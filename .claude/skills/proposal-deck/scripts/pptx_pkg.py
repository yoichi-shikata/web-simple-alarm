# -*- coding: utf-8 -*-
"""Self-contained OOXML package operations for .pptx decks.

This replaces the helpers this skill used to borrow from Claude Code's bundled
`pptx` skill, so the exact same code runs under Claude Code, Codex or any other
agent -- identical tooling is what makes identical output possible.

Only the Python standard library is required here. XML is parsed for reading
and validation only; slide XML is always edited as text, never re-serialised
through ElementTree, because that rewrites namespace prefixes and PowerPoint
then refuses the file.
"""
import glob
import io
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.dom.minidom

SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"


def _read(p):
    return io.open(p, encoding="utf-8").read()


def _write(p, s):
    io.open(p, "w", encoding="utf-8").write(s)


def _sld_ids(pres_xml):
    return [(int(i), r) for i, r in re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"', pres_xml)]


def _rel_map(rels_xml):
    """rId -> Target for every Relationship."""
    out = {}
    for m in re.finditer(r'<Relationship\b[^>]*/>', rels_xml):
        tag = m.group(0)
        rid = re.search(r'\bId="([^"]+)"', tag)
        tgt = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and tgt:
            out[rid.group(1)] = tgt.group(1)
    return out


def slide_order(unpacked):
    """Slide file names ('slide7') in presentation order."""
    pres = _read(os.path.join(unpacked, "ppt", "presentation.xml"))
    rels = _rel_map(_read(os.path.join(unpacked, "ppt", "_rels", "presentation.xml.rels")))
    return [os.path.basename(rels[r])[:-4] for _, r in _sld_ids(pres) if r in rels]


# ---------------------------------------------------------------- duplicate
def add_slide(unpacked, src, after=None):
    """Duplicate ppt/slides/<src> (e.g. 'slide4.xml') and register it everywhere
    a slide must be registered. The copy's notes link is dropped so two slides
    never share one notes page. Returns the new file name."""
    sd = os.path.join(unpacked, "ppt", "slides")
    src = src if src.endswith(".xml") else src + ".xml"
    if not os.path.exists(os.path.join(sd, src)):
        raise FileNotFoundError(os.path.join(sd, src))
    nums = [int(m) for m in (re.findall(r'^slide(\d+)\.xml$', f) for f in os.listdir(sd)) for m in m]
    new = "slide%d.xml" % (max(nums) + 1)

    shutil.copyfile(os.path.join(sd, src), os.path.join(sd, new))
    src_rels = os.path.join(sd, "_rels", src + ".rels")
    if os.path.exists(src_rels):
        r = _read(src_rels)
        r = re.sub(r'<Relationship\b[^>]*relationships/notesSlide"[^>]*/>', "", r)
        _write(os.path.join(sd, "_rels", new + ".rels"), r)

    ct_p = os.path.join(unpacked, "[Content_Types].xml")
    ct = _read(ct_p)
    ct = ct.replace("</Types>", '<Override PartName="/ppt/slides/%s" ContentType="%s"/></Types>' % (new, SLIDE_CT))
    _write(ct_p, ct)

    prels_p = os.path.join(unpacked, "ppt", "_rels", "presentation.xml.rels")
    prels = _read(prels_p)
    rid = "rId%d" % (max(int(x) for x in re.findall(r'Id="rId(\d+)"', prels)) + 1)
    prels = prels.replace("</Relationships>",
                          '<Relationship Id="%s" Type="%s" Target="slides/%s"/></Relationships>' % (rid, SLIDE_REL, new))
    _write(prels_p, prels)

    pres_p = os.path.join(unpacked, "ppt", "presentation.xml")
    pres = _read(pres_p)
    ids = _sld_ids(pres)
    entry = '<p:sldId id="%d" r:id="%s"/>' % (max([i for i, _ in ids] + [255]) + 1, rid)
    if after:
        after = after if after.endswith(".xml") else after + ".xml"
        rmap = _rel_map(prels)
        arid = next((k for k, v in rmap.items() if v == "slides/" + after), None)
        m = re.search(r'<p:sldId id="\d+" r:id="%s"/>' % arid, pres) if arid else None
        if not m:
            raise ValueError("--after slide not in deck: %s" % after)
        pres = pres[:m.end()] + entry + pres[m.end():]
    else:
        pres = pres.replace("</p:sldIdLst>", entry + "</p:sldIdLst>")
    _write(pres_p, pres)
    return new


# ---------------------------------------------------------------- clean
_ORPHAN_DIRS = ("slides", "notesSlides", "media", "charts", "embeddings",
                "diagrams", "tags", "ink", "drawings")


def _resolve(base_dir, target):
    return os.path.normpath(os.path.join(base_dir, target)).replace("\\", "/")


def clean(unpacked):
    """Drop slides missing from <p:sldIdLst>, then repeatedly drop any part no
    remaining relationship points at (the media a deleted slide was holding,
    its notes page, charts...). Run only after the slide list is final."""
    unpacked = os.path.abspath(unpacked)
    removed = []
    keep = set(slide_order(unpacked))
    prels_p = os.path.join(unpacked, "ppt", "_rels", "presentation.xml.rels")
    prels = _read(prels_p)
    for f in os.listdir(os.path.join(unpacked, "ppt", "slides")):
        m = re.match(r'^(slide\d+)\.xml$', f)
        if m and m.group(1) not in keep:
            prels = re.sub(r'<Relationship\b[^>]*Target="slides/%s"[^>]*/>' % re.escape(f), "", prels)
    _write(prels_p, prels)

    while True:
        referenced = set()
        for rels in glob.glob(os.path.join(unpacked, "**", "_rels", "*.rels"), recursive=True):
            base = os.path.dirname(os.path.dirname(rels))
            for m in re.finditer(r'<Relationship\b[^>]*/>', _read(rels)):
                tag = m.group(0)
                if 'TargetMode="External"' in tag:
                    continue
                t = re.search(r'\bTarget="([^"]+)"', tag)
                if t:
                    tgt = t.group(1)
                    full = os.path.join(unpacked, tgt.lstrip("/")) if tgt.startswith("/") \
                        else _resolve(base, tgt)
                    referenced.add(os.path.normpath(full))
        victims = []
        for d in _ORPHAN_DIRS:
            for f in glob.glob(os.path.join(unpacked, "ppt", d, "*")):
                if os.path.isfile(f) and os.path.normpath(f) not in referenced:
                    victims.append(f)
        if not victims:
            break
        for f in victims:
            os.remove(f)
            rels = os.path.join(os.path.dirname(f), "_rels", os.path.basename(f) + ".rels")
            if os.path.exists(rels):
                os.remove(rels)
            removed.append(os.path.relpath(f, unpacked).replace("\\", "/"))

    ct_p = os.path.join(unpacked, "[Content_Types].xml")
    ct = _read(ct_p)
    for part in re.findall(r'<Override PartName="([^"]+)"', ct):
        if not os.path.exists(os.path.join(unpacked, part.lstrip("/"))):
            ct = re.sub(r'<Override PartName="%s"[^>]*/>' % re.escape(part), "", ct)
    _write(ct_p, ct)
    return removed


# ---------------------------------------------------------------- validate
def validate(pptx):
    """Structural checks that catch what makes PowerPoint refuse or 'repair' a
    deck: unparsable or empty parts, dangling relationships, parts without a
    content type, broken slide list, duplicate shape ids. Returns a list of
    problems; empty means pass."""
    errs = []
    try:
        z = zipfile.ZipFile(pptx)
    except zipfile.BadZipFile as e:
        return ["not a zip: %s" % e]
    names = set(z.namelist())
    for n in names:
        if n.endswith((".xml", ".rels")):
            data = z.read(n)
            if not data.strip():
                errs.append("%s: empty part" % n)
                continue
            try:
                xml.dom.minidom.parseString(data)
            except Exception as e:  # noqa: BLE001 - report any parser failure
                errs.append("%s: XML parse error: %s" % (n, e))
    for n in names:
        if not n.endswith(".rels"):
            continue
        base = os.path.dirname(os.path.dirname(n))
        for m in re.finditer(r'<Relationship\b[^>]*/>', z.read(n).decode("utf-8", "replace")):
            tag = m.group(0)
            if 'TargetMode="External"' in tag:
                continue
            t = re.search(r'\bTarget="([^"]+)"', tag).group(1)
            full = t.lstrip("/") if t.startswith("/") else os.path.normpath(os.path.join(base, t)).replace("\\", "/")
            if full not in names:
                errs.append("%s: target missing: %s" % (n, t))
    ct = z.read("[Content_Types].xml").decode("utf-8")
    defaults = set(e.lower() for e in re.findall(r'<Default Extension="([^"]+)"', ct))
    overrides = set(p.lstrip("/") for p in re.findall(r'<Override PartName="([^"]+)"', ct))
    for n in names:
        if n.endswith("/") or n == "[Content_Types].xml":
            continue
        if n not in overrides and n.rsplit(".", 1)[-1].lower() not in defaults:
            errs.append("%s: no content type" % n)
    for p in overrides:
        if p not in names:
            errs.append("[Content_Types].xml: override for missing part /%s" % p)
    pres = z.read("ppt/presentation.xml").decode("utf-8")
    prels = _rel_map(z.read("ppt/_rels/presentation.xml.rels").decode("utf-8"))
    seen = set()
    for i, r in _sld_ids(pres):
        if i in seen:
            errs.append("presentation.xml: duplicate sldId %d" % i)
        seen.add(i)
        if r not in prels:
            errs.append("presentation.xml: sldId r:id %s has no relationship" % r)
            continue
        part = "ppt/" + prels[r]
        if part not in names:
            errs.append("presentation.xml: slide part missing: %s" % part)
            continue
        sx = z.read(part).decode("utf-8", "replace")
        if "<p:sld" not in sx:
            errs.append("%s: not a slide" % part)
        sids = re.findall(r'<p:cNvPr id="(\d+)"', sx)
        dup = sorted(set(s for s in sids if sids.count(s) > 1), key=int)
        if dup:
            errs.append("%s: duplicate shape ids %s (PowerPoint will 'repair' the file)" % (part, ",".join(dup)))
    return errs


# ---------------------------------------------------------------- text dump
def text(pptx):
    """Plain-text dump in presentation order, notes included -- for proofreading
    and for grepping out leftovers from the previous deal."""
    z = zipfile.ZipFile(pptx)
    pres = z.read("ppt/presentation.xml").decode("utf-8")
    prels = _rel_map(z.read("ppt/_rels/presentation.xml.rels").decode("utf-8"))
    out = []
    for k, (_, r) in enumerate(_sld_ids(pres), 1):
        part = "ppt/" + prels[r]
        sx = z.read(part).decode("utf-8")
        out.append("<!-- Slide %d: %s -->" % (k, os.path.basename(part)))
        for p in re.findall(r'<a:p>.*?</a:p>|<a:p .*?</a:p>', sx, re.S):
            t = "".join(re.findall(r'<a:t>([^<]*)</a:t>', p))
            if t.strip():
                out.append(t)
        rels_part = "ppt/slides/_rels/%s.rels" % os.path.basename(part)
        if rels_part in z.namelist():
            for tgt in _rel_map(z.read(rels_part).decode("utf-8")).values():
                if "notesSlide" in tgt:
                    npart = os.path.normpath(os.path.join("ppt/slides", tgt)).replace("\\", "/")
                    if npart in z.namelist():
                        nt = "".join(re.findall(r'<a:t>([^<]*)</a:t>', z.read(npart).decode("utf-8")))
                        if nt.strip():
                            out.append("[notes] " + nt)
        out.append("")
    return "\n".join(out)


# ---------------------------------------------------------------- rendering
def soffice_bin():
    """Locate LibreOffice on Linux, macOS or Windows."""
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return p
    return None


def to_pdf(pptx, outdir=None, timeout=300):
    """Convert with a throwaway user profile: a shared/default profile is the
    usual reason headless soffice hangs or silently converts nothing."""
    sb = soffice_bin()
    if not sb:
        raise RuntimeError("LibreOffice not found -- run scripts/setup_env.py")
    pptx = os.path.abspath(pptx)
    outdir = os.path.abspath(outdir or os.path.dirname(pptx))
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    if os.path.exists(pdf):
        os.remove(pdf)
    prof = tempfile.mkdtemp(prefix="lo_profile_")
    env = dict(os.environ)
    if platform.system() == "Linux":
        env.setdefault("SAL_USE_VCLPLUGIN", "svp")
    try:
        uri = "file:///" + prof.replace("\\", "/").lstrip("/")
        subprocess.run([sb, "-env:UserInstallation=" + uri, "--headless", "--convert-to", "pdf",
                        "--outdir", outdir, pptx], env=env, timeout=timeout,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    finally:
        shutil.rmtree(prof, ignore_errors=True)
    if not os.path.exists(pdf) or os.path.getsize(pdf) == 0:
        raise RuntimeError("PDF conversion failed for %s. If LibreOffice is installed but this "
                           "still fails on Linux, the Impress component is probably missing "
                           "(apt package libreoffice-impress)." % pptx)
    return pdf


def pdf_to_images(pdf, outdir, dpi=130):
    """Rasterise with PyMuPDF (pip-only, works on Windows) or fall back to
    poppler's pdftoppm. Returns image paths in page order."""
    os.makedirs(outdir, exist_ok=True)
    for f in glob.glob(os.path.join(outdir, "p-*.jpg")):
        os.remove(f)
    try:
        try:
            import pymupdf as fitz  # PyMuPDF >= 1.24
        except ImportError:
            import fitz  # older PyMuPDF
        doc = fitz.open(pdf)
        width = len(str(doc.page_count))
        paths = []
        for i, page in enumerate(doc, 1):
            p = os.path.join(outdir, "p-%0*d.jpg" % (width, i))
            page.get_pixmap(dpi=dpi).save(p)
            paths.append(p)
        return paths
    except ImportError:
        pass
    if not shutil.which("pdftoppm"):
        raise RuntimeError("neither PyMuPDF nor pdftoppm available -- run scripts/setup_env.py")
    subprocess.run(["pdftoppm", "-jpeg", "-r", str(dpi), pdf, os.path.join(outdir, "p")], check=True)
    return sorted(glob.glob(os.path.join(outdir, "p-*.jpg")))


def render(pptx, outdir="qa-img", dpi=130):
    pdf = to_pdf(pptx)
    return pdf_to_images(pdf, outdir, dpi)


def thumbnail(pptx, prefix="thumbs", cols=3, width=420):
    """Labelled contact sheet(s), 12 slides per sheet, for choosing layouts."""
    from PIL import Image, ImageDraw
    tmp = tempfile.mkdtemp(prefix="thumbs_")
    try:
        imgs = pdf_to_images(to_pdf(pptx, tmp), tmp, dpi=60)
        z = zipfile.ZipFile(pptx)
        pres = z.read("ppt/presentation.xml").decode("utf-8")
        prels = _rel_map(z.read("ppt/_rels/presentation.xml.rels").decode("utf-8"))
        labels = [os.path.basename(prels[r]) for _, r in _sld_ids(pres)]
        out = []
        for s in range(0, len(imgs), 12):
            chunk = imgs[s:s + 12]
            thumbs = []
            for p in chunk:
                im = Image.open(p)
                h = int(im.height * width / im.width)
                thumbs.append(im.resize((width, h)))
            th = thumbs[0].height
            rows = (len(thumbs) + cols - 1) // cols
            sheet = Image.new("RGB", (cols * (width + 20) + 20, rows * (th + 50) + 20), "white")
            d = ImageDraw.Draw(sheet)
            for k, im in enumerate(thumbs):
                x = 20 + (k % cols) * (width + 20)
                y = 20 + (k // cols) * (th + 50)
                d.text((x, y), "%d  %s" % (s + k + 1, labels[s + k] if s + k < len(labels) else ""), fill="black")
                sheet.paste(im, (x, y + 18))
                d.rectangle([x, y + 18, x + width - 1, y + 18 + th - 1], outline="gray")
            name = "%s-%d.jpg" % (prefix, s // 12 + 1)
            sheet.save(name, quality=85)
            out.append(name)
        return out
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
