# -*- coding: utf-8 -*-
"""Emit slide XML in an existing deck's own idiom, and do the package surgery.

The deck we build from is the design system. Rather than re-deriving a look
with a generator library, we write the same shapes the template's author drew
-- same fonts, same fills, same EMU grid -- so a new slide sitting next to an
original one is indistinguishable.

Defaults below are the VRAIN Solution proposal template. For another template,
run inspect_template.py and override the module-level constants; see
references/generic-template.md.
"""
import io
import os
import re

# ---------------------------------------------------------------- tokens
FONT = "メイリオ"
NAVY = "112C4B"      # headers, emphasis, primary blocks
RED = "7F201B"       # problems, deadlines, anything the reader must not miss
BLUE = "2E5A8C"      # secondary blocks
LTBLUE = "E9EEF4"    # conclusion banners, "after" states
LTGRAY = "F2F2F2"    # zebra rows, "before" states
GRAY = "595959"      # body copy
BORDER = "BFBFBF"
MIDGRAY = "8C8C8C"   # inert steps in a flow
WHITE = "FFFFFF"

# A4 landscape, 9906000 x 6858000 EMU.
L = 484632            # content left edge
W = 8961120           # content width (right edge = 9445752)
TOP = 1645920         # content top edge, below the lead-message placeholder
FOOT = 6100000        # keep content above this; the footer sits at 6627812

HDR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
       '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
       ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
       ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
       '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
       '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
       '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>')
FTR = '</p:spTree></p:cSld><a:clrMapOvr><a:masterClrMapping/></a:clrMapOvr></p:sld>'

# effectRef idx="0": the template's own shapes say idx="2" plus an empty
# <a:effectLst/> override, which PowerPoint honours but LibreOffice does not --
# under a theme whose effect #2 is a shadow, every card grows one. idx 0 = none.
STYLE = ('<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
         '<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>'
         '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
         '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>')


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def ids(start=4):
    """Shape ids. 1 is the group, 2/3 are the two placeholders."""
    n = start
    while True:
        yield n
        n += 1


# ---------------------------------------------------------------- runs & paragraphs
def rpr(sz, b, color, i=0, u=0):
    return ('<a:rPr lang="ja-JP" altLang="en-US" sz="%d" b="%d"%s%s dirty="0">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/></a:rPr>'
            % (sz, b, ' i="1"' if i else '', ' u="sng"' if u else '', color, FONT, FONT, FONT))


def R(t, sz=1100, b=0, color=GRAY, i=0, u=0):
    """One run. sz is in hundredths of a point: 1100 = 11pt."""
    return '<a:r>%s<a:t>%s</a:t></a:r>' % (rpr(sz, b, color, i, u), esc(t))


def P(runs, algn="l", spc_aft=None, spc_bef=None, lnspc=None):
    """One paragraph. Mixing colours/weights in a line = several runs in one P."""
    if isinstance(runs, (list, tuple)):
        runs = "".join(runs)
    props = ""
    if lnspc:
        props += '<a:lnSpc><a:spcPct val="%d"/></a:lnSpc>' % lnspc
    if spc_bef:
        props += '<a:spcBef><a:spcPts val="%d"/></a:spcBef>' % spc_bef
    if spc_aft:
        props += '<a:spcAft><a:spcPts val="%d"/></a:spcAft>' % spc_aft
    return '<a:p><a:pPr algn="%s">%s</a:pPr>%s</a:p>' % (algn, props, runs or '<a:endParaRPr/>')


def _xfrm(x, y, cx, cy):
    return '<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>' % (
        int(x), int(y), int(cx), int(cy))


# ---------------------------------------------------------------- primitives
def shape(sid, name, x, y, cx, cy, fill=None, line=None, prst="rect",
          paras=None, anchor="ctr", adj=None, lw=12700):
    f = '<a:noFill/>' if fill is None else '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill
    ln = ('<a:ln><a:noFill/></a:ln>' if line is None else
          '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>' % (lw, line))
    body = "".join(paras) if paras else '<a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p>'
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>%s<a:prstGeom prst="%s">%s</a:prstGeom>%s%s<a:effectLst/></p:spPr>%s'
            '<p:txBody><a:bodyPr wrap="square" lIns="54864" tIns="27432" rIns="54864" bIns="27432"'
            ' rtlCol="0" anchor="%s"/><a:lstStyle/>%s</p:txBody></p:sp>'
            % (sid, name, _xfrm(x, y, cx, cy), prst, adj or '<a:avLst/>', f, ln, STYLE, anchor, body))


def textbox(sid, name, x, y, cx, cy, paras, anchor="t"):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288"'
            ' anchor="%s"/><a:lstStyle/>%s</p:txBody></p:sp>'
            % (sid, name, _xfrm(x, y, cx, cy), anchor, "".join(paras)))


def placeholders(title, message, title_sz=2000, msg_sz=1150, idx_title=12, idx_msg=13):
    """The layout's own title + lead-message slots. Leaving these as placeholders
    (rather than drawing text boxes) is what keeps the header rule, logo and
    CONFIDENTIAL badge inherited from the layout."""
    t = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Text Placeholder 1"/>'
         '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
         '<p:nvPr><p:ph type="body" sz="quarter" idx="%d"/></p:nvPr></p:nvSpPr><p:spPr/>'
         '<p:txBody><a:bodyPr/><a:lstStyle/><a:p>%s</a:p></p:txBody></p:sp>'
         % (idx_title, R(title, sz=title_sz, b=1, color=NAVY)))
    m = ('<p:sp><p:nvSpPr><p:cNvPr id="3" name="Text Placeholder 2"/>'
         '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
         '<p:nvPr><p:ph type="body" sz="quarter" idx="%d"/></p:nvPr></p:nvSpPr><p:spPr/>'
         '<p:txBody><a:bodyPr/><a:lstStyle/><a:p>%s</a:p></p:txBody></p:sp>'
         % (idx_msg, R(message, sz=msg_sz, b=0, color=GRAY)))
    return t + m


# ---------------------------------------------------------------- composites
def card(ids_, x, y, cx, cy, header, lines, header_fill=NAVY, hdr_h=420624,
         body_sz=1050, spc_aft=800, pad_x=155448, top_gap=548640):
    """White bordered panel, coloured header bar, bullet body.
    `lines` items are plain strings, or (text, bold, colour) to emphasise one."""
    out = shape(next(ids_), "Card", x, y, cx, cy, fill=WHITE, line=BORDER)
    out += shape(next(ids_), "CardHdr", x, y, cx, hdr_h, fill=header_fill)
    out += textbox(next(ids_), "CardHdrTx", x, y, cx, hdr_h,
                   [P(R(header, sz=1300, b=1, color=WHITE), algn="ctr")], anchor="ctr")
    paras = []
    for ln in lines:
        txt, b, col = (ln, 0, GRAY) if isinstance(ln, str) else ln
        paras.append(P(R(txt, sz=body_sz, b=b, color=col), spc_aft=spc_aft))
    out += textbox(next(ids_), "CardBody", x + pad_x, y + top_gap,
                   cx - 2 * pad_x, cy - top_gap - 91440, paras)
    return out


def banner(ids_, x, y, cx, cy, runs, fill=LTBLUE, line=NAVY, algn="ctr"):
    """The '⇒ so what' strip that closes a slide."""
    return shape(next(ids_), "Banner", x, y, cx, cy, fill=fill, line=line,
                 paras=[P(runs, algn=algn)], anchor="ctr")


def point_row(ids_, x, y, cx, cy, label, body_lines, fill=LTGRAY,
              label_w=2606040, label_x_pad=201168):
    """Tinted band + navy edge tab + bold label + body. Stack 3 of these for a
    '<本構成のポイント>' block."""
    out = shape(next(ids_), "Row", x, y, cx, cy, fill=fill)
    out += shape(next(ids_), "RowTab", x, y, 91440, cy, fill=NAVY)
    out += textbox(next(ids_), "RowLabel", x + label_x_pad, y, label_w, cy,
                   [P(R(label, sz=1150, b=1, color=NAVY))], anchor="ctr")
    bx = x + label_x_pad + label_w + 219456
    out += textbox(next(ids_), "RowBody", bx, y, x + cx - bx - 91440, cy,
                   [P(R(t, sz=1050, color=GRAY)) for t in body_lines], anchor="ctr")
    return out


def flow(ids_, labels, y=1783080, bw=1627632, bh=960120, x0=548640, pitch=1911096,
         arrow_y=None, arrow_w=228600, arrow_h=182880):
    """Left-to-right process chain. `labels` are (text, fill, textcolour);
    newlines split the text into lines. 5 boxes is the comfortable maximum."""
    out = ""
    if arrow_y is None:
        arrow_y = y + (bh - arrow_h) // 2
    for k, (label, fill, col) in enumerate(labels):
        x = x0 + k * pitch
        out += shape(next(ids_), "Box%d" % k, x, y, bw, bh, fill=fill, prst="roundRect",
                     paras=[P(R(t, sz=1050, b=1, color=col), algn="ctr")
                            for t in label.split("\n")], anchor="ctr")
        if k < len(labels) - 1:
            out += shape(next(ids_), "Arrow%d" % k, x + bw + 27432, arrow_y,
                         arrow_w, arrow_h, fill=BORDER, prst="rightArrow")
    return out


def gantt(ids_, rows, months, year_spans, x=L, y=1737360, width=W,
          label_w=2011680, hdr_h=320040, row_h=480000):
    """Schedule chart. `months` are the column labels (ints or strings),
    `year_spans` are (start_index, length, label), and `rows` are
    (row_label, start_index, length, bar_colour, bar_text)."""
    track_x = x + label_w
    mw = (width - label_w) / float(len(months))
    out = shape(next(ids_), "CornerY", x, y, label_w, hdr_h, fill=NAVY, line=WHITE)
    out += shape(next(ids_), "CornerM", x, y + hdr_h, label_w, hdr_h, fill=NAVY, line=WHITE,
                 paras=[P(R("工程", sz=1000, b=1, color=WHITE), algn="ctr")])
    for s, ln, label in year_spans:
        out += shape(next(ids_), "Year", track_x + s * mw, y, ln * mw, hdr_h,
                     fill=NAVY, line=WHITE,
                     paras=[P(R(label, sz=1000, b=1, color=WHITE), algn="ctr")])
    for k, m in enumerate(months):
        out += shape(next(ids_), "Month", track_x + k * mw, y + hdr_h, mw, hdr_h,
                     fill=LTGRAY, line=BORDER,
                     paras=[P(R(str(m), sz=850, color=GRAY), algn="ctr")])
    ry = y + 2 * hdr_h
    for k, (label, s, ln, col, bar_text) in enumerate(rows):
        fill = LTGRAY if k % 2 else WHITE
        out += shape(next(ids_), "PhLabel", x, ry, label_w, row_h, fill=fill, line=BORDER,
                     paras=[P(R(label, sz=1000, b=1, color=NAVY))], anchor="ctr")
        out += shape(next(ids_), "Track", track_x, ry, width - label_w, row_h,
                     fill=fill, line=BORDER)
        out += shape(next(ids_), "Bar", track_x + s * mw + 18288,
                     ry + (row_h - hdr_h) // 2, ln * mw - 36576, hdr_h,
                     fill=col, prst="roundRect",
                     paras=[P(R(bar_text, sz=850, b=1, color=WHITE), algn="ctr")], anchor="ctr")
        ry += row_h
    return out


# ---------------------------------------------------------------- tables
def cell(paras, fill=WHITE, anchor="ctr"):
    return ('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>%s</a:txBody>'
            '<a:tcPr marL="54864" marR="54864" marT="27432" marB="27432" anchor="%s">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:tcPr></a:tc>'
            % ("".join(paras), anchor, fill))


def tcell(text, sz=1000, b=0, color=GRAY, fill=WHITE, algn="l", anchor="ctr"):
    """Cell from text; a newline starts a new paragraph (never concatenate
    list items into one paragraph -- spacing collapses)."""
    return cell([P(R(t, sz=sz, b=b, color=color), algn=algn) for t in text.split("\n")],
                fill=fill, anchor=anchor)


def header_row(titles, fill=NAVY, sz=1000, algns=None):
    algns = algns or ["ctr"] * len(titles)
    return [tcell(t, sz=sz, b=1, color=WHITE, fill=fill, algn=a)
            for t, a in zip(titles, algns)]


def table(sid, x, y, widths, rows, row_heights):
    """Row heights are minimums -- PowerPoint grows a row to fit its text, so
    size the column widths and let the heights settle."""
    grid = "".join('<a:gridCol w="%d"/>' % w for w in widths)
    trs = "".join('<a:tr h="%d">%s</a:tr>' % (h, "".join(c))
                  for h, c in zip(row_heights, rows))
    return ('<p:graphicFrame><p:nvGraphicFramePr>'
            '<p:cNvPr id="%d" name="Table"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
            '<p:nvPr/></p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            '<a:tbl><a:tblPr/><a:tblGrid>%s</a:tblGrid>%s</a:tbl></a:graphicData></a:graphic>'
            '</p:graphicFrame>'
            % (sid, int(x), int(y), sum(widths), sum(row_heights), grid, trs))


def zebra(rows_src, cell_specs, start_fill=WHITE, alt_fill=LTGRAY):
    """Build body rows with alternating fills. cell_specs is a list of dicts of
    tcell kwargs (minus text/fill), one per column."""
    out = []
    for k, values in enumerate(rows_src):
        fill = alt_fill if k % 2 else start_fill
        out.append([tcell(v, fill=fill, **spec) for v, spec in zip(values, cell_specs)])
    return out


# ---------------------------------------------------------------- geometry helpers
SLIDE_W, SLIDE_H = 9906000, 6858000


def cols(n, gap=219456, x=None, width=None):
    """[(x, w), ...] for n equal columns across the content width. Use this
    instead of hardcoding A4 positions like 5074920, so layouts survive a
    template with a different slide size."""
    x = L if x is None else x
    width = W if width is None else width
    w = (width - gap * (n - 1)) // n
    return [(x + k * (w + gap), w) for k in range(n)]


def configure_for(unpacked, margin=None):
    """Read the deck's slide size and re-derive L/W/FOOT from it. Call once,
    right after unpacking a non-VRAIN template (VRAIN values are the defaults).
    Composite defaults bound at import time are not touched -- pass explicit
    coordinates computed from L/W/cols()."""
    global SLIDE_W, SLIDE_H, L, W, FOOT
    pres = io.open(os.path.join(unpacked, "ppt", "presentation.xml"), encoding="utf-8").read()
    m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
    SLIDE_W, SLIDE_H = int(m.group(1)), int(m.group(2))
    L = int(SLIDE_W * 484632 / 9906000) if margin is None else margin
    W = SLIDE_W - 2 * L
    FOOT = int(SLIDE_H * 6100000 / 6858000)
    return SLIDE_W, SLIDE_H


def check_bounds(unpacked, slides=None, tol=0):
    """Report shapes that run off the slide. Coordinates past the edge are
    written, not clamped -- the shape just silently isn't fully visible.
    Returns a list of (slide, shape_name, right, bottom); empty = clean."""
    pres = io.open(os.path.join(unpacked, "ppt", "presentation.xml"), encoding="utf-8").read()
    m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
    sw, sh = int(m.group(1)), int(m.group(2))
    sd = os.path.join(unpacked, "ppt", "slides")
    names = slides or sorted(f[:-4] for f in os.listdir(sd) if re.fullmatch(r"slide\d+\.xml", f))
    bad = []
    for s in names:
        xml = io.open(os.path.join(sd, s + ".xml"), encoding="utf-8").read()
        for sp in re.finditer(r'<p:(sp|graphicFrame|pic)>.*?</p:\1>', xml, re.S):
            blk = sp.group(0)
            xf = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>', blk)
            if not xf:
                continue
            x0, y0, cx, cy = map(int, xf.groups())
            if x0 < -tol or y0 < -tol or x0 + cx > sw + tol or y0 + cy > sh + tol:
                nm = re.search(r'name="([^"]*)"', blk)
                bad.append((s, nm.group(1) if nm else "?", x0 + cx, y0 + cy))
    for s, nm, r, b in bad:
        print("  OFF-SLIDE %s %-14s right=%d (max %d) bottom=%d (max %d)" % (s, nm, r, sw, b, sh))
    return bad


# ---------------------------------------------------------------- writing & packaging
def build(path, title, message, body, with_placeholders=True, **ph):
    """Write a slide. Set with_placeholders=False for a template whose layout has
    no title/message body placeholders -- draw those with textbox() instead."""
    head = placeholders(title, message, **ph) if with_placeholders else ""
    xml = HDR + head + body + FTR
    with io.open(path, "w", encoding="utf-8") as f:
        f.write(xml)
    return len(xml)


def set_para(xml, old_text, new_text):
    """Replace a paragraph whose combined run text equals `old_text`, collapsing
    it to a single run that keeps the first run's formatting. Use this on the
    cover/divider slides, where a phrase is often split across several runs
    (e.g. '〇〇株式' + '会社　御中') and a plain string replace would miss it."""
    def repl(m):
        s = m.group(0)
        if "".join(re.findall(r'<a:t>([^<]*)</a:t>', s)) != old_text:
            return s
        # Keep the first run's formatting. Runs may legitimately carry no
        # <a:rPr> at all (python-pptx writes them that way) -- inherit then.
        rpr_m = re.search(r'<a:r>(<a:rPr.*?</a:rPr>|<a:rPr[^>]*/>)?', s, re.S)
        rpr_xml = (rpr_m.group(1) or '') if rpr_m else ''
        ppr = re.match(r'<a:p>(<a:pPr.*?</a:pPr>|<a:pPr[^>]*/>)?', s, re.S).group(1) or ''
        return '<a:p>%s<a:r>%s<a:t>%s</a:t></a:r></a:p>' % (ppr, rpr_xml, esc(new_text))
    out = re.sub(r'<a:p>.*?</a:p>', repl, xml, flags=re.S)
    if out == xml:
        raise ValueError("no paragraph matched: %r" % old_text)
    return out


def paragraph_texts(xml):
    """Non-empty paragraphs' combined text, in document order. Use it to find
    what a template's cover actually says before replacing it -- the strings
    differ for every past deal, so match what is there, never what you expect."""
    out = []
    for p in re.findall(r'<a:p>.*?</a:p>|<a:p .*?</a:p>', xml, re.S):
        t = "".join(re.findall(r'<a:t>([^<]*)</a:t>', p))
        if t.strip():
            out.append(t)
    return out


def run_texts(xml):
    """Every <a:t> in document order (slide-number field text included)."""
    return re.findall(r'<a:t>([^<]*)</a:t>', xml)


def swap_runs(xml, pairs):
    """Replace whole run texts, applied simultaneously so a rotation like
    A->B, B->C does not cascade."""
    tok = {a: "\x00%d\x00" % k for k, (a, _) in enumerate(pairs)}
    for a, _ in pairs:
        xml = xml.replace('<a:t>%s</a:t>' % esc(a), '<a:t>%s</a:t>' % tok[a])
    for (a, b) in pairs:
        xml = xml.replace('<a:t>%s</a:t>' % tok[a], '<a:t>%s</a:t>' % esc(b))
    return xml


def edit_slide(unpacked, name, fn):
    """Apply fn(xml) -> xml to ppt/slides/<name>.xml, writing only if fn succeeds.

    Use this rather than open(path, "w").write(transform(...)): opening for
    write truncates the file first, so a transform that raises leaves a
    0-byte slide behind and a deck PowerPoint refuses to open.
        d.edit_slide("unpacked", "slide1", lambda x: d.set_para(x, "旧", "新"))
    """
    path = os.path.join(unpacked, "ppt", "slides", name + ".xml")
    xml = io.open(path, encoding="utf-8").read()
    new = fn(xml)
    if not new or "<p:sld" not in new:
        raise ValueError("edit_slide(%s): transform returned no slide XML" % name)
    io.open(path, "w", encoding="utf-8").write(new)
    return new


def reorder_slides(unpacked, order):
    """Rewrite <p:sldIdLst> to exactly `order` (names like 'slide7', no .xml).
    Slides left out are dropped from the deck; run the pptx skill's clean.py
    afterwards to collect the media they were holding."""
    rels_p = os.path.join(unpacked, "ppt", "_rels", "presentation.xml.rels")
    rels = io.open(rels_p, encoding="utf-8").read()
    rid = {t: i for i, t in re.findall(
        r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"', rels)}
    missing = [s for s in order if s + ".xml" not in rid]
    if missing:
        raise ValueError("not in presentation.xml.rels: %s" % missing)
    pres_p = os.path.join(unpacked, "ppt", "presentation.xml")
    pres = io.open(pres_p, encoding="utf-8").read()
    lst = "".join('<p:sldId id="%d" r:id="%s"/>' % (256 + k, rid[s + ".xml"])
                  for k, s in enumerate(order))
    out = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>',
                 '<p:sldIdLst>%s</p:sldIdLst>' % lst, pres, flags=re.S)
    if out == pres:
        raise ValueError("no <p:sldIdLst> in presentation.xml")
    io.open(pres_p, "w", encoding="utf-8").write(out)
    return len(order)


def unpack(pptx_path, dest="unpacked"):
    import zipfile
    zipfile.ZipFile(pptx_path).extractall(dest)
    return dest


def pack(unpacked, out_pptx):
    """Zip from inside the directory, and remove any previous output first --
    zip updates in place, so parts you deleted would otherwise survive."""
    import subprocess
    out_pptx = os.path.abspath(out_pptx)
    if os.path.exists(out_pptx):
        os.remove(out_pptx)
    subprocess.run(["zip", "-qXr", out_pptx, "."], cwd=unpacked, check=True)
    return out_pptx
