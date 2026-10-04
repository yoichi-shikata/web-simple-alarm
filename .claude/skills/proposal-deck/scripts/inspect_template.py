#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dump a deck's design system: layout usage, shape geometry, fills, fonts, text.

Reading a template by eye through unzipped XML is slow and easy to get wrong.
This prints the few things that actually decide how to rebuild a slide --
which layout each slide uses, where its shapes sit in EMU, what colours and
fonts it uses, and which slides carry media (those must be reused verbatim,
never regenerated).

Usage:
  inspect_template.py deck.pptx                 # per-slide summary
  inspect_template.py deck.pptx --slides 4,6,9  # full geometry for those slides
  inspect_template.py deck.pptx --layouts       # layout names + placeholder idx
  inspect_template.py deck.pptx --tokens        # colour/font frequency
"""
import argparse
import collections
import re
import sys
import zipfile

SHAPE_RE = re.compile(r'<p:sp>.*?</p:sp>|<p:graphicFrame>.*?</p:graphicFrame>|<p:pic>.*?</p:pic>', re.S)
XFRM_RE = re.compile(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>')


def slide_nums(z):
    ns = [int(re.search(r'slide(\d+)\.xml', n).group(1))
          for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)]
    return sorted(ns)


def read(z, name):
    try:
        return z.read(name).decode("utf-8")
    except KeyError:
        return ""


def shape_rows(xml):
    for m in SHAPE_RE.finditer(xml):
        s = m.group(0)
        nm = re.search(r'name="([^"]*)"', s)
        xf = XFRM_RE.search(s)
        prst = re.search(r'<a:prstGeom prst="([^"]+)"', s)
        fill = re.search(r'<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/></a:solidFill>', s)
        ph = re.search(r'<p:ph[^>]*idx="(\d+)"', s)
        txt = " / ".join(re.findall(r'<a:t>([^<]*)</a:t>', s))
        yield {"name": nm.group(1) if nm else "?",
               "prst": prst.group(1) if prst else ("table" if "<a:tbl>" in s else "-"),
               "xfrm": tuple(int(v) for v in xf.groups()) if xf else None,
               "fill": fill.group(1).upper() if fill else None,
               "ph": ph.group(1) if ph else None,
               "text": txt}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--slides", help="comma-separated slide numbers for full geometry")
    ap.add_argument("--layouts", action="store_true")
    ap.add_argument("--tokens", action="store_true")
    a = ap.parse_args()

    z = zipfile.ZipFile(a.pptx)
    sz = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', read(z, "ppt/presentation.xml"))
    print("slide size: %s x %s EMU%s" % (sz.group(1), sz.group(2),
          "  (A4)" if sz.group(1) == "9906000" else ""))
    nums = slide_nums(z)
    print("slides: %d" % len(nums))

    if a.layouts:
        print("\n== layouts ==")
        for n in sorted(x for x in z.namelist() if re.fullmatch(r'ppt/slideLayouts/slideLayout\d+\.xml', x)):
            x = read(z, n)
            nm = re.search(r'name="([^"]*)"', x)
            phs = sorted(set(re.findall(r'<p:ph[^>]*idx="(\d+)"', x)), key=int)
            print("  %-22s %-28s ph idx: %s" % (n.split("/")[-1], nm.group(1) if nm else "?",
                                                ",".join(phs) or "-"))

    if a.tokens:
        cols, fonts = collections.Counter(), collections.Counter()
        for n in nums:
            x = read(z, "ppt/slides/slide%d.xml" % n)
            cols.update(c.upper() for c in re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', x))
            fonts.update(re.findall(r'<a:latin typeface="([^"]+)"', x))
        print("\n== colours (top 12) ==")
        for c, k in cols.most_common(12):
            print("  #%s  x%d" % (c, k))
        if not cols:
            print("  (no literal RGB on slides -- colours come from the theme via"
                  " <a:schemeClr>; read ppt/theme/theme1.xml <a:clrScheme>)")
        print("== fonts ==")
        for f, k in fonts.most_common(6):
            print("  %-20s x%d" % (f, k))
        if not fonts:
            print("  (no explicit typeface on slides -- see <a:fontScheme> in ppt/theme/theme1.xml)")

    print("\n== per-slide ==")
    print("  %-5s %-16s %-6s %-5s %s" % ("slide", "layout", "media", "shps", "title"))
    for n in nums:
        rels = read(z, "ppt/slides/_rels/slide%d.xml.rels" % n)
        lay = re.search(r'slideLayout(\d+)\.xml', rels)
        media = len(re.findall(r'Target="\.\./media/', rels))
        x = read(z, "ppt/slides/slide%d.xml" % n)
        rows = list(shape_rows(x))
        title = next((r["text"] for r in rows if r["text"]), "")
        print("  %-5d %-16s %-6s %-5d %s" % (
            n, "slideLayout%s" % lay.group(1) if lay else "?",
            media or "-", len(rows), title[:46]))
    print("\n  media>0 のスライドは画像を抱えている → 書き直さず丸ごと流用する")

    if a.slides:
        for n in [int(s) for s in a.slides.split(",")]:
            print("\n== slide%d geometry ==" % n)
            for r in shape_rows(read(z, "ppt/slides/slide%d.xml" % n)):
                print("  %-22s %-12s %-42s fill=%-7s ph=%-4s %s" % (
                    r["name"][:22], r["prst"], str(r["xfrm"]) if r["xfrm"] else "(inherited)",
                    r["fill"] or "-", r["ph"] or "-", r["text"][:44]))


if __name__ == "__main__":
    sys.exit(main())
