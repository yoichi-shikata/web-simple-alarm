#!/usr/bin/env python3
"""End-to-end self-test of the skill's tooling. Run it once per new machine or
agent (Claude Code, Codex, ...) before a real job:

    python tests/run_e2e.py            # structure + rendering
    python tests/run_e2e.py --no-render

Exercises: unpack -> add_slide -> generate body XML -> edit cover -> reorder
(dropping two slides, one of which holds an image) -> clean -> pack ->
check_bounds -> validate -> text dump -> render. Exits non-zero on failure.
"""
import os
import shutil
import sys
import tempfile
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))
import deck_lib as d  # noqa: E402
import pptx_pkg as pk  # noqa: E402


def re_slide(name):
    return re.fullmatch(r"ppt/slides/slide\d+\.xml", name) is not None


def check(cond, msg):
    print(("  ok   " if cond else "  FAIL ") + msg)
    if not cond:
        sys.exit(1)


def main():
    render = "--no-render" not in sys.argv
    tmp = tempfile.mkdtemp(prefix="deck_e2e_")
    cwd = os.getcwd()
    try:
        os.chdir(tmp)
        shutil.copy(os.path.join(HERE, "fixture_template.pptx"), "template.pptx")
        d.unpack("template.pptx", "unpacked")
        d.configure_for("unpacked")          # fixture is 4:3, not A4
        check(pk.slide_order("unpacked") == ["slide%d" % i for i in range(1, 6)], "fixture has 5 slides")

        # an extra image-bearing slide we will delete, to prove media cleanup
        img_slide = pk.add_slide("unpacked", "slide5.xml")
        new = pk.add_slide("unpacked", "slide2.xml", after="slide1.xml")
        check(pk.slide_order("unpacked")[1] == new[:-4], "add_slide --after places the copy")

        i = d.ids()
        (x1, w1), (x2, w2) = d.cols(2)
        b = d.textbox(next(i), "T", d.L, 600000, d.W, 500000,
                      [d.P(d.R("セルフテスト", sz=2000, b=1, color=d.NAVY))])
        b += d.card(i, x1, d.TOP, w1, 2560320, "＜ご状況＞", ["・新設ライン"])
        b += d.card(i, x2, d.TOP, w2, 2560320, "＜課題＞", [("・目視では不可", 1, d.RED)], header_fill=d.RED)
        b += d.flow(i, [("A", d.MIDGRAY, d.WHITE), ("B", d.NAVY, d.WHITE)], x0=d.L, y=4400000)
        rows = [d.header_row(["区分", "仕様", "数量"])] + d.zebra(
            [("カメラ", "2,400万画素", "4台")], [{"b": 1}, {}, {"algn": "ctr"}])
        b += d.table(next(i), d.L, 5600000, [2000000, d.W - 4000000, 2000000], rows, [300000, 300000])
        d.build("unpacked/ppt/slides/" + new, "", "", b, with_placeholders=False)

        # a failing edit must not destroy the slide
        before = open("unpacked/ppt/slides/slide1.xml", encoding="utf-8").read()
        try:
            d.edit_slide("unpacked", "slide1", lambda x: d.set_para(x, "存在しない段落", "x"))
        except ValueError:
            pass
        check(open("unpacked/ppt/slides/slide1.xml", encoding="utf-8").read() == before,
              "failed edit leaves slide intact")
        d.edit_slide("unpacked", "slide1", lambda x: d.set_para(x, "前案件 表紙", "新案件 ご提案書"))

        d.reorder_slides("unpacked", ["slide1", new[:-4], "slide5"])   # drops 2,3,4 and the image copy
        gone = pk.clean("unpacked")
        check(img_slide in [os.path.basename(g) for g in gone], "dropped slide removed by clean")
        media_left = os.listdir("unpacked/ppt/media") if os.path.isdir("unpacked/ppt/media") else []
        check(len(media_left) == 1, "kept slide5's image survives, orphan copy's does not (media=%s)" % media_left)

        check(d.check_bounds("unpacked") == [], "no shape runs off the slide")
        d.pack("unpacked", "out.pptx")
        errs = pk.validate("out.pptx")
        for e in errs:
            print("       " + e)
        check(not errs, "structural validation")
        txt = pk.text("out.pptx")
        check("新案件 ご提案書" in txt and "セルフテスト" in txt and "会社概要" in txt, "expected text present")
        check("前案件" not in txt, "no leftovers from the template deal")
        with zipfile.ZipFile("out.pptx") as z:
            check(sum(1 for n in z.namelist() if n.startswith("ppt/slides/slide")) == 3, "exactly 3 slide parts")

        if render:
            pages = pk.render("out.pptx", "qa")
            check(len(pages) == 3, "rendered 3 pages")

        # the reference example, on a fixture shaped like the VRAIN template
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            "build_deck", os.path.join(HERE, "..", "examples", "two-station-line", "build_deck.py"))
        bd = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(bd)
        bd.main(os.path.join(HERE, "fixture_vrain_like.pptx"), "example.pptx")
        check(pk.validate("example.pptx") == [], "reference example builds and validates")
        with zipfile.ZipFile("example.pptx") as z:
            n = sum(1 for f in z.namelist() if re_slide(f))
        check(n == 30, "reference example has 30 slides (got %d)" % n)
        print("\nALL PASSED")
    finally:
        os.chdir(cwd)
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
