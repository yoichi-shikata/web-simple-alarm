#!/usr/bin/env python3
"""Deck operations CLI. Same commands under Claude Code, Codex, or a terminal.

  pptx_tool.py add_slide  unpacked/ slide4.xml [--after slide4.xml]
  pptx_tool.py clean      unpacked/
  pptx_tool.py validate   out.pptx
  pptx_tool.py text       out.pptx
  pptx_tool.py render     out.pptx [outdir]        -> page images for review
  pptx_tool.py thumbnail  deck.pptx [prefix]       -> labelled contact sheets
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pptx_pkg as pk  # noqa: E402


def main(argv):
    if len(argv) < 2:
        sys.exit(__doc__)
    cmd, args = argv[1], argv[2:]
    if cmd == "add_slide":
        after = None
        if "--after" in args:
            k = args.index("--after"); after = args[k + 1]; args = args[:k] + args[k + 2:]
        new = pk.add_slide(args[0], args[1], after=after)
        print("Created ppt/slides/%s from %s" % (new, args[1]))
    elif cmd == "clean":
        gone = pk.clean(args[0])
        print("removed %d orphaned part(s)" % len(gone))
        for g in gone:
            print("  " + g)
    elif cmd == "validate":
        errs = pk.validate(args[0])
        if errs:
            for e in errs:
                print("FAIL " + e)
            sys.exit(1)
        print("All validations PASSED!")
    elif cmd == "text":
        sys.stdout.write(pk.text(args[0]) + "\n")
    elif cmd == "render":
        paths = pk.render(args[0], args[1] if len(args) > 1 else "qa-img")
        print("%d pages ->" % len(paths))
        for p in paths:
            print(os.path.abspath(p))
    elif cmd == "thumbnail":
        for p in pk.thumbnail(args[0], args[1] if len(args) > 1 else "thumbs"):
            print(p)
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main(sys.argv)
