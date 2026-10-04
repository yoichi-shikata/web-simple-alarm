#!/usr/bin/env python3
"""Install this skill for the agents on this machine.

  python install.py              # Codex (both current and legacy paths)
  python install.py --claude     # + Claude Code user skills
  python install.py --all        # Codex + Claude Code
  python install.py --dry-run

Codex reads skills from ~/.agents/skills (current) and older builds from
~/.codex/skills; Claude Code reads ~/.claude/skills. We copy real files rather
than symlinking: symlinks are unreliable on Windows and some agents do not
follow them when scanning for SKILL.md.
"""
import os
import shutil
import sys

SKILL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NAME = os.path.basename(SKILL)
IGNORE = shutil.ignore_patterns("__pycache__", "*.pyc", "*_work", "qa", "*.pdf")


def main():
    args = set(sys.argv[1:])
    targets = ["~/.agents/skills", "~/.codex/skills"]
    if "--claude" in args or "--all" in args:
        targets.append("~/.claude/skills")
    for t in targets:
        dest = os.path.join(os.path.expanduser(t), NAME)
        if os.path.abspath(dest) == SKILL:
            print("skip (source)  %s" % dest)
            continue
        print("%s %s" % ("would install" if "--dry-run" in args else "install      ", dest))
        if "--dry-run" in args:
            continue
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        if os.path.exists(dest):
            shutil.rmtree(dest)
        shutil.copytree(SKILL, dest, ignore=IGNORE)
    if "--dry-run" not in args:
        print("\n次に:  python %s/scripts/setup_env.py   （LibreOffice 等の準備）" %
              os.path.join(os.path.expanduser(targets[0]), NAME))
        print("Codex では  $%s  で呼び出せます。" % NAME)


if __name__ == "__main__":
    main()
