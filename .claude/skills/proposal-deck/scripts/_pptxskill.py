# -*- coding: utf-8 -*-
"""Locate the bundled `pptx` skill's scripts directory.

Its path contains a per-install UUID (…/skills/synced/<uuid>/pptx/scripts) that
differs between sessions, so hardcoding it breaks on the next container. Glob
for it instead, and fail with an actionable message rather than a stray
ModuleNotFoundError deep in a conversion.
"""
import glob
import os
import sys

_CANDIDATES = (
    "~/.claude/skills/synced/*/pptx/scripts",
    "~/.claude/skills/*/pptx/scripts",
    "~/.claude/skills/pptx/scripts",
    "/root/.claude/skills/synced/*/pptx/scripts",
)


def scripts_dir():
    for pat in _CANDIDATES:
        for hit in sorted(glob.glob(os.path.expanduser(pat))):
            if os.path.isdir(os.path.join(hit, "office")):
                return hit
    raise RuntimeError(
        "pptx skill scripts not found. Looked in: %s\n"
        "The proposal-deck skill reuses that skill's add_slide/clean/validate/"
        "thumbnail helpers." % ", ".join(_CANDIDATES))


def ensure_on_path():
    d = scripts_dir()
    if d not in sys.path:
        sys.path.insert(0, d)
    return d


def tool(name):
    """Absolute path to one of the pptx skill's scripts, e.g. tool('clean.py')."""
    p = os.path.join(scripts_dir(), name)
    if not os.path.exists(p):
        raise RuntimeError("pptx skill has no %s (looked at %s)" % (name, p))
    return p
