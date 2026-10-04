#!/usr/bin/env python3
"""Thin passthrough to the pptx skill's helper scripts, path-independent.

  pptx_tool.py add_slide  unpacked/ slide4.xml [--after slideN.xml]
  pptx_tool.py clean      unpacked/
  pptx_tool.py validate   out.pptx --original template.pptx
  pptx_tool.py thumbnail  deck.pptx prefix
"""
import os
import pathlib
import subprocess
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from _pptxskill import scripts_dir, tool  # noqa: E402

SCRIPTS = {"add_slide": "add_slide.py", "clean": "clean.py",
           "validate": "office/validate.py", "thumbnail": "thumbnail.py"}

if len(sys.argv) < 2 or sys.argv[1] not in SCRIPTS:
    sys.exit(__doc__)

env = dict(os.environ, PYTHONPATH=scripts_dir() + os.pathsep + os.environ.get("PYTHONPATH", ""))
sys.exit(subprocess.run([sys.executable, tool(SCRIPTS[sys.argv[1]])] + sys.argv[2:], env=env).returncode)
