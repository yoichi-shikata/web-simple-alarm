#!/usr/bin/env python3
"""Run soffice through the pptx skill's sandbox-aware wrapper.

A bare `soffice` call fails here (blocked AF_UNIX sockets, no bootstrappable
user profile), so always go through run_soffice rather than subprocess-ing
soffice directly.
"""
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from _pptxskill import ensure_on_path  # noqa: E402

ensure_on_path()
from office.soffice import run_soffice  # noqa: E402

sys.exit(run_soffice(sys.argv[1:]).returncode)
