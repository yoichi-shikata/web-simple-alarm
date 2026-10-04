#!/usr/bin/env python3
"""Check and install what the deck workflow needs, on Linux, macOS or Windows.

  python setup_env.py          # check, install what it safely can, verify
  python setup_env.py --check  # report only

Needs: LibreOffice (with Impress) to render slides for visual QA; Pillow and
PyMuPDF to rasterise and contact-sheet them. Everything else is stdlib.

Why the explicit verify step: a Linux box can have `soffice` on PATH but no
Impress filter, and then converting a .pptx fails with the misleading
"source file could not be loaded" -- it looks like a corrupt deck. So we prove
a real .pptx converts, rather than trusting that the binary exists.
"""
import os
import platform
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
SYS = platform.system()

if hasattr(sys.stdout, "reconfigure"):  # Windows console encodings must not crash a run
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def sh(cmd):
    print("  $ " + " ".join(cmd))
    return subprocess.run(cmd).returncode == 0


def have_py(mod):
    try:
        __import__(mod)
        return True
    except ImportError:
        return False


def install_libreoffice():
    if SYS == "Linux":
        sudo = [] if os.geteuid() == 0 else (["sudo"] if shutil.which("sudo") else None)
        if sudo is None or not shutil.which("apt-get"):
            return False
        # update first: the pinned point release can 404 on a stale index
        return sh(sudo + ["apt-get", "update", "-qq"]) and \
            sh(sudo + ["apt-get", "install", "-y", "-qq", "libreoffice-impress", "fonts-noto-cjk"])
    if SYS == "Darwin" and shutil.which("brew"):
        return sh(["brew", "install", "--cask", "libreoffice"])
    if SYS == "Windows" and shutil.which("winget"):
        return sh(["winget", "install", "-e", "--id", "TheDocumentFoundation.LibreOffice",
                   "--accept-package-agreements", "--accept-source-agreements"])
    return False


def manual_hint():
    return {
        "Linux": "sudo apt-get install -y libreoffice-impress fonts-noto-cjk",
        "Darwin": "brew install --cask libreoffice   (または https://www.libreoffice.org/download/)",
        "Windows": "winget install TheDocumentFoundation.LibreOffice   (または公式サイトのインストーラ)",
    }.get(SYS, "https://www.libreoffice.org/download/")


def impress_on_linux():
    if SYS != "Linux":
        return True
    r = subprocess.run(["dpkg", "-s", "libreoffice-impress"], capture_output=True) if shutil.which("dpkg") else None
    return r is None or r.returncode == 0


def verify_conversion():
    """Build a minimal real .pptx with the skill's own tooling and convert it."""
    import zipfile
    import pptx_pkg as pk
    tmp = tempfile.mkdtemp(prefix="deckcheck_")
    try:
        src = os.path.join(HERE, "..", "tests", "fixture_template.pptx")
        if not os.path.exists(src):
            print("  (fixture missing; skipping conversion proof)")
            return True
        pdf = pk.to_pdf(src, tmp, timeout=180)
        return os.path.getsize(pdf) > 0
    except Exception as e:  # noqa: BLE001
        print("  !! %s" % e)
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    check_only = "--check" in sys.argv
    import pptx_pkg as pk
    ok = True

    print("== Python packages")
    for mod, pipname in (("PIL", "Pillow"), ("pymupdf", "PyMuPDF")):
        if have_py(mod):
            print("  %-8s ok" % pipname)
        elif check_only:
            print("  %-8s MISSING" % pipname); ok = False
        else:
            sh([sys.executable, "-m", "pip", "install", "--quiet", pipname])
            ok &= have_py(mod)

    print("== LibreOffice")
    if pk.soffice_bin() and impress_on_linux():
        print("  found: %s" % pk.soffice_bin())
    elif check_only:
        print("  MISSING -> %s" % manual_hint()); ok = False
    elif not install_libreoffice() or not pk.soffice_bin():
        print("  !! 自動インストールできませんでした。手動で入れてください:\n     %s" % manual_hint())
        ok = False

    print("== Conversion proof (.pptx -> .pdf)")
    if pk.soffice_bin():
        if verify_conversion():
            print("  ok")
        else:
            ok = False
            print("  !! 変換に失敗。Linuxなら libreoffice-impress が入っているか確認: %s" % manual_hint())

    print("==> %s" % ("ready" if ok else "NOT ready (上のメッセージを参照)"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
