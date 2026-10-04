#!/usr/bin/env bash
# Bootstrap everything the deck workflow needs. Safe to re-run.
#
# Why this exists: the cloud container ships libreoffice-core WITHOUT Impress,
# so `soffice --convert-to pdf` on a .pptx fails with the misleading
# "source file could not be loaded" — it looks like a corrupt deck, but the
# filter is simply missing. pdftoppm, the CJK fonts and the python tooling
# are absent too. Diagnosing this from scratch costs ~6 turns; this costs one.
set -uo pipefail

need_apt=0
command -v pdftoppm >/dev/null 2>&1 || need_apt=1
dpkg -s libreoffice-impress >/dev/null 2>&1 || need_apt=1
fc-list 2>/dev/null | grep -qi "noto.*cjk" || need_apt=1

if [ "$need_apt" = "1" ]; then
  echo "==> apt: libreoffice-impress / poppler-utils / fonts-noto-cjk"
  apt-get update -qq
  # The pinned point release can 404 until `apt-get update` runs, hence the order.
  apt-get install -y -qq libreoffice-impress poppler-utils fonts-noto-cjk
else
  echo "==> apt deps already present"
fi

python3 - <<'PY' || pip install --quiet "markitdown[pptx]" python-pptx defusedxml Pillow lxml
import pptx, markitdown, defusedxml, PIL, lxml
PY

echo "==> verifying"
command -v pdftoppm >/dev/null && echo "  pdftoppm   ok"
command -v markitdown >/dev/null && echo "  markitdown ok"
python3 -c "import pptx;print('  python-pptx ok', pptx.__version__)"
fc-list 2>/dev/null | grep -ci "noto.*cjk" | xargs -I{} echo "  CJK fonts  {} faces"

# Prove the Impress filter actually loads a .pptx — the check that matters.
tmp=$(mktemp -d)
python3 -c "
from pptx import Presentation
p=Presentation(); s=p.slides.add_slide(p.slide_layouts[0]); s.shapes.title.text='検証'
p.save('$tmp/t.pptx')"
if python3 "$(dirname "$0")/_soffice.py" --headless --convert-to pdf --outdir "$tmp" "$tmp/t.pptx" >/dev/null 2>&1 \
   && [ -s "$tmp/t.pdf" ]; then
  echo "  soffice→pdf ok"
else
  echo "  !! soffice could not convert a .pptx — Impress still missing?" >&2
  rm -rf "$tmp"; exit 1
fi
rm -rf "$tmp"
echo "==> ready"
