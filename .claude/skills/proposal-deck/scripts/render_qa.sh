#!/usr/bin/env bash
# render_qa.sh <deck.pptx> [outdir]  -> page images for visual inspection.
# Re-run in full after every fix: pdftoppm reads the PDF, so the PDF must be
# regenerated from the edited .pptx or you will review stale pixels.
set -euo pipefail
deck="$1"; outdir="${2:-qa-img}"; here="$(cd "$(dirname "$0")" && pwd)"
pdf="${deck%.pptx}.pdf"
rm -rf "$outdir" "$pdf"; mkdir -p "$outdir"
python3 "$here/_soffice.py" --headless --convert-to pdf --outdir "$(dirname "$deck")" "$deck" >/dev/null
pdftoppm -jpeg -r 130 "$pdf" "$outdir/p"
echo "$(ls "$outdir" | wc -l) pages ->"; ls -1 "$PWD/$outdir"/p-*.jpg
