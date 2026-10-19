#!/usr/bin/env bash
# Рендер .docx в PDF и PNG через LibreOffice — проверить вёрстку глазами.
#
#   scripts/render_check.sh работа.docx [каталог]        # как есть, с пометками правок
#   ACCEPT=1 scripts/render_check.sh работа.docx [каталог]  # с принятыми правками
#
# Результат: <каталог>/<имя>.pdf и <каталог>/<имя>-pages/page-N.png
set -euo pipefail

src="$1"
out="${2:-out}"
here="$(cd "$(dirname "$0")" && pwd)"
mkdir -p "$out"
name="$(basename "${src%.*}")"

if [[ "${ACCEPT:-0}" == "1" ]]; then
  accepted="$out/$name-accepted.docx"
  python3 "$here/accept_changes.py" "$src" -o "$accepted"
  src="$accepted"
  name="$name-accepted"
fi

# отдельный профиль LibreOffice: не конфликтует с открытым офисом
profile_dir="$(mktemp -d)"
trap 'rm -rf "$profile_dir"' EXIT
soffice -env:UserInstallation="file://$profile_dir" --headless --convert-to pdf \
  --outdir "$out" "$src" >/dev/null 2>&1

pdf="$out/$name.pdf"
[[ -f "$pdf" ]] || { echo "Не удалось сконвертировать $src" >&2; exit 1; }

if command -v pdftoppm >/dev/null; then
  rm -rf "$out/$name-pages" && mkdir -p "$out/$name-pages"
  pdftoppm -r "${DPI:-60}" -png "$pdf" "$out/$name-pages/page"
fi
pages="$(pdfinfo "$pdf" 2>/dev/null | awk '/^Pages:/ {print $2}')"
echo "PDF: $pdf (страниц: ${pages:-?})"
[[ -d "$out/$name-pages" ]] && echo "PNG: $out/$name-pages/"
