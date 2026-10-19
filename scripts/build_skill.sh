#!/usr/bin/env bash
# Собирает самодостаточный скилл для загрузки в claude.ai (Настройки → Возможности → Навыки).
#
#   scripts/build_skill.sh [out/normokontrol-skill.zip]
#
# В архиве: SKILL.md, scripts/, profiles/, references/ — каталог скилла и есть NK.
set -euo pipefail

root="$(cd "$(dirname "$0")/.." && pwd)"
zip_path="$(realpath -m "${1:-$root/out/normokontrol-skill.zip}")"
tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT

dst="$tmp/normokontrol"
mkdir -p "$dst"
cp "$root/.claude/skills/normokontrol/SKILL.md" "$dst/"
cp -r "$root/profiles" "$root/references" "$dst/"
mkdir -p "$dst/scripts"
cp "$root"/scripts/*.py "$root"/scripts/*.sh "$dst/scripts/"
rm -f "$dst/scripts/build_skill.sh"
cp "$root/requirements.txt" "$dst/"

mkdir -p "$(dirname "$zip_path")"
rm -f "$zip_path"
(cd "$tmp" && zip -qr "$zip_path" normokontrol)
echo "Скилл собран: $zip_path"
