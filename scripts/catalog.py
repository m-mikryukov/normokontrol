#!/usr/bin/env python3
"""Что умеем оформлять: типы документов, профили и стандарты с их статусом.

    python3 scripts/catalog.py            # меню для первого вопроса пользователю
    python3 scripts/catalog.py --standards # стандарты: статус, дата введения, когда проверен
    python3 scripts/catalog.py --check     # согласованность каталога и профилей (для тестов)
"""

from __future__ import annotations

import argparse
import sys
from datetime import date

import yaml

import nk_common as nk

READINESS = {"ready": "готово", "beta": "бета", "limited": "частично", "planned": "не поддерживается"}
STALE_DAYS = 365


def load_catalog() -> dict:
    return yaml.safe_load((nk.PROFILES_DIR / "catalog.yaml").read_text(encoding="utf-8"))


def profile_files():
    return sorted(p for p in nk.PROFILES_DIR.glob("*.yaml") if p.name != "catalog.yaml")


def derived_profiles(base_id: str) -> list[dict]:
    """Профили вузов и советов, наследующие base_id (напрямую или через цепочку)."""
    out = []
    for path in profile_files():
        raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        chain, cur = [], raw
        while cur and cur.get("extends"):
            chain.append(cur["extends"])
            try:
                cur = yaml.safe_load(nk._find_profile(cur["extends"], path.parent).read_text(encoding="utf-8"))
            except FileNotFoundError:
                break
        if base_id in chain and raw.get("id") and not raw.get("document_types"):
            out.append(raw)
    return out


def readiness(entry: dict) -> str:
    if entry.get("readiness"):
        return entry["readiness"]
    if not entry.get("profile"):
        return "planned"
    return nk.load_profile(entry["profile"]).get("readiness", "beta")


def render_menu(cat: dict) -> str:
    lines = ["Что нужно оформить?", ""]
    for n, (key, e) in enumerate(cat["document_types"].items(), 1):
        state = READINESS[readiness(e)]
        if e.get("profile"):
            prof = nk.load_profile(e["profile"])
            stds = ", ".join(cat["standards"][s]["title"].split(" «")[0] for s in prof.get("standards", []))
            lines.append(f"{n}. {e['title']} — {state}; {stds}; профиль `{e['profile']}`")
        else:
            lines.append(f"{n}. {e['title']} — {state}")
        if e.get("note"):
            lines.append(f"   {' '.join(e['note'].split())}")
        for d in derived_profiles(e["profile"]) if e.get("profile") else []:
            lines.append(f"   • есть профиль: {d.get('name')} (`{d['id']}`)")
    return "\n".join(lines) + "\n"


def render_standards(cat: dict, today: date | None = None) -> str:
    today = today or date.today()
    lines = ["| Стандарт | Для чего | Статус | Действует с | Проверено |", "|---|---|---|---|---|"]
    for key, s in list(cat["standards"].items()) + list(cat.get("regulations", {}).items()):
        checked = s.get("checked")
        stale = checked and (today - checked).days > STALE_DAYS
        mark = f"{checked} ⚠ перепроверить" if stale else str(checked)
        since = s.get("in_force_since", "—")
        lines.append(f"| {s['title'].split(' «')[0]} | {' '.join(s['scope'].split())} | "
                     f"{' '.join(s['status'].split())} | {since} | {mark} |")
    return "\n".join(lines) + "\n"


def check(cat: dict) -> list[str]:
    errors = []
    known_types = set(cat["document_types"])
    for key, e in cat["document_types"].items():
        if not e.get("profile"):
            continue
        try:
            prof = nk.load_profile(e["profile"])
        except FileNotFoundError:
            errors.append(f"{key}: нет профиля {e['profile']}")
            continue
        if key not in prof.get("document_types", []):
            errors.append(f"{key}: профиль {e['profile']} не заявляет этот тип документа")
        for s in prof.get("standards", []):
            if s not in cat["standards"]:
                errors.append(f"{e['profile']}: стандарт {s} не описан в каталоге")
    for path in profile_files():
        raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        if raw.get("id") != path.stem:
            errors.append(f"{path.name}: id «{raw.get('id')}» не совпадает с именем файла")
        for t in raw.get("document_types", []) or []:
            if t not in known_types:
                errors.append(f"{path.name}: неизвестный тип документа {t}")
        if raw.get("readiness") and raw["readiness"] not in READINESS:
            errors.append(f"{path.name}: неизвестная готовность {raw['readiness']}")
    for key, s in cat["standards"].items():
        for field in ("title", "scope", "status", "checked", "source"):
            if not s.get(field):
                errors.append(f"стандарт {key}: не заполнено поле {field}")
    return errors


def main(argv=None):
    ap = argparse.ArgumentParser(description="Каталог документов и стандартов")
    ap.add_argument("--standards", action="store_true", help="таблица стандартов со статусом")
    ap.add_argument("--check", action="store_true", help="проверить согласованность каталога")
    args = ap.parse_args(argv)
    cat = load_catalog()
    if args.check:
        errors = check(cat)
        for e in errors:
            print(f"ошибка: {e}")
        print("Каталог согласован." if not errors else f"Ошибок: {len(errors)}")
        return 1 if errors else 0
    sys.stdout.write(render_standards(cat) if args.standards else render_menu(cat))
    return 0


if __name__ == "__main__":
    sys.exit(main())
