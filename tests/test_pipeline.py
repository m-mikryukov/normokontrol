"""Сквозная проверка: учебный документ → проверка → исправление → повторная проверка.

    pip install -r requirements.txt pytest && python3 -m pytest tests
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "tests"))

import nk_common as nk  # noqa: E402
from audit import Audit  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from fix_layout import Fixer  # noqa: E402
from make_sample import build  # noqa: E402
from nk_tracked import accept_all  # noqa: E402


@pytest.fixture()
def sample(tmp_path):
    path = tmp_path / "bad.docx"
    build(str(path))
    return path


def audit(path):
    profile = nk.load_profile("gost-7.32-2017")
    return Audit(nk.open_docx(path), profile).run()


def fix(src, dst, track=True):
    profile = nk.load_profile("gost-7.32-2017")
    doc = nk.open_docx(src)
    Fixer(doc, profile, track=track).run()
    doc.save(dst)


def rules(findings, fixer=None):
    return {f.rule for f in findings if fixer is None or f.fix == fixer}


def test_audit_finds_typical_errors(sample):
    found = rules(audit(sample))
    for rule in ("page.margins", "page.size", "page_numbers.missing", "text.font", "text.size",
                 "text.line_spacing", "text.alignment", "text.indent", "headings.struct.case",
                 "headings.struct.dot", "headings.struct.new_page", "headings.section.number_dot",
                 "figure.caption_format", "figure.caption_dot", "table.caption_position",
                 "figure.ref_missing", "bibliography.heading", "bibliography.ref_missing"):
        assert rule in found, rule


def test_fix_clears_all_auto_findings(sample, tmp_path):
    out = tmp_path / "fixed.docx"
    fix(sample, out)
    assert rules(audit(out), "auto") == set()
    # то, что на автора и Claude, остаётся в отчёте
    left = rules(audit(out))
    assert {"figure.ref_missing", "table.caption_position", "bibliography.heading"} <= left


def test_fix_is_tracked_and_keeps_text(sample, tmp_path):
    out = tmp_path / "fixed.docx"
    fix(sample, out)
    doc = nk.open_docx(out)
    body = doc.element.body
    assert list(body.iter(qn("w:pPrChange"))), "нет отслеживаемых правок абзацев"
    assert list(body.iter(qn("w:rPrChange"))), "нет отслеживаемых правок шрифта"
    assert list(body.iter(qn("w:del"))) and list(body.iter(qn("w:ins")))
    for ch in body.iter(qn("w:rPrChange")):
        assert ch.getparent()[-1] is ch

    # после принятия правок текст автора тот же, кроме служебных знаков оформления
    accept_all(body)
    texts = [nk.para_text(p) for p in body.iter(qn("w:p"))]
    assert "Рисунок 1 — Схема процесса" in texts
    assert "Таблица 1 — Исходные данные" in texts
    assert "1 Теоретическая часть" in texts
    assert any(t.startswith("Цель работы — анализ процессов [1].") for t in texts)


def test_fix_without_tracking(sample, tmp_path):
    out = tmp_path / "plain.docx"
    fix(sample, out, track=False)
    body = nk.open_docx(out).element.body
    assert not list(body.iter(qn("w:pPrChange")))
    assert not list(body.iter(qn("w:del")))
    assert rules(audit(out), "auto") == set()


def test_fix_is_idempotent(sample, tmp_path):
    once, twice = tmp_path / "1.docx", tmp_path / "2.docx"
    fix(sample, once)
    profile = nk.load_profile("gost-7.32-2017")
    doc = nk.open_docx(once)
    fixer = Fixer(doc, profile).run()
    assert fixer.rev.count == 0, fixer.summary()


def test_profile_inheritance(tmp_path):
    (tmp_path / "vuz-2026.yaml").write_text(
        "name: Тест\nid: vuz-2026\nextends: gost-7.32-2017\ntext:\n  size_pt_min: 14\n", encoding="utf-8")
    p = nk.load_profile(str(tmp_path / "vuz-2026.yaml"))
    assert p["text"]["size_pt_min"] == 14
    assert p["text"]["font"] == "Times New Roman"
    assert p["page"]["margins_mm"]["left"] == 30


@pytest.mark.skipif(not shutil.which("soffice"), reason="нет LibreOffice")
def test_libreoffice_opens_result(sample, tmp_path):
    out = tmp_path / "fixed.docx"
    fix(sample, out)
    subprocess.run([str(ROOT / "scripts" / "render_check.sh"), str(out), str(tmp_path / "render")],
                   check=True, capture_output=True, timeout=180)
    assert (tmp_path / "render" / "fixed.pdf").stat().st_size > 1000
