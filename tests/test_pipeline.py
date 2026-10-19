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


# ---------------------------------------------------------------- профили по типам документов

from make_sample import build_dissertation  # noqa: E402
import catalog  # noqa: E402


def audit_with(path, profile_id):
    return Audit(nk.open_docx(path), nk.load_profile(profile_id)).run()


def fix_with(src, dst, profile_id):
    doc = nk.open_docx(src)
    Fixer(doc, nk.load_profile(profile_id)).run()
    doc.save(dst)


def test_catalog_is_consistent():
    assert catalog.check(catalog.load_catalog()) == []


def test_every_document_type_has_profile_or_is_planned():
    cat = catalog.load_catalog()
    for key, e in cat["document_types"].items():
        assert e.get("profile") or e.get("readiness") == "planned", key


def test_dissertation_profile(tmp_path):
    src, out = tmp_path / "diss.docx", tmp_path / "diss_fixed.docx"
    build_dissertation(str(src))
    pid = "gost-r-7.0.11-2011-dissertation"
    found = rules(audit_with(src, pid))
    for rule in ("page.margins", "page_numbers.missing", "text.size_max", "text.indent",
                 "headings.section.new_page", "headings.section.alignment", "headings.struct.new_page"):
        assert rule in found, rule
    # ГОСТ Р 7.0.11 не задаёт гарнитуру и регистр — такие правила не проверяем
    assert not {"text.font", "headings.struct.case", "headings.struct.bold"} & found
    fix_with(src, out, pid)
    assert rules(audit_with(out, pid), "auto") == set()
    # номер страницы — в верхнем колонтитуле
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    doc = nk.open_docx(out)
    headers = [r.target_part.element for r in doc.part.rels.values() if r.reltype == RT.HEADER]
    assert any("PAGE" in (t.text or "") for h in headers for t in h.iter(qn("w:instrText")))


def test_student_profile_does_not_require_nir_elements(sample):
    missing = [f.message for f in audit_with(sample, "student-work-gost-7.32-2017") if f.rule == "structure.missing"]
    assert not any("ИСПОЛНИТЕЛЕЙ" in m or "РЕФЕРАТ" in m for m in missing)
    nir = [f.message for f in audit_with(sample, "gost-7.32-2017") if f.rule == "structure.missing"]
    assert any("РЕФЕРАТ" in m for m in nir)


def test_null_unsets_inherited_rule(tmp_path):
    (tmp_path / "x.yaml").write_text("id: x\nextends: gost-7.32-2017\ntext:\n  font: null\n", encoding="utf-8")
    p = nk.load_profile(str(tmp_path / "x.yaml"))
    assert "font" not in p["text"] and p["text"]["size_pt_min"] == 12


def test_page_break_after_heading_text_is_not_before(tmp_path):
    from docx import Document
    from docx.enum.text import WD_BREAK

    doc = Document()
    doc.add_paragraph("Текст")
    doc.add_paragraph("ВВЕДЕНИЕ").runs[0].add_break(WD_BREAK.PAGE)
    blocks = nk.iter_blocks(doc)
    assert not nk.starts_new_page(doc, blocks, 1)
    doc.add_paragraph("ЗАКЛЮЧЕНИЕ")
    blocks = nk.iter_blocks(doc)
    assert nk.starts_new_page(doc, blocks, 2)


# ---------------------------------------------------------------- случаи из реальных диссертаций

def _doc():
    from docx import Document
    return Document()


def _classified(doc, profile_id="gost-r-7.0.11-2011-dissertation"):
    blocks = nk.iter_blocks(doc)
    nk.classify(doc, blocks, nk.load_profile(profile_id))
    return [b for b in blocks if b.kind == "p"]


def test_heading_style_on_paragraphs_does_not_make_them_chapters():
    doc = _doc()
    doc.add_paragraph("ВВЕДЕНИЕ")
    doc.add_heading("ГЛАВА 1. ОБЗОР", level=1)
    doc.add_heading("1.2 Механизмы", level=1)       # параграф, оформленный стилем главы
    doc.add_heading("Выводы 1 главы", level=1)
    levels = {b.text: b.extra.get("level") for b in _classified(doc) if b.role == "heading_section"}
    assert levels == {"ГЛАВА 1. ОБЗОР": 1, "1.2 Механизмы": 2, "Выводы 1 главы": 2}


def test_inline_appendix_title_ends_bibliography():
    doc = _doc()
    doc.add_paragraph("СПИСОК ЛИТЕРАТУРЫ")
    doc.add_paragraph("Иванов, И. И. Книга. – Москва, 2020. – 10 с.")
    doc.add_heading("Приложение А. Классификация механизмов", level=1)
    doc.add_paragraph("Текст приложения.")
    roles = [b.role for b in _classified(doc)]
    assert roles == ["heading_struct", "bib_entry", "heading_appendix", "body"]


def test_split_caption_and_continuation_roles():
    doc = _doc()
    doc.add_paragraph("ВВЕДЕНИЕ")
    doc.add_paragraph("Рисунок 8 – Траектория изменения уровня")
    doc.add_paragraph("отрасли России (1920-2021).")
    doc.add_paragraph("Окончание рисунка 7")
    roles = [b.role for b in _classified(doc)]
    assert roles[1:] == ["caption_figure", "caption_tail", "continuation"]


def test_formula_layout_table_needs_no_caption():
    doc = _doc()
    doc.add_paragraph("ВВЕДЕНИЕ")
    t = doc.add_table(rows=1, cols=2)
    t.cell(0, 0).text = "A = B + C"
    t.cell(0, 1).text = "(2)"
    found = Audit(doc, nk.load_profile("gost-r-7.0.11-2011-dissertation")).run()
    assert "table.no_caption" not in rules(found)


def test_landscape_sections_with_opposite_binding(tmp_path):
    from docx.enum.section import WD_ORIENT
    from docx.shared import Mm

    doc = _doc()
    s = doc.sections[0]
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.left_margin, s.right_margin, s.top_margin, s.bottom_margin = Mm(25), Mm(10), Mm(20), Mm(20)
    for top, bottom in ((25, 10), (10, 25)):
        ns = doc.add_section()
        ns.orientation = WD_ORIENT.LANDSCAPE
        ns.page_width, ns.page_height = Mm(297), Mm(210)
        ns.left_margin, ns.right_margin, ns.top_margin, ns.bottom_margin = Mm(20), Mm(20), Mm(top), Mm(bottom)
    found = rules(Audit(doc, nk.load_profile("gost-r-7.0.11-2011-dissertation")).run())
    assert "page.landscape_binding" in found and "page.landscape_margins" not in found


def test_citations_parsing():
    from audit import parse_citations

    assert parse_citations("см. [3; 7], [4–6], [10, с. 81], [2, с. 5; 9], [3, 7]") == [3, 7, 4, 5, 6, 10, 2, 9, 3, 7]
