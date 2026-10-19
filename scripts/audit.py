#!/usr/bin/env python3
"""Проверка оформления .docx по профилю. Документ не меняется.

    python3 scripts/audit.py работа.docx [--profile gost-7.32-2017] [--json out.json]

Каждое нарушение помечено, кто его исправляет:
  авто   — scripts/fix_layout.py;
  Claude — смысловое оформление по скиллу (подписи, заголовки, библиография);
  автор  — содержание, которое инструмент не трогает.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import OrderedDict
from dataclasses import asdict, dataclass

from docx.enum.text import WD_ALIGN_PARAGRAPH
import nk_common as nk

FIXERS = {"auto": "авто", "claude": "Claude", "author": "автор"}
SEVERITY = {"error": "ошибка", "warning": "предупреждение", "info": "замечание"}
CATEGORIES = ["Страница", "Нумерация страниц", "Основной текст", "Заголовки", "Рисунки",
              "Таблицы", "Формулы", "Ссылки и источники", "Структура"]

ALIGN = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY, "center": WD_ALIGN_PARAGRAPH.CENTER,
         "left": WD_ALIGN_PARAGRAPH.LEFT, "right": WD_ALIGN_PARAGRAPH.RIGHT}
ALIGN_RU = {WD_ALIGN_PARAGRAPH.JUSTIFY: "по ширине", WD_ALIGN_PARAGRAPH.CENTER: "по центру",
            WD_ALIGN_PARAGRAPH.LEFT: "по левому краю", WD_ALIGN_PARAGRAPH.RIGHT: "по правому краю",
            WD_ALIGN_PARAGRAPH.DISTRIBUTE: "распределённое", None: "по левому краю"}

A4 = (210.0, 297.0)


def parse_citations(text: str) -> list[int]:
    """Номера источников из ссылок [5], [5, с. 12], [3; 7], [4–6], [2, с. 5; 9]."""
    cited: list[int] = []
    for m in re.finditer(r"\[([^\[\]]{1,200})\]", text):
        body = m.group(1)
        if not re.match(r"^\s*\d", body):
            continue
        parts = re.split(r";", body)
        if len(parts) == 1 and not re.search(r"\b(с|С|p|P)\.", body):
            parts = re.split(r",", body)  # [3, 7] без страниц
        for part in parts:
            rng = re.match(r"^\s*(\d+)\s*[–—\-]\s*(\d+)\s*$", part)
            one = re.match(r"^\s*(\d+)\b", part)
            if rng:
                a, b = int(rng.group(1)), int(rng.group(2))
                if 0 < b - a < 50:
                    cited += list(range(a, b + 1))
            elif one:
                cited.append(int(one.group(1)))
    return cited


def is_layout_table(b) -> bool:
    """Таблица-обёртка для формулы и её номера — подпись ей не нужна."""
    el = b.el
    if any(True for _ in el.iter(f"{{{nk.M_NS}}}oMath")):
        return True
    texts = [nk.para_text(p).strip() for p in el.iter(nk.qn("w:p"))]
    return any(re.fullmatch(r"\([А-ЯA-Z]?\.?\d+(\.\d+)*\)", t) for t in texts if t)


def caption_spacing_ok(doc, b, need: float) -> bool:
    ls = nk.eff_para(doc, b.obj, "line_spacing")
    if ls is None:  # не задан нигде — в Word это одинарный
        ls = 1.0
    return isinstance(ls, float) and abs(ls - need) <= 0.01


@dataclass
class Finding:
    rule: str
    category: str
    severity: str
    fix: str
    message: str
    where: str | None = None


class Audit:
    def __init__(self, doc, profile: dict):
        self.doc = doc
        self.p = profile
        self.fonts = nk.Fonts(doc)
        self.blocks = nk.iter_blocks(doc)
        nk.classify(doc, self.blocks, profile)
        self.findings: list[Finding] = []
        self._landscape_binding: dict[str, list[int]] = {}

    def add(self, rule, category, severity, fix, message, where=None):
        self.findings.append(Finding(rule, category, severity, fix, message, where))

    # ------------------------------------------------------------ запуск

    def run(self) -> list[Finding]:
        self.check_page()
        self._report_landscape_binding()
        self.check_page_numbers()
        for i, b in enumerate(self.blocks):
            if b.kind != "p":
                continue
            role = b.role
            if role in ("body", "list", "bib_entry"):
                self.check_text(b)
                self.check_runs(b, "Основной текст")
                if role == "body":
                    self.check_pseudo_heading(b)
            elif role == "caption_tail":
                cat = "Рисунки" if self.blocks[i - 1].role == "caption_figure" else "Таблицы"
                self.add("caption.split", cat, "error", "claude",
                         "Подпись разбита клавишей Enter на два абзаца: вторая строка стала обычным текстом. "
                         "Нужен один абзац (перенос строки — автоматический).", b.where())
            elif role == "note":
                self.check_runs(b, "Основной текст", size=False)
            elif role == "heading_struct":
                self.check_struct_heading(i, b)
            elif role == "heading_appendix":
                self.check_appendix_heading(i, b)
            elif role == "heading_section":
                self.check_section_heading(i, b)
            elif role == "caption_figure":
                self.check_caption(i, b, "figure")
            elif role == "caption_table":
                self.check_caption(i, b, "table")
            elif role == "figure":
                self.check_figure(i, b)
            elif role == "table_text":
                self.check_runs(b, "Таблицы", size=False,
                                size_min=nk.prof(self.p, "text.table_text_size_pt_min"))
        self.check_tables_have_captions()
        self.check_merged_captions()
        self.check_structure()
        self.check_cross_refs()
        self.check_bibliography()
        return self.findings

    # ------------------------------------------------------------ страница

    def check_page(self):
        tol = nk.prof(self.p, "page.tolerance_mm", 0.5)
        margins = nk.prof(self.p, "page.margins_mm", {})
        names = {"left": "левое", "right": "правое", "top": "верхнее", "bottom": "нижнее"}
        for n, s in enumerate(self.doc.sections, 1):
            w, h = nk.mm(s.page_width), nk.mm(s.page_height)
            landscape = w > h
            if nk.prof(self.p, "page.size") == "A4":
                short, long_ = sorted((w, h))
                if abs(short - A4[0]) > 1 or abs(long_ - A4[1]) > 1:
                    self.add("page.size", "Страница", "error", "auto",
                             f"Раздел {n}: формат {w:.0f}×{h:.0f} мм, нужен A4 (210×297 мм).")
            if landscape:
                self._check_landscape(n, s, margins, tol)
                continue
            for side, need in margins.items():
                have = nk.mm(getattr(s, f"{side}_margin"))
                if abs(have - need) > tol:
                    self.add("page.margins", "Страница", "error", "auto",
                             f"Раздел {n}: {names[side]} поле {have:.1f} мм, нужно {need} мм.")

    def _check_landscape(self, n, s, margins, tol):
        """Альбомный лист: поля портретной страницы, повёрнутые на 90°.

        Поле под переплёт (левое) у альбомного листа оказывается сверху или снизу —
        стандарт сторону не задаёт, но во всей работе она должна быть одна.
        """
        if not margins:
            return
        have = {k: nk.mm(getattr(s, f"{k}_margin")) for k in ("left", "right", "top", "bottom")}
        want_tb = sorted((margins.get("left", 0), margins.get("right", 0)))
        want_lr = sorted((margins.get("top", 0), margins.get("bottom", 0)))
        ok_tb = all(abs(a - b) <= tol for a, b in zip(sorted((have["top"], have["bottom"])), want_tb))
        ok_lr = all(abs(a - b) <= tol for a, b in zip(sorted((have["left"], have["right"])), want_lr))
        if not (ok_tb and ok_lr):
            self.add("page.landscape_margins", "Страница", "error", "claude",
                     f"Раздел {n} (альбомный): поля {have['left']:.0f}/{have['right']:.0f}/{have['top']:.0f}/"
                     f"{have['bottom']:.0f} мм (л/п/в/н) не совпадают с повёрнутыми полями книжного листа.")
            return
        binding = "top" if have["top"] > have["bottom"] else "bottom"
        self._landscape_binding.setdefault(binding, []).append(n)

    def _report_landscape_binding(self):
        sides = self._landscape_binding
        if len(sides) > 1:
            minority = min(sides.values(), key=len)
            side = "снизу" if minority is sides.get("bottom") else "сверху"
            self.add("page.landscape_binding", "Страница", "warning", "claude",
                     f"Альбомные листы повёрнуты по-разному: в разделах {', '.join(map(str, minority))} поле "
                     f"под переплёт {side}, в остальных — с другой стороны. На этих листах широкое поле "
                     "окажется не со стороны переплёта.")

    # ------------------------------------------------------------ номера страниц

    def _page_field_paragraphs(self):
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.oxml.ns import qn

        found = []
        for rel in self.doc.part.rels.values():
            if rel.reltype not in (RT.FOOTER, RT.HEADER) or rel.is_external:
                continue
            root = rel.target_part.element
            for p_el in root.iter(qn("w:p")):
                instr = " ".join([f.get(qn("w:instr")) or "" for f in p_el.iter(qn("w:fldSimple"))] +
                                 [t.text or "" for t in p_el.iter(qn("w:instrText"))])
                if re.search(r"\bPAGE\b", instr):
                    found.append((rel.reltype == RT.FOOTER, p_el))
        return found

    def check_page_numbers(self):
        cfg = nk.prof(self.p, "page_numbers")
        if not cfg:
            return
        found = self._page_field_paragraphs()
        if not found:
            where = "верхнем" if cfg.get("position", "").startswith("top") else "нижнем"
            self.add("page_numbers.missing", "Нумерация страниц", "error", "auto",
                     f"Нет номера страницы (нужен в {where} колонтитуле).")
            return
        pos = cfg.get("position", "bottom_center")
        for in_footer, p_el in found:
            if pos.startswith("bottom") and not in_footer:
                self.add("page_numbers.position", "Нумерация страниц", "error", "claude",
                         "Номер страницы в верхнем колонтитуле, нужен в нижнем.")
            if pos.startswith("top") and in_footer:
                self.add("page_numbers.position", "Нумерация страниц", "error", "claude",
                         "Номер страницы в нижнем колонтитуле, нужен в верхнем.")
            jc = p_el.find(f"{nk.qn('w:pPr')}/{nk.qn('w:jc')}")
            jc_val = None if jc is None else jc.get(nk.qn("w:val"))
            if pos.endswith("center") and jc_val != "center" and "\t" not in nk.para_text(p_el):
                self.add("page_numbers.alignment", "Нумерация страниц", "error", "auto",
                         "Номер страницы не по центру.")
        if cfg.get("title_page_shown") is False and self._has_title_page():
            s = self.doc.sections[0]
            if not s.different_first_page_header_footer:
                self.add("page_numbers.title_page", "Нумерация страниц", "error", "auto",
                         "На титульном листе будет виден номер страницы: включите «особый колонтитул "
                         "для первой страницы».")
        for s in self.doc.sections:
            fmt = s._sectPr.find(nk.qn("w:pgNumType"))
            val = None if fmt is None else fmt.get(nk.qn("w:fmt"))
            if cfg.get("style") == "arabic" and val not in (None, "decimal"):
                self.add("page_numbers.style", "Нумерация страниц", "error", "claude",
                         f"Формат номеров страниц «{val}», нужны арабские цифры.")
            start = None if fmt is None else fmt.get(nk.qn("w:start"))
            if cfg.get("title_page_counted") and s is self.doc.sections[0] and start not in (None, "1"):
                self.add("page_numbers.start", "Нумерация страниц", "warning", "claude",
                         f"Нумерация начинается с {start}: титульный лист должен входить в общую нумерацию.")

    def _has_title_page(self) -> bool:
        return any(b.role == "title" for b in self.blocks)

    # ------------------------------------------------------------ текст

    def check_text(self, b: nk.Block):
        t = self.p.get("text", {})
        if "line_spacing" in t:
            ls = nk.eff_para(self.doc, b.obj, "line_spacing")
            if not isinstance(ls, float):
                desc = "одинарный" if ls is None else f"точно {ls.pt:.0f} пт"
                self.add("text.line_spacing", "Основной текст", "error", "auto",
                         f"Межстрочный интервал {desc}, нужен {t['line_spacing']}.", b.where())
            elif abs(ls - t["line_spacing"]) > 0.01:
                self.add("text.line_spacing", "Основной текст", "error", "auto",
                         f"Межстрочный интервал {ls:g}, нужен {t['line_spacing']}.", b.where())
        if "alignment" in t:
            al = nk.eff_para(self.doc, b.obj, "alignment")
            need = ALIGN[t["alignment"]]
            if (al or WD_ALIGN_PARAGRAPH.LEFT) != need:
                self.add("text.alignment", "Основной текст", "error", "auto",
                         f"Выравнивание {ALIGN_RU.get(al, al)}, нужно {ALIGN_RU[need]}.", b.where())
        if "first_line_indent_cm" in t and b.role == "body":
            self._check_indent(b, t["first_line_indent_cm"], "text.indent", "Основной текст")

    def _check_indent(self, b, need_cm, rule, category, fix="auto"):
        lo, hi = nk.indent_range(need_cm)
        ind = nk.eff_para(self.doc, b.obj, "first_line_indent")
        have = 0.0 if ind is None else ind.cm
        if have < lo - 0.05 or have > hi + 0.05:
            need = f"{lo:g} см" if lo == hi else f"от {lo:g} до {hi:g} см"
            self.add(rule, category, "error", fix,
                     f"Абзацный отступ {have:.2f} см, нужен {need}.", b.where())

    def check_runs(self, b: nk.Block, category: str, size=True, size_min=None):
        t = self.p.get("text", {})
        font = t.get("font")
        size_min = size_min if size_min is not None else (t.get("size_pt_min") if size else None)
        size_max = t.get("size_pt_max") if size and category == "Основной текст" else None
        bad_fonts, bad_sizes, big_sizes, colored = set(), set(), set(), False
        for r in nk.text_runs(b):
            if font:
                f = self.fonts.font(r, b.obj)
                if f != font:
                    bad_fonts.add(f)
            if size_min:
                sz = self.fonts.size_pt(r, b.obj)
                # индексы и сноски допускаются мельче
                vert = r.find(f"{nk.qn('w:rPr')}/{nk.qn('w:vertAlign')}")
                if sz < size_min and vert is None:
                    bad_sizes.add(sz)
            if size_max and self.fonts.size_pt(r, b.obj) > size_max:
                big_sizes.add(self.fonts.size_pt(r, b.obj))
            if t.get("color") == "black" and not nk.is_black(self.fonts.color(r, b.obj)):
                colored = True
        if bad_fonts:
            self.add("text.font", category, "error", "auto",
                     f"Шрифт {', '.join(sorted(bad_fonts))}, нужен {font}.", b.where())
        if bad_sizes:
            self.add("text.size", category, "error", "auto",
                     f"Кегль {', '.join(f'{s:g}' for s in sorted(bad_sizes))} пт, нужно не меньше {size_min} пт.",
                     b.where())
        if big_sizes:
            self.add("text.size_max", category, "error", "auto",
                     f"Кегль {', '.join(f'{s:g}' for s in sorted(big_sizes))} пт, нужно не больше {size_max} пт.",
                     b.where())
        if colored:
            self.add("text.color", category, "error", "auto", "Цветной текст, нужен чёрный.", b.where())

    def check_pseudo_heading(self, b: nk.Block):
        text = b.text.strip()
        if len(text) > 120 or text[-1:] in ".,;:!?" or not text:
            return
        runs = nk.text_runs(b)
        all_bold = bool(runs) and all(self.fonts.flag(r, b.obj, "b") for r in runs)
        numbered = re.match(r"^\d+(\.\d+)*\.?\s+\S", text)
        if all_bold or numbered and len(text) < 80:
            self.add("headings.no_style", "Заголовки", "warning", "claude",
                     "Похоже на заголовок, но стиль «Заголовок» не применён — он не попадёт в содержание.",
                     b.where())

    # ------------------------------------------------------------ заголовки

    def _heading_common(self, b, cfg, rule):
        text = b.text.strip()
        if cfg.get("trailing_dot") is False and text.endswith(".") and not text.endswith(".."):
            self.add(f"{rule}.dot", "Заголовки", "error", "auto", "Точка в конце заголовка.", b.where())
        if "alignment" in cfg:
            al = nk.eff_para(self.doc, b.obj, "alignment") or WD_ALIGN_PARAGRAPH.LEFT
            if al != ALIGN[cfg["alignment"]]:
                self.add(f"{rule}.alignment", "Заголовки", "error", "auto",
                         f"Заголовок выровнен {ALIGN_RU.get(al, al)}, нужно {ALIGN_RU[ALIGN[cfg['alignment']]]}.",
                         b.where())
        runs = nk.text_runs(b)
        if cfg.get("bold") and runs and not all(self.fonts.flag(r, b.obj, "b") for r in runs):
            self.add(f"{rule}.bold", "Заголовки", "error", "auto", "Заголовок не полужирный.", b.where())
        self.check_runs(b, "Заголовки")

    def check_struct_heading(self, i, b):
        cfg = nk.prof(self.p, "headings.structural_elements", {})
        self._heading_common(b, cfg, "headings.struct")
        if cfg.get("case") == "upper" and b.text.strip() != b.text.strip().upper():
            runs = nk.text_runs(b)
            if not all(self.fonts.flag(r, b.obj, "caps") for r in runs):
                self.add("headings.struct.case", "Заголовки", "error", "auto",
                         "Заголовок структурного элемента не прописными буквами.", b.where())
        self._check_indent(b, 0, "headings.struct.indent", "Заголовки")
        if cfg.get("new_page"):
            self._check_new_page(i, b, "headings.struct.new_page", "Структурный элемент")
        bib = nk.normalize_heading(nk.prof(self.p, "bibliography.heading", ""))
        aliases = {nk.normalize_heading(a) for a in nk.prof(self.p, "bibliography.heading_aliases", [])}
        struct_alts = (nk.prof(self.p, "headings.structural_elements.aliases", {}) or {}).get(
            nk.prof(self.p, "bibliography.heading", ""), [])
        renamed_ok = nk.normalize_heading(b.text) in {nk.normalize_heading(a) for a in struct_alts}
        if nk.normalize_heading(b.text) in aliases and bib and not renamed_ok:
            self.add("bibliography.heading", "Ссылки и источники", "error", "claude",
                     f"Заголовок списка источников должен быть «{nk.prof(self.p, 'bibliography.heading')}».",
                     b.where())

    def _check_new_page(self, i, b, rule, what):
        if nk.starts_new_page(self.doc, self.blocks, i):
            return
        if nk.rendered_break_before(b):
            self.add(f"{rule}_soft", "Заголовки", "warning", "auto",
                     f"{what} начинается с новой страницы только за счёт пустых абзацев или переполнения "
                     "страницы — при правке текста съедет. Нужен разрыв страницы или «с новой страницы».",
                     b.where())
        else:
            self.add(rule, "Заголовки", "error", "auto", f"{what} начинается не с новой страницы.", b.where())

    def check_appendix_heading(self, i, b):
        cfg = nk.prof(self.p, "appendices", {})
        al = nk.eff_para(self.doc, b.obj, "alignment") or WD_ALIGN_PARAGRAPH.LEFT
        if cfg.get("alignment") and al != ALIGN[cfg["alignment"]]:
            self.add("appendix.alignment", "Заголовки", "error", "auto",
                     f"Заголовок приложения выровнен {ALIGN_RU.get(al, al)}, нужно {ALIGN_RU[ALIGN[cfg['alignment']]]}.",
                     b.where())
        if cfg.get("new_page"):
            self._check_new_page(i, b, "appendix.new_page", "Приложение")
        self.check_runs(b, "Заголовки")

    def check_section_heading(self, i, b):
        cfg = nk.prof(self.p, "headings.sections", {})
        self._heading_common(b, cfg, "headings.section")
        if b.extra.get("level") in cfg.get("new_page_levels", []):
            self._check_new_page(i, b, "headings.section.new_page", "Глава (раздел)")
        if "first_line_indent_cm" in cfg:
            self._check_indent(b, cfg["first_line_indent_cm"], "headings.section.indent", "Заголовки")
        if cfg.get("number_trailing_dot") is False and re.match(r"^\d+(\.\d+)*\.\s", b.text.strip()):
            self.add("headings.section.number_dot", "Заголовки", "error", "auto",
                     "Точка после номера раздела.", b.where())

    # ------------------------------------------------------------ рисунки и таблицы

    def check_caption(self, i, b, kind):
        cat = "Рисунки" if kind == "figure" else "Таблицы"
        cfg = nk.prof(self.p, "figures" if kind == "figure" else "tables", {})
        m = nk.match_caption(b.text, kind)
        label, sep, title = m.group("label"), m.group("sep"), m.group("title")
        sample = cfg.get("caption_format", "").replace("{n}", "1").replace("{title}", "Название")
        if label != nk.CANON_LABEL[kind]:
            self.add(f"{kind}.caption_format", cat, "error", "auto",
                     f"«{label}» вместо «{nk.CANON_LABEL[kind]}» (образец: «{sample}»).", b.where())
        if m.group("sp") != " ":
            self.add(f"{kind}.caption_format", cat, "error", "auto",
                     f"Нет пробела между словом и номером (образец: «{sample}»).", b.where())
        if sep != " — ":
            marks = sep.strip() or "ничего"
            self.add(f"{kind}.caption_format", cat, "error", "auto",
                     f"Между номером и названием «{marks}» вместо тире «—» (образец: «{sample}»).", b.where())
        if cfg.get("caption_trailing_dot") is False and title.rstrip().endswith(".") \
                and not re.search(r"\b\w{1,2}\.$", title.rstrip()):
            self.add(f"{kind}.caption_dot", cat, "error", "auto", "Точка в конце подписи.", b.where())
        if "caption_alignment" in cfg:
            al = nk.eff_para(self.doc, b.obj, "alignment") or WD_ALIGN_PARAGRAPH.LEFT
            need = ALIGN[cfg["caption_alignment"]]
            if al != need:
                self.add(f"{kind}.caption_alignment", cat, "error", "auto",
                         f"Подпись выровнена {ALIGN_RU.get(al, al)}, нужно {ALIGN_RU[need]}.", b.where())
        if "caption_first_line_indent_cm" in cfg:
            self._check_indent(b, cfg["caption_first_line_indent_cm"], f"{kind}.caption_indent", cat)
        if "caption_line_spacing" in cfg:
            if not caption_spacing_ok(self.doc, b, cfg["caption_line_spacing"]):
                self.add(f"{kind}.caption_spacing", cat, "error", "auto",
                         f"Межстрочный интервал подписи должен быть {cfg['caption_line_spacing']:g}.", b.where())
        self.check_runs(b, cat)
        # положение подписи
        if kind == "figure":
            before, after = nk.prev_block(self.blocks, i), nk.next_block(self.blocks, i)
            if not (before is not None and before.kind == "p" and before.role == "figure"):
                if after is not None and after.kind == "p" and after.role == "figure":
                    self.add("figure.caption_position", cat, "error", "claude",
                             "Подпись рисунка стоит над рисунком, нужна под ним.", b.where())
                else:
                    self.add("figure.caption_orphan", cat, "warning", "claude",
                             "Рядом с подписью не найден рисунок.", b.where())
        else:
            after, before = nk.next_block(self.blocks, i), nk.prev_block(self.blocks, i)
            if not (after is not None and after.kind == "tbl"):
                if before is not None and before.kind == "tbl":
                    self.add("table.caption_position", cat, "error", "claude",
                             "Подпись таблицы стоит под таблицей, нужна над ней.", b.where())
                else:
                    self.add("table.caption_orphan", cat, "warning", "claude",
                             "Сразу после подписи нет таблицы.", b.where())

    def check_figure(self, i, b):
        cfg = nk.prof(self.p, "figures", {})
        if cfg.get("image_alignment"):
            al = nk.eff_para(self.doc, b.obj, "alignment") or WD_ALIGN_PARAGRAPH.LEFT
            if al != ALIGN[cfg["image_alignment"]]:
                self.add("figure.alignment", "Рисунки", "error", "auto",
                         f"Рисунок выровнен {ALIGN_RU.get(al, al)}, нужно {ALIGN_RU[ALIGN[cfg['image_alignment']]]}.",
                         f"абз. {b.num} (рисунок)")
            ind = nk.eff_para(self.doc, b.obj, "first_line_indent")
            if ind is not None and abs(ind.cm) > 0.05:
                self.add("figure.indent", "Рисунки", "error", "auto",
                         f"У абзаца с рисунком отступ {ind.cm:.2f} см — рисунок сдвинут от центра.",
                         f"абз. {b.num} (рисунок)")
        after = nk.next_block(self.blocks, i)
        if not (after is not None and after.kind == "p" and after.role == "caption_figure"):
            before = nk.prev_block(self.blocks, i)
            if not (before is not None and before.kind == "p" and before.role == "caption_figure"):
                self.add("figure.no_caption", "Рисунки", "error", "author",
                         "У рисунка нет подписи «Рисунок N — Название».", f"абз. {b.num} (рисунок)")

    def check_tables_have_captions(self):
        for i, b in enumerate(self.blocks):
            if b.kind != "tbl" or b.role == "title":
                continue
            before = nk.prev_block(self.blocks, i)
            if before is not None and before.kind == "p" and before.role == "caption_table":
                continue
            if before is not None and before.kind == "p" and before.text.strip().lower().startswith("продолжение таблицы"):
                continue
            # подпись под таблицей уже отмечена в check_caption
            after = nk.next_block(self.blocks, i)
            if after is not None and after.kind == "p" and after.role == "caption_table":
                continue
            if is_layout_table(b):
                continue
            first = next((x for x in self.blocks[i + 1:] if x.kind == "p" and x.in_table and x.text.strip()), None)
            where = f"таблица, начинается с «{nk.excerpt(first.text, 40)}»" if first else "таблица"
            self.add("table.no_caption", "Таблицы", "error", "author",
                     "У таблицы нет подписи «Таблица N — Название».", where)

    # ------------------------------------------------------------ структура и ссылки

    def check_structure(self):
        present = {nk.normalize_heading(b.text) for b in self.blocks
                   if b.kind == "p" and b.role == "heading_struct"}
        aliases = nk.prof(self.p, "headings.structural_elements.aliases", {}) or {}
        for must in nk.prof(self.p, "headings.structural_elements.required", []):
            if nk.normalize_heading(must) in present:
                continue
            alt = next((a for a in aliases.get(must, []) if nk.normalize_heading(a) in present), None)
            if alt:
                self.add("structure.renamed", "Структура", "info", "author",
                         f"Элемент «{must}» назван «{alt}». Стандарт перечисляет элементы, но не требует "
                         "дословных заголовков; проверьте требования совета.")
            else:
                self.add("structure.missing", "Структура", "warning", "author",
                         f"Не найден структурный элемент «{must}» (или он назван иначе).")
        bib = nk.normalize_heading(nk.prof(self.p, "bibliography.heading", ""))
        aliases = {nk.normalize_heading(a) for a in nk.prof(self.p, "bibliography.heading_aliases", [])}
        bib_alts = {nk.normalize_heading(a) for a in nk.prof(self.p, "bibliography.heading_aliases", [])}
        if bib and not (present & ({bib} | bib_alts)):
            self.add("structure.no_bibliography", "Структура", "warning", "author",
                     "Не найден список использованных источников.")
        toc_heading = next((t for t in ("СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ") if t in present), None)
        if toc_heading and not any(b.kind == "p" and b.role == "toc" for b in self.blocks):
            self.add("structure.empty_toc", "Структура", "warning", "claude", f"После «{toc_heading}» нет перечня разделов.")
        has_auto_toc = any(b.kind == "p" and b.in_toc_sdt for b in self.blocks) or \
            any("TOC" in (t.text or "") for t in self.doc.element.body.iter(nk.qn("w:instrText")))
        if toc_heading and not has_auto_toc:
            self.add("structure.manual_toc", "Структура", "info", "claude",
                     f"«{toc_heading.capitalize()}» набрано вручную: номера страниц могут не совпасть. "
                     "Лучше собрать автоматически.")

    def _captions(self, role, kind):
        nums = []
        for b in self.blocks:
            if b.kind == "p" and b.role == role:
                nums.append((nk.match_caption(b.text, kind).group("num"), b))
        return nums

    def _is_continuation(self, b, n, kind) -> bool:
        word = "рисунк" if kind == "figure" else "таблиц"
        i = self.blocks.index(b)
        for x in self.blocks[max(0, i - 4):i]:
            if x.kind == "p" and re.match(rf"^(окончание|продолжение)\s+{word}\w*\s+{re.escape(n)}\b",
                                          x.text.strip(), re.I):
                return True
        return False

    def _merged_caption_numbers(self, kind) -> list[str]:
        """Подписи, слитые с соседним текстом через перенос строки, — номера тоже существуют."""
        out = []
        for b in self.blocks:
            if b.kind != "p" or b.role.startswith("caption") or "\n" not in b.text:
                continue
            for line in b.text.split("\n")[1:]:
                m = nk.match_caption(line, kind)
                if m:
                    out.append(m.group("num"))
        return out

    def check_merged_captions(self):
        for b in self.blocks:
            if b.kind != "p" or b.in_table or "\n" not in b.text:
                continue
            lines = b.text.split("\n")
            for kind, cat in (("figure", "Рисунки"), ("table", "Таблицы")):
                if any(nk.match_caption(line, kind) for line in lines[1:]):
                    self.add(f"{kind}.caption_merged", cat, "error", "claude",
                             "Подпись слита с соседним текстом через перенос строки (Shift+Enter), "
                             "а не стоит отдельным абзацем: её не видят автоматические списки и выравнивание.",
                             b.where())

    def check_cross_refs(self):
        body_text = "\n".join(b.text for b in self.blocks
                              if b.kind == "p" and b.role in ("body", "list", "table_text", "note"))
        for kind, role, cat, word, cfg_key in (
                ("figure", "caption_figure", "Рисунки", r"рис(?:\.|ун(?:ок|ка|ке|ком|ки|ках|ков|кам|ками))", "figures"),
                ("table", "caption_table", "Таблицы", r"табл(?:\.|иц(?:а|ы|е|у|ей|ах|ам|ами))", "tables")):
            caps = self._captions(role, kind)
            numbers = [n for n, _ in caps]
            seen = set()
            continued = set()
            for n, b in caps:
                if n in seen:
                    if self._is_continuation(b, n, kind):
                        continued.add(id(b))
                        if kind == "figure":
                            self.add("figure.continuation", cat, "warning", "claude",
                                     f"Рисунок {n} на нескольких листах: наименование повторено. По ГОСТ Р 2.105-2019, "
                                     f"п. 6.9.7 — «Рисунок {n} — Название (лист 1 из 2)», далее «Рисунок {n} (лист 2 из 2)»; "
                                     "слова «Окончание рисунка» не нужны.", b.where())
                    else:
                        self.add(f"{kind}.duplicate_number", cat, "error", "author",
                                 f"Номер {n} повторяется.", b.where())
                seen.add(n)
            numbers = [n for n, b in caps if id(b) not in continued]
            numbers += self._merged_caption_numbers(kind)
            plain = [int(n) for n in numbers if n.isdigit()]
            if plain and plain != list(range(1, len(plain) + 1)) and len(plain) == len(numbers):
                self.add(f"{kind}.numbering_order", cat, "warning", "author",
                         f"Нумерация идёт не по порядку: {', '.join(map(str, plain))}.")
            refs = set()
            for m in re.finditer(rf"\b{word}\s*((?:[А-ЯA-Z]\.)?\d+(?:\.\d+)*)", body_text, re.I):
                refs.add(m.group(1))
            for r in sorted(refs - set(numbers)):
                self.add(f"{kind}.ref_missing", cat, "error", "author",
                         f"В тексте есть ссылка на {'рисунок' if kind == 'figure' else 'таблицу'} {r}, "
                         "но такой подписи нет.")
            if nk.prof(self.p, f"{cfg_key}.references_required"):
                for n, b in caps:
                    if n not in refs:
                        self.add(f"{kind}.not_referenced", cat, "warning", "author",
                                 f"На {'рисунок' if kind == 'figure' else 'таблицу'} {n} нет ссылки в тексте.",
                                 b.where())

    def check_bibliography(self):
        entries = [b for b in self.blocks if b.kind == "p" and b.role == "bib_entry" and b.text.strip()]
        if not entries:
            return
        text = "\n".join(b.text for b in self.blocks if b.kind == "p" and b.role in (
            "body", "list", "table_text", "note", "caption_figure", "caption_table", "formula"))
        cited: list[int] = []
        cited = parse_citations(text)
        n = len(entries)
        fmt = nk.prof(self.p, "bibliography.entry_number")
        if fmt:
            need_dot = fmt.endswith(".")
            for b in entries:
                m = re.match(r"^\s*(\d+)(\.?)\s", b.text)
                if m and bool(m.group(2)) != need_dot:
                    self.add("bibliography.entry_number", "Ссылки и источники", "error", "auto",
                             f"Номер записи нужен в виде «{fmt.replace('{n}', m.group(1))}».", b.where())
        if nk.prof(self.p, "bibliography.citations_required") is False:
            cited = []
        if cited:
            over = sorted({c for c in cited if c > n or c < 1})
            if over:
                self.add("bibliography.ref_missing", "Ссылки и источники", "error", "author",
                         f"В тексте есть ссылки на источники {', '.join(map(str, over))}, а в списке "
                         f"{plural(n, 'запись', 'записи', 'записей')}.")
            unused = [i for i in range(1, n + 1) if i not in set(cited)]
            if unused:
                self.add("bibliography.unused", "Ссылки и источники", "warning", "author",
                         f"На источники {', '.join(map(str, unused))} нет ссылок в тексте.")
            if nk.prof(self.p, "bibliography.order") == "by_citation":
                first_seen = list(OrderedDict.fromkeys(c for c in cited if 1 <= c <= n))
                if first_seen != sorted(first_seen):
                    self.add("bibliography.order", "Ссылки и источники", "warning", "claude",
                             "Источники в списке расположены не в порядке первого упоминания в тексте.")
        elif nk.prof(self.p, "bibliography.citations_required") is not False:
            self.add("bibliography.no_refs", "Ссылки и источники", "warning", "author",
                     "В тексте не найдено ссылок вида [1] на источники.")
        self.add("bibliography.review", "Ссылки и источники", "info", "claude",
                 f"Записей в списке: {n}. Описания сверить с {nk.prof(self.p, 'bibliography.standard')} "
                 "(см. references/bibliography.md).")


# ---------------------------------------------------------------- отчёт


def standards_lines(profile: dict) -> list[str]:
    """Строки «по каким стандартам проверяли» со статусом из каталога."""
    try:
        import yaml

        cat = yaml.safe_load((nk.PROFILES_DIR / "catalog.yaml").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    out = []
    for sid in profile.get("standards", []):
        st = cat.get("standards", {}).get(sid)
        if st:
            out.append(f"- {st['title']} — {' '.join(st['status'].split())} (проверено {st['checked']}).")
    if profile.get("readiness") in ("beta", "limited"):
        out.append("- Профиль в статусе «" + {"beta": "бета", "limited": "частично"}[profile["readiness"]]
                   + "»: часть требований проверяется только вручную.")
    return (["Стандарты:"] + out + [""]) if out else []


def plural(n: int, one: str, few: str, many: str) -> str:
    """plural(5, "ошибка", "ошибки", "ошибок") -> "5 ошибок"."""
    if n % 10 == 1 and n % 100 != 11:
        word = one
    elif 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        word = few
    else:
        word = many
    return f"{n} {word}"


def group(findings: list[Finding]):
    """Склеивает одинаковые нарушения: одна строка — список мест."""
    out: "OrderedDict[tuple, list]" = OrderedDict()
    for f in findings:
        out.setdefault((f.category, f.severity, f.fix, f.rule, f.message), []).append(f.where)
    return out


def render_markdown(findings: list[Finding], path: str, profile: dict, max_places: int = 8) -> str:
    grouped = group(findings)
    errors = sum(1 for f in findings if f.severity == "error")
    warnings = sum(1 for f in findings if f.severity == "warning")
    by_fix = {k: len({key for key in grouped if key[2] == k}) for k in FIXERS}
    lines = [f"# Нормоконтроль: {path}", "",
             f"Профиль: {profile.get('name', '?')}.",
             *standards_lines(profile),
             f"Найдено: {plural(errors, 'ошибка', 'ошибки', 'ошибок')}, "
             f"{plural(warnings, 'предупреждение', 'предупреждения', 'предупреждений')} "
             f"— {plural(len(grouped), 'пункт', 'пункта', 'пунктов')} в отчёте. "
             f"Исправит скрипт — {by_fix['auto']}, смысловая правка — {by_fix['claude']}, "
             f"на автора — {by_fix['author']}.", ""]
    for cat in CATEGORIES:
        items = [(k, v) for k, v in grouped.items() if k[0] == cat]
        if not items:
            continue
        lines += [f"## {cat}", ""]
        for (_, sev, fix, _, msg), places in items:
            places = [p for p in places if p]
            tail = ""
            if places:
                shown = "; ".join(places[:max_places])
                more = f" и ещё {len(places) - max_places}" if len(places) > max_places else ""
                tail = f" — {shown}{more}"
            lines.append(f"- [{FIXERS[fix]}, {SEVERITY[sev]}] {msg}{tail}")
        lines.append("")
    if not findings:
        lines.append("Нарушений не найдено.")
    return "\n".join(lines).rstrip() + "\n"


def audit_file(path, profile_name="gost-7.32-2017"):
    profile = nk.load_profile(profile_name)
    doc = nk.open_docx(path)
    return Audit(doc, profile).run(), profile


def main(argv=None):
    ap = argparse.ArgumentParser(description="Проверка оформления .docx по профилю")
    ap.add_argument("docx")
    ap.add_argument("--profile", default="gost-7.32-2017", help="id профиля или путь к .yaml")
    ap.add_argument("--json", help="сохранить нарушения в JSON")
    ap.add_argument("--out", help="сохранить отчёт Markdown в файл")
    ap.add_argument("--max-places", type=int, default=8, help="сколько мест показывать на строку")
    args = ap.parse_args(argv)

    findings, profile = audit_file(args.docx, args.profile)
    report = render_markdown(findings, args.docx, profile, args.max_places)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(report)
    sys.stdout.write(report)
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump([asdict(f) for f in findings], fh, ensure_ascii=False, indent=2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
