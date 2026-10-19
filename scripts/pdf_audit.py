#!/usr/bin/env python3
"""Проверка оформления по готовому PDF. Документ не меняется.

    python3 scripts/pdf_audit.py работа.pdf [--profile gost-r-7.0.11-2011-dissertation] [--json out.json]

PDF — это уже свёрстанные страницы: стилей, разрывов и отступов в нём нет. Поэтому
проверяется то, что видно на странице: формат, поля (по границам текста), кегль и
интервал (по символам), положение номера страницы, структура, подписи, ссылки,
список литературы. Исправлять PDF инструмент не умеет — для правок нужен .docx.

Измерения приблизительные: поля оцениваются по крайним символам на типичных
страницах, интервал — по расстоянию между строками. Допуски заданы ниже.
"""

from __future__ import annotations

import argparse
import json
import re
import statistics
import subprocess
import sys
from collections import Counter
from dataclasses import asdict, dataclass, field

import nk_common as nk
from audit import Finding, parse_citations, render_markdown

PT_PER_MM = 72 / 25.4
TOL_SIDE_MM = 2.5      # левое и правое поле: текст по ширине доходит до края
TOL_VERT_MM = 4.0      # верх и низ: Word добавляет интервал над первой строкой
SAMPLE_PAGES = 30      # сколько страниц мерить (равномерно по документу)
LINE_HEIGHT = 1.15     # высота строки Times New Roman в долях кегля (одинарный интервал Word)


@dataclass
class Line:
    text: str
    x0: float
    x1: float
    top: float
    bottom: float
    size: float
    font: str


@dataclass
class Page:
    num: int
    width: float
    height: float
    text: str = ""
    lines: list[Line] = field(default_factory=list)


# ---------------------------------------------------------------- чтение PDF


def page_texts(path: str) -> list[str]:
    # -layout: строки по положению на странице; без него текст по ширине теряет порядок слов
    out = subprocess.run(["pdftotext", "-layout", "-enc", "UTF-8", path, "-"], capture_output=True, check=True).stdout
    pages = out.decode("utf-8", "replace").split("\f")
    pages = pages[:-1] if pages and not pages[-1].strip() else pages
    # пробелы выравнивания внутри строки схлопываем: «ВВЕДЕНИЕ      5» → «ВВЕДЕНИЕ 5»
    return ["\n".join(re.sub(r"\s{2,}", " ", ln).strip() for ln in pg.splitlines()) for pg in pages]


def measure_pages(path: str, nums: list[int]) -> dict[int, Page]:
    import pdfplumber

    result = {}
    with pdfplumber.open(path) as pdf:
        for n in nums:
            p = pdf.pages[n - 1]
            page = Page(n, float(p.width), float(p.height))
            words = p.extract_words(extra_attrs=["fontname", "size"], keep_blank_chars=False, use_text_flow=True)
            rows: dict[int, list] = {}
            for w in words:
                rows.setdefault(round(w["top"] / 2), []).append(w)
            for key in sorted(rows):
                ws = sorted(rows[key], key=lambda w: w["x0"])
                sizes = Counter(round(w["size"], 1) for w in ws)
                fonts = Counter(re.sub(r"^[A-Z]{6}\+", "", w["fontname"]) for w in ws)
                page.lines.append(Line(" ".join(w["text"] for w in ws), min(w["x0"] for w in ws),
                                       max(w["x1"] for w in ws), min(w["top"] for w in ws),
                                       max(w["bottom"] for w in ws), sizes.most_common(1)[0][0],
                                       fonts.most_common(1)[0][0]))
            result[n] = page
    return result


# ---------------------------------------------------------------- проверка


class PdfAudit:
    def __init__(self, path: str, profile: dict):
        self.path = path
        self.p = profile
        self.texts = page_texts(path)
        self.n_pages = len(self.texts)
        self.findings: list[Finding] = []
        self.stats: dict = {"pages": self.n_pages}

    def add(self, rule, category, severity, fix, message, where=None):
        self.findings.append(Finding(rule, category, severity, fix, message, where))

    def run(self):
        sample = self._sample_pages()
        self.measured = measure_pages(self.path, sorted(set(sample + [1, 2])))
        self.check_page_size()
        self.check_margins_and_text(sample)
        self.check_page_numbers()
        self.check_structure()
        self.check_captions()
        self.check_bibliography()
        return self.findings

    def _sample_pages(self) -> list[int]:
        first, last = min(4, self.n_pages), self.n_pages
        if last - first < SAMPLE_PAGES:
            return list(range(first, last + 1))
        step = (last - first) / SAMPLE_PAGES
        return sorted({int(first + i * step) for i in range(SAMPLE_PAGES)})

    # --- страница

    def check_page_size(self):
        sizes = Counter((round(p.width / PT_PER_MM), round(p.height / PT_PER_MM)) for p in self.measured.values())
        (w, h), _ = sizes.most_common(1)[0]
        self.stats["page_mm"] = f"{w}×{h}"
        if nk.prof(self.p, "page.size") == "A4" and sorted((w, h)) != [210, 297] and \
                not (abs(min(w, h) - 210) <= 1 and abs(max(w, h) - 297) <= 1):
            self.add("page.size", "Страница", "error", "author",
                     f"Формат страницы {w}×{h} мм, нужен A4 (210×297 мм).")

    def _body_lines(self, page: Page) -> list[Line]:
        """Строки основного текста: без номера страницы и колонтитулов."""
        lines = [ln for ln in page.lines if not re.fullmatch(r"\s*\d{1,3}\s*", ln.text)]
        return [ln for ln in lines if len(ln.text) > 3]

    def check_margins_and_text(self, sample):
        lefts, rights, tops, bottoms, sizes, fonts, pitches = [], [], [], [], Counter(), Counter(), []
        for n in sample:
            page = self.measured[n]
            if page.width > page.height:
                continue  # альбомные листы меряем отдельно — пропускаем
            body = self._body_lines(page)
            if len(body) < 15:
                continue  # страница с рисунком или таблицей
            lefts.append(min(ln.x0 for ln in body))
            rights.append(page.width - max(ln.x1 for ln in body))
            tops.append(min(ln.top for ln in body))
            bottoms.append(page.height - max(ln.bottom for ln in body))
            for ln in body:
                sizes[ln.size] += len(ln.text)
                fonts[ln.font] += len(ln.text)
            main = sizes.most_common(1)[0][0]
            seq = sorted(ln.top for ln in body if abs(ln.size - main) < 0.3)
            pitches += [b - a for a, b in zip(seq, seq[1:]) if 0 < b - a < main * 3]
        if not lefts:
            self.add("pdf.no_text_pages", "Страница", "info", "claude",
                     "Не нашлось страниц со сплошным текстом для измерений.")
            return
        mm = lambda v: statistics.median(v) / PT_PER_MM  # noqa: E731
        # верх и низ — по самым заполненным страницам: на неполных текст обрывается раньше
        low = lambda v: sorted(v)[max(0, len(v) // 5 - 1)] / PT_PER_MM  # noqa: E731
        have = {"left": mm(lefts), "right": mm(rights), "top": low(tops), "bottom": low(bottoms)}
        self.stats["margins_mm"] = {k: round(v, 1) for k, v in have.items()}
        names = {"left": "левое", "right": "правое", "top": "верхнее", "bottom": "нижнее"}
        for side, need in (nk.prof(self.p, "page.margins_mm", {}) or {}).items():
            tol = TOL_SIDE_MM if side in ("left", "right") else TOL_VERT_MM
            # сверху и снизу текст начинается чуть дальше поля из-за интервала — меньше поля быть не может
            delta = have[side] - need
            if side in ("top", "bottom") and 0 <= delta <= tol * 2:
                continue
            if delta > tol:
                self.add("page.margins_wider", "Страница", "info", "author",
                         f"{names[side].capitalize()} поле около {have[side]:.0f} мм — шире нормы ({need} мм). "
                         "Текст от этого не страдает; советы такое обычно принимают.")
            elif delta < -tol:
                self.add("page.margins", "Страница", "error", "author",
                         f"{names[side].capitalize()} поле около {have[side]:.0f} мм (по границе текста), "
                         f"нужно {need} мм.")
        size = sizes.most_common(1)[0][0]
        font = fonts.most_common(1)[0][0]
        self.stats["font"] = font
        self.stats["size_pt"] = size
        t = self.p.get("text", {})
        if t.get("font") and t["font"].replace(" ", "").lower() not in font.replace(" ", "").replace("-", "").lower() \
                and not (t["font"] == "Times New Roman" and "times" in font.lower()):
            self.add("text.font", "Основной текст", "error", "author", f"Основной шрифт {font}, нужен {t['font']}.")
        if t.get("size_pt_min") and size < t["size_pt_min"] - 0.3:
            self.add("text.size", "Основной текст", "error", "author",
                     f"Основной кегль {size:g} пт, нужно не меньше {t['size_pt_min']} пт.")
        if t.get("size_pt_max") and size > t["size_pt_max"] + 0.3:
            self.add("text.size_max", "Основной текст", "error", "author",
                     f"Основной кегль {size:g} пт, нужно не больше {t['size_pt_max']} пт.")
        if pitches and t.get("line_spacing"):
            spacing = statistics.median(pitches) / (size * LINE_HEIGHT)
            self.stats["line_spacing"] = round(spacing, 2)
            if abs(spacing - t["line_spacing"]) > 0.15:
                self.add("text.line_spacing", "Основной текст", "error", "author",
                         f"Межстрочный интервал около {spacing:.2f}, нужен {t['line_spacing']}.")

    # --- номера страниц

    def _page_number_position(self, page: Page) -> str | None:
        for ln in page.lines:
            if not re.fullmatch(r"\s*\d{1,3}\s*", ln.text):
                continue
            cx = (ln.x0 + ln.x1) / 2
            vert = "top" if ln.top < page.height * 0.15 else ("bottom" if ln.bottom > page.height * 0.85 else None)
            if vert is None:
                continue
            horiz = "center" if abs(cx - page.width / 2) < page.width * 0.08 else (
                "right" if cx > page.width / 2 else "left")
            return f"{vert}_{horiz}"
        return None

    def check_page_numbers(self):
        cfg = nk.prof(self.p, "page_numbers")
        if not cfg:
            return
        positions = Counter(self._page_number_position(p) for n, p in self.measured.items() if n > 2)
        pos, count = positions.most_common(1)[0]
        self.stats["page_number"] = pos
        need = cfg.get("position")
        ru = {"top_center": "вверху по центру", "bottom_center": "внизу по центру", "top_right": "вверху справа",
              "bottom_right": "внизу справа", "top_left": "вверху слева", "bottom_left": "внизу слева"}
        if pos is None:
            self.add("page_numbers.missing", "Нумерация страниц", "error", "author", "Номера страниц не найдены.")
        elif need and pos != need:
            self.add("page_numbers.position", "Нумерация страниц", "warning", "author",
                     f"Номер страницы {ru.get(pos, pos)}, по стандарту — {ru.get(need, need)}. "
                     "На практике советы часто принимают другое положение — сверьтесь с их требованиями.")
        if cfg.get("title_page_shown") is False and self._page_number_position(self.measured[1]):
            self.add("page_numbers.title_page", "Нумерация страниц", "error", "author",
                     "На титульном листе стоит номер страницы.")

    # --- структура

    def _first_text_line(self, n: int) -> str:
        for line in self.texts[n - 1].splitlines():
            s = line.strip()
            if s and not re.fullmatch(r"\d{1,3}", s):
                return s
        return ""

    def _find_heading(self, names: set[str], start: int = 1):
        """(страница, текст) первых строк, совпадающих с названием элемента."""
        found = []
        for n in range(start, self.n_pages + 1):
            for line in self.texts[n - 1].splitlines():
                raw = line.strip()
                if raw[:1].isupper() and nk.normalize_heading(raw) in names:
                    found.append((n, raw))
        return found

    def _find_variant(self, required: str, start: int):
        """Заголовок-вариант: «ОБЩИЕ ВЫВОДЫ И ЗАКЛЮЧЕНИЕ» для «ЗАКЛЮЧЕНИЕ»."""
        for n in range(start, self.n_pages + 1):
            for line in self.texts[n - 1].splitlines():
                if not re.search(r"\.{4,}|\s\d{1,3}$", line) and nk.heading_variant(line, required):
                    return n, line.strip()
        return None

    def _language(self) -> str:
        sample = "".join(self.texts[2:min(self.n_pages, 30)])
        cyr = len(re.findall(r"[А-Яа-яЁё]", sample))
        lat = len(re.findall(r"[A-Za-z]", sample))
        return "ru" if cyr >= lat else "foreign"

    def check_structure(self):
        cfg = nk.prof(self.p, "headings.structural_elements", {}) or {}
        self.stats["language"] = self._language()
        if self.stats["language"] != "ru":
            self.toc_end = 2
            self.add("structure.foreign_language", "Структура", "info", "author",
                     "Работа написана не на русском языке: названия структурных элементов не проверялись. "
                     "По ГОСТ Р 7.0.11-2011, п. 5.1.3, к ней прилагают дополнительный титульный лист на русском.")
            return
        aliases = cfg.get("aliases", {}) or {}
        toc_page = None
        for t in ("ОГЛАВЛЕНИЕ", "СОДЕРЖАНИЕ"):
            hits = self._find_heading({t})
            if hits:
                toc_page = hits[0][0]
                self.stats["toc"] = t
                break
        self.toc_end = (toc_page or 1) + self._toc_length(toc_page)
        bib = nk.prof(self.p, "bibliography.heading", "")
        for must in cfg.get("required", []):
            names = {nk.normalize_heading(must)} | {nk.normalize_heading(a) for a in aliases.get(must, [])}
            if must == bib:
                names |= self._bib_names()
            start = 1 if must in ("ОГЛАВЛЕНИЕ", "СОДЕРЖАНИЕ") else self.toc_end + 1
            hits = self._find_heading(names, start)
            if not hits and must in ("ОГЛАВЛЕНИЕ", "СОДЕРЖАНИЕ"):
                other = "СОДЕРЖАНИЕ" if must == "ОГЛАВЛЕНИЕ" else "ОГЛАВЛЕНИЕ"
                if self._find_heading({other}):
                    self.add("structure.renamed", "Структура", "warning", "author",
                             f"Перечень разделов назван «{other}», по стандарту — «{must}».")
                    continue
            if not hits:
                variant = self._find_variant(must, start)
                if variant:
                    self.add("structure.renamed", "Структура", "info", "author",
                             f"Элемент «{must}» назван «{variant[1]}».", f"с. {variant[0]}")
                    hits = [variant]
                else:
                    self.add("structure.missing", "Структура", "warning", "author",
                             f"Не найден структурный элемент «{must}».")
                    continue
            n, text = hits[0]
            if nk.normalize_heading(text) != nk.normalize_heading(must) and not any(
                    f.rule == "structure.renamed" and f"«{must}»" in f.message for f in self.findings):
                self.add("structure.renamed", "Структура", "info", "author",
                         f"Элемент «{must}» назван «{text}».", f"с. {n}")
            if re.match(r"^\s*(\d{1,2}|[IVX]+)\.\s*", text) and must in ("ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ"):
                self.add("headings.struct.numbered", "Заголовки", "info", "author",
                         f"«{text}» пронумерован как глава; обычно введение и заключение не нумеруют.", f"с. {n}")
            if cfg.get("new_page") and nk.normalize_heading(self._first_text_line(n)) != nk.normalize_heading(text):
                self.add("headings.struct.new_page", "Заголовки", "error", "author",
                         f"«{text}» начинается не с новой страницы.", f"с. {n}")
        sec = nk.prof(self.p, "headings.sections", {}) or {}
        if 1 in (sec.get("new_page_levels") or []):
            for n in range(self.toc_end + 1, self.n_pages + 1):
                for line in self.texts[n - 1].splitlines():
                    if re.match(r"^\s*(ГЛАВА|Глава)\s+\d+\.?\s*[.:—–-]?\s*[А-ЯЁA-Z«]", line) and \
                            not re.search(r"\.{4,}|\s\d{1,3}$", line):
                        if not re.match(r"^\s*(ГЛАВА|Глава)\s+\d+", self._first_text_line(n)):
                            self.add("headings.section.new_page", "Заголовки", "error", "author",
                                     f"«{nk.excerpt(line.strip(), 50)}» начинается не с новой страницы.", f"с. {n}")

    def _toc_length(self, toc_page) -> int:
        """Сколько страниц после первой занимает оглавление (строки с номерами страниц в конце)."""
        if not toc_page:
            return 0
        extra = 0
        for n in range(toc_page + 1, min(toc_page + 5, self.n_pages + 1)):
            lines = [x for x in self.texts[n - 1].splitlines() if x.strip()]
            numbered = sum(bool(re.search(r"\.{3,}\s*\d+\s*$|\s\d{1,3}\s*$|^\d{1,3}$", x)) for x in lines)
            if lines and numbered > len(lines) * 0.3:
                extra += 1
            else:
                break
        return extra

    # --- подписи

    def _caption_lines(self, kind: str):
        out = []
        for n in range(self.toc_end + 1, self.n_pages + 1):
            for line in self.texts[n - 1].splitlines():
                m = nk.match_caption(line, kind)
                if m and not re.search(r"\.{4,}", line):  # строки оглавления и списков иллюстраций — не подписи
                    out.append((n, line.strip(), m))
        return out

    @property
    def appendix_page(self):
        word = (nk.prof(self.p, "appendices.heading_word") or "ПРИЛОЖЕНИЕ").upper()
        bib = self._bib_page() or self.toc_end + 1
        for n in range(bib, self.n_pages + 1):
            if re.search(rf"(?m)^\s*{word}\s+[А-ЯA-Z1-9]\b", self.texts[n - 1].upper()):
                return n
        return None

    def check_captions(self):
        body = "\n".join(self.texts[self.toc_end:self._bib_page() or self.n_pages])
        for kind, cat, word in (("figure", "Рисунки", r"рис(?:\.|ун(?:ок|ка|ке|ком|ки|ках|ков|кам|ками))"),
                                ("table", "Таблицы", r"табл(?:\.|иц(?:а|ы|е|у|ей|ах|ам|ами))")):
            caps = self._caption_lines(kind)
            if not caps:
                continue
            seps = Counter(m.group("sep").strip() or "пробел" for _, _, m in caps)
            labels = Counter(m.group("label") for _, _, m in caps)
            self.stats[f"{kind}_captions"] = {"count": len(caps), "sep": dict(seps), "label": dict(labels)}
            en = [(n, line) for n, line, m in caps if m.group("sep").strip() == "–"]
            bad_sep = [(n, line) for n, line, m in caps if m.group("sep").strip() not in ("—", "–")]
            if en:
                self.add(f"{kind}.caption_en_dash", cat, "info", "claude",
                         f"В {len(en)} подписях короткое тире «–» вместо «—». Типографская мелочь; правится автозаменой.")
            if bad_sep:
                marks = ", ".join(f"«{k}» — {v}" for k, v in seps.items() if k not in ("—", "–"))
                self.add(f"{kind}.caption_format", cat, "warning", "claude",
                         f"Между номером и названием не тире: {marks} подписей (образец: "
                         f"«{nk.CANON_LABEL[kind]} 1 — Название»).",
                         "; ".join(f"с. {n}" for n, _ in bad_sep[:5]))
            bad_label = [(n, line) for n, line, m in caps if m.group("label") != nk.CANON_LABEL[kind]]
            if bad_label:
                self.add(f"{kind}.caption_label", cat, "warning", "claude",
                         f"Сокращение «{bad_label[0][1].split()[0]}» вместо «{nk.CANON_LABEL[kind]}».",
                         "; ".join(f"с. {n}" for n, _ in bad_label[:5]))
            numbers, seen = [], set()
            for n, line, m in caps:
                num = m.group("num")
                cont = re.search(r"\((лист|окончание|продолжение)", line, re.I)
                if num in seen and not cont:
                    if self.appendix_page and n >= self.appendix_page and num.isdigit():
                        self.add(f"{kind}.appendix_numbering", cat, "warning", "claude",
                                 f"В приложениях {'рисунки' if kind == 'figure' else 'таблицы'} нумеруются заново без "
                                 "буквы приложения; нужно «А.1», «Б.1» (ГОСТ Р 2.105-2019, п. 6.8.2, 6.9.3).",
                                 f"с. {n}")
                    else:
                        self.add(f"{kind}.duplicate_number", cat, "warning", "author",
                                 f"Номер {num} встречается в подписях дважды (повтор подписи или перенос без "
                                 "пометки «лист 2 из 2»).", f"с. {n}")
                seen.add(num)
                numbers.append(num)
            refs = nk.expand_refs(body, word)
            caption_texts = {line for _, line, _ in caps}
            missing = sorted(refs - set(numbers), key=lambda x: [int(p) if p.isdigit() else 0 for p in x.split(".")])
            if missing:
                # в PDF подпись может быть частью картинки — поэтому предупреждение, а не ошибка
                self.add(f"{kind}.ref_missing", cat, "warning", "author",
                         f"В тексте есть ссылки на {'рисунки' if kind == 'figure' else 'таблицы'} "
                         f"{', '.join(missing[:12])}, а подписей с такими номерами в тексте PDF нет "
                         "(проверьте: подпись могла остаться внутри картинки).")
            # «На таблицу N нет ссылки»: ищем упоминания вне самой подписи
            body_wo_caps = "\n".join(x for x in body.splitlines() if x.strip() not in caption_texts)
            mentioned = nk.expand_refs(body_wo_caps, word)
            for num in dict.fromkeys(numbers):
                if num not in mentioned and not re.match(r"[А-ЯA-Z]", num):  # в приложениях — отдельно
                    self.add(f"{kind}.not_referenced", cat, "warning", "author",
                             f"На {'рисунок' if kind == 'figure' else 'таблицу'} {num} нет ссылки в тексте.")

    # --- список литературы

    def _bib_page(self):
        names = {nk.normalize_heading(nk.prof(self.p, "bibliography.heading", ""))}
        names |= {nk.normalize_heading(a) for a in nk.prof(self.p, "bibliography.heading_aliases", [])}
        hits = self._find_heading(names, self.toc_end + 1 if hasattr(self, "toc_end") else 1)
        return hits[-1][0] if hits else None

    def check_bibliography(self):
        start = self._bib_page()
        if not start:
            return
        word = (nk.prof(self.p, "appendices.heading_word") or "ПРИЛОЖЕНИЕ").upper()
        lines, collecting = [], False
        for n in range(start, self.n_pages + 1):
            for line in self.texts[n - 1].splitlines():
                t = line.strip()
                if not collecting:
                    if nk.normalize_heading(t) in self._bib_names():
                        collecting = True
                    continue
                if re.match(rf"^{word}\s+[А-ЯA-Z]\b", t.upper()):
                    collecting = False
                    break
                if t and not re.fullmatch(r"\d{1,3}", t):
                    lines.append(t)
            if lines and not collecting:
                break
        entries = self._split_entries(lines)
        self.stats["bib_entries"] = len(entries)
        if not entries:
            return
        body = "\n".join(self.texts[self.toc_end:start - 1])
        cited = parse_citations(body)
        n = len(entries)
        author_year = len(re.findall(r"\([А-ЯЁA-Z][\w-]+(?: et al\.| и др\.)?,? (?:19|20)\d{2}[a-zа-я]?\)", body))
        self.stats["citations"] = {"numeric": len(cited), "author_year": author_year}
        if author_year > len(cited):
            self.add("bibliography.author_year", "Ссылки и источники", "info", "author",
                     "Ссылки в формате «(Автор, год)»; сверка номеров не проводилась. ГОСТ Р 7.0.5-2008 "
                     "допускает такие отсылки, если записи в списке не нумерованы (п. 7.5).")
            cited = []
        if nk.prof(self.p, "bibliography.citations_required") is not False and cited:
            typos = sorted({c for c in cited if c > 2 * n + 10})
            over = sorted({c for c in cited if n < c <= 2 * n + 10})
            if over:
                self.add("bibliography.ref_missing", "Ссылки и источники", "error", "author",
                         f"Ссылки на источники {', '.join(map(str, over[:15]))}, а в списке {n} записей.")
            if typos:
                self.add("bibliography.ref_typo", "Ссылки и источники", "warning", "author",
                         f"Подозрительные номера в ссылках: {', '.join(map(str, typos[:10]))} — вероятно, пропущен "
                         "разделитель («[122125]» вместо «[122, 125]»).")
            unused = [i for i in range(1, n + 1) if i not in set(cited)]
            self.stats["bib_uncited"] = len(unused)
            if unused:
                footnotes = len(re.findall(r"(?m)^\s*\d{1,3}\s+[А-ЯЁA-Z][\w-]+,? [А-ЯA-Z]\.", body))
                caveat = (" В работе есть постраничные сноски с описаниями источников — часть «неиспользованных» "
                          "может быть процитирована в них.") if footnotes >= 5 else ""
                self.add("bibliography.unused", "Ссылки и источники", "warning", "author",
                         f"На {len(unused)} из {n} источников нет ссылок вида [N] в тексте: "
                         f"{', '.join(map(str, unused[:20]))}{'…' if len(unused) > 20 else ''}.{caveat}")
        nodot = [i + 1 for i, e in enumerate(entries) if not re.search(r"[.)]\s*$", e)]
        if nodot and len(nodot) < n * 0.5:
            self.add("bibliography.no_final_dot", "Ссылки и источники", "warning", "claude",
                     f"Без точки в конце: записи {', '.join(map(str, nodot[:20]))}.")
        if nk.prof(self.p, "bibliography.order") in ("any", "alphabetical"):
            self._check_alphabet(entries)

    def _bib_names(self):
        names = {nk.normalize_heading(nk.prof(self.p, "bibliography.heading", ""))}
        return names | {nk.normalize_heading(a) for a in nk.prof(self.p, "bibliography.heading_aliases", [])}

    @staticmethod
    def _split_entries(lines: list[str]) -> list[str]:
        """Записи списка по номерам «1.», «2.» в начале строки; без номеров — пусто."""
        entries, expected = [], 1
        for line in lines:
            m = re.match(rf"^{expected}\s*[.)]?(\s+\S|\s*$)", line)
            if m and (m.group(1).strip() or line.strip() in (f"{expected}.", f"{expected})")):
                entries.append(line[m.end() - 1:] if m.group(1).strip() else "")
                expected += 1
            elif entries:
                entries[-1] += " " + line
        return entries if len(entries) >= 5 else []

    def _check_alphabet(self, entries):
        def key(t):
            return re.sub(r"^[«\"„“'(\[]+", "", t).lower().replace("ё", "е")

        def author(t):
            return re.split(r"[,.]", t)[0]

        def is_latin(t):
            return bool(re.match(r"[A-Za-z]", key(t)))

        # список упорядочен по алфавиту, если большинство соседних пар идёт по порядку
        pairs = [(a, b) for a, b in zip(entries, entries[1:]) if is_latin(a) == is_latin(b)]
        if not pairs:
            return
        ordered = sum(key(a) <= key(b) for a, b in pairs) / len(pairs)
        self.stats["bib_alpha_share"] = round(ordered, 2)
        if ordered < 0.8:
            return  # список не алфавитный (по ссылкам или иной) — алфавит не проверяем
        bad = [i + 1 for i, (a, b) in enumerate(zip(entries, entries[1:]))
               if is_latin(a) == is_latin(b) and author(a) != author(b) and key(a) > key(b)]
        if bad:
            self.add("bibliography.alphabet", "Ссылки и источники", "warning", "claude",
                     f"Список алфавитный, но порядок нарушен после записей {', '.join(map(str, bad[:15]))}.")
        latin_first = [i for i, e in enumerate(entries) if is_latin(e)]
        cyr_after = [i for i, e in enumerate(entries) if not is_latin(e) and latin_first and i > latin_first[0]]
        if latin_first and cyr_after and len(cyr_after) > 2:
            self.add("bibliography.languages", "Ссылки и источники", "warning", "claude",
                     "Иноязычные записи перемешаны с русскими; по ГОСТ Р 7.0.11, п. 5.6.7 — отдельный ряд после русских.")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Проверка оформления по PDF")
    ap.add_argument("pdf")
    ap.add_argument("--profile", default="gost-r-7.0.11-2011-dissertation")
    ap.add_argument("--json", help="сохранить нарушения и измерения в JSON")
    args = ap.parse_args(argv)
    profile = nk.load_profile(args.profile)
    a = PdfAudit(args.pdf, profile)
    findings = a.run()
    report = render_markdown(findings, args.pdf, profile)
    stats = "\n".join(f"- {k}: {v}" for k, v in a.stats.items())
    sys.stdout.write(report + "\n## Измерения (по PDF, приблизительно)\n\n" + stats + "\n")
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({"findings": [asdict(f) for f in findings], "stats": a.stats}, fh, ensure_ascii=False, indent=2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
