"""Общие функции: профиль, обход документа, эффективное оформление, роли абзацев."""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from pathlib import Path

import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.shared import Emu, Pt, Twips
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent.parent
PROFILES_DIR = ROOT / "profiles"

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


# ---------------------------------------------------------------- профиль


def _merge(base: dict, over: dict) -> dict:
    out = copy.deepcopy(base)
    for key, value in over.items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = _merge(out[key], value)
        else:
            out[key] = copy.deepcopy(value)
    return out


def _find_profile(name: str, near: Path | None = None) -> Path:
    candidates = [Path(name)]
    if near is not None:
        candidates.append(near / f"{name}.yaml")
    candidates.append(PROFILES_DIR / f"{name}.yaml")
    for path in candidates:
        if path.is_file():
            return path
    raise FileNotFoundError(f"Профиль не найден: {name}")


def _drop_none(node):
    if isinstance(node, dict):
        return {k: _drop_none(v) for k, v in node.items() if v is not None}
    return node


def _load_raw(name: str, near: Path | None = None) -> dict:
    path = _find_profile(name, near)
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    parent = data.pop("extends", None)
    if parent:
        data = _merge(_load_raw(parent, path.parent), data)
    return data


def load_profile(name: str, near: Path | None = None) -> dict:
    """Загружает профиль и применяет наследование через `extends`.

    Значение `null` означает «правило не задано»: так профиль-наследник
    отменяет требование базового, которого нет в его стандарте.
    """
    return _drop_none(_load_raw(name, near))


def indent_range(value) -> tuple[float, float]:
    """Отступ в профиле — число (1.25) или диапазон ([1.25, 1.7]); первое значение — целевое."""
    if isinstance(value, (list, tuple)):
        return float(value[0]), float(value[-1])
    return float(value), float(value)


def prof(profile: dict, dotted: str, default=None):
    node = profile
    for part in dotted.split("."):
        if not isinstance(node, dict) or part not in node:
            return default
        node = node[part]
    return node


# ---------------------------------------------------------------- текст абзаца

_SKIP_ANCESTORS = {qn("w:del"), qn("w:moveFrom"), qn("w:txbxContent")}


def visible_runs(p_el) -> list:
    """Видимые прогоны абзаца по порядку: без удалённых и без надписей."""
    out = []
    for r in p_el.iter(qn("w:r")):
        anc = r.getparent()
        skip = False
        while anc is not None and anc is not p_el:
            if anc.tag in _SKIP_ANCESTORS:
                skip = True
                break
            anc = anc.getparent()
        if not skip:
            out.append(r)
    return out


def run_text(r_el) -> str:
    parts = []
    for child in r_el:
        if child.tag == qn("w:t"):
            parts.append(child.text or "")
        elif child.tag == qn("w:tab"):
            parts.append("\t")
        elif child.tag in (qn("w:br"), qn("w:cr")):
            parts.append("\n")
        elif child.tag == qn("w:noBreakHyphen"):
            parts.append("-")
    return "".join(parts)


def para_text(p_el) -> str:
    return "".join(run_text(r) for r in visible_runs(p_el))


def has_drawing(p_el) -> bool:
    for tag in ("w:drawing", "w:pict", "w:object"):
        for el in p_el.iter(qn(tag)):
            # картинка внутри надписи не делает абзац рисунком
            return True
    return False


def has_math(p_el) -> bool:
    return any(True for _ in p_el.iter(f"{{{M_NS}}}oMath"))


def _break_positions(p_el):
    """(разрыв до текста, разрыв после текста) — по ручным разрывам страницы в абзаце."""
    before = after = False
    seen_text = False
    for r in visible_runs(p_el):
        for child in r:
            if child.tag == qn("w:br") and child.get(qn("w:type")) == "page":
                if seen_text:
                    after = True
                else:
                    before = True
            elif child.tag == qn("w:t") and (child.text or "").strip():
                seen_text = True
                after = False
    return before, after


def has_page_break(p_el) -> bool:
    """Заканчивается ли абзац разрывом страницы (ручным или разрывом раздела)."""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
        return True
    before, after = _break_positions(p_el)
    has_text = bool(para_text(p_el).strip())
    return after or (before and not has_text)


def excerpt(text: str, n: int = 50) -> str:
    text = " ".join(text.split())
    return text if len(text) <= n else text[: n - 1].rstrip() + "…"


# ---------------------------------------------------------------- обход документа


@dataclass
class Block:
    kind: str                 # "p" — абзац, "tbl" — таблица
    el: object                # lxml-элемент
    obj: object               # Paragraph или Table
    num: int = 0              # номер абзаца в документе (с 1), для таблиц 0
    in_table: bool = False
    in_toc_sdt: bool = False
    role: str = "body"
    text: str = ""
    extra: dict = field(default_factory=dict)

    def where(self) -> str:
        return f"абз. {self.num} «{excerpt(self.text)}»" if self.text.strip() else f"абз. {self.num}"


def _sdt_is_toc(sdt) -> bool:
    for gal in sdt.iter(qn("w:docPartGallery")):
        if "Contents" in (gal.get(qn("w:val")) or ""):
            return True
    return False


def iter_blocks(doc) -> list[Block]:
    body = doc._body
    blocks: list[Block] = []
    counter = [0]

    def add_p(p_el, in_table=False, in_toc=False):
        counter[0] += 1
        blocks.append(Block("p", p_el, Paragraph(p_el, body), counter[0], in_table, in_toc,
                            text=para_text(p_el)))

    def walk(container, in_toc=False):
        for child in container.iterchildren():
            if child.tag == qn("w:p"):
                add_p(child, in_toc=in_toc)
            elif child.tag == qn("w:tbl"):
                blocks.append(Block("tbl", child, Table(child, body)))
                for p_el in child.iter(qn("w:p")):
                    anc, nested_box = p_el.getparent(), False
                    while anc is not child:
                        if anc.tag == qn("w:txbxContent"):
                            nested_box = True
                            break
                        anc = anc.getparent()
                    if not nested_box:
                        add_p(p_el, in_table=True, in_toc=in_toc)
            elif child.tag == qn("w:sdt"):
                content = child.find(qn("w:sdtContent"))
                if content is not None:
                    walk(content, in_toc or _sdt_is_toc(child))

    walk(doc.element.body)
    return blocks


# ---------------------------------------------------------------- эффективное оформление


def _style_chain(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def _doc_defaults(doc, kind: str):
    """kind: 'pPr' или 'rPr' из w:docDefaults."""
    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is None:
        return None
    holder = dd.find(qn("w:pPrDefault" if kind == "pPr" else "w:rPrDefault"))
    return None if holder is None else holder.find(qn(f"w:{kind}"))


def _ppr_value(pPr, attr):
    """Значение свойства абзаца из «сырого» w:pPr (для docDefaults)."""
    if pPr is None:
        return None
    if attr == "alignment":
        jc = pPr.find(qn("w:jc"))
        return None if jc is None else jc.get(qn("w:val"))
    if attr == "first_line_indent":
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            return None
        if ind.get(qn("w:hanging")) is not None:
            return Twips(-int(ind.get(qn("w:hanging"))))
        if ind.get(qn("w:firstLine")) is not None:
            return Twips(int(ind.get(qn("w:firstLine"))))
        return None
    if attr == "line_spacing":
        sp = pPr.find(qn("w:spacing"))
        if sp is None or sp.get(qn("w:line")) is None:
            return None
        line = int(sp.get(qn("w:line")))
        rule = sp.get(qn("w:lineRule")) or "auto"
        return line / 240 if rule == "auto" else Twips(line)
    if attr == "page_break_before":
        el = pPr.find(qn("w:pageBreakBefore"))
        return None if el is None else _on(el)
    return None


_JC_MAP = {"both": WD_ALIGN_PARAGRAPH.JUSTIFY, "center": WD_ALIGN_PARAGRAPH.CENTER,
           "left": WD_ALIGN_PARAGRAPH.LEFT, "start": WD_ALIGN_PARAGRAPH.LEFT,
           "right": WD_ALIGN_PARAGRAPH.RIGHT, "end": WD_ALIGN_PARAGRAPH.RIGHT,
           "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE}


def eff_para(doc, paragraph: Paragraph, attr: str):
    """Эффективное свойство абзаца: прямое → стиль и его базовые → docDefaults."""
    value = getattr(paragraph.paragraph_format, attr)
    if value is not None:
        return value
    try:
        style = paragraph.style
    except Exception:  # абзац вне основной части (колонтитул без стилей)
        style = None
    for st in _style_chain(style):
        value = getattr(st.paragraph_format, attr)
        if value is not None:
            return value
    raw = _ppr_value(_doc_defaults(doc, "pPr"), attr)
    if attr == "alignment" and isinstance(raw, str):
        return _JC_MAP.get(raw)
    return raw


def style_names(paragraph: Paragraph) -> list[str]:
    try:
        return [st.name or "" for st in _style_chain(paragraph.style)]
    except Exception:
        return []


def heading_level(doc, paragraph: Paragraph):
    """Уровень заголовка по стилю или outlineLvl; None — не заголовок."""
    for name in style_names(paragraph):
        m = re.match(r"^(heading|заголовок)\s*(\d)$", name.strip(), re.I)
        if m:
            return int(m.group(2))
    pPr = paragraph._p.pPr
    lvl = None if pPr is None else pPr.find(qn("w:outlineLvl"))
    if lvl is None:
        try:
            for st in _style_chain(paragraph.style):
                sp = st.element.pPr
                lvl = None if sp is None else sp.find(qn("w:outlineLvl"))
                if lvl is not None:
                    break
        except Exception:
            lvl = None
    if lvl is not None:
        val = int(lvl.get(qn("w:val")))
        return val + 1 if val < 9 else None
    return None


def has_numbering(paragraph: Paragraph) -> bool:
    pPr = paragraph._p.pPr
    if pPr is not None and pPr.find(qn("w:numPr")) is not None:
        return True
    try:
        for st in _style_chain(paragraph.style):
            sp = st.element.pPr
            if sp is not None and sp.find(qn("w:numPr")) is not None:
                return True
    except Exception:
        pass
    return False


def _on(el) -> bool:
    return (el.get(qn("w:val")) or "true").lower() not in ("0", "false", "off")


class Fonts:
    """Разрешение свойств прогона по цепочке стилей и теме документа."""

    def __init__(self, doc):
        self.doc = doc
        self.theme = {"minor": None, "major": None}
        try:
            theme_part = doc.part.part_related_by(RT.THEME)
            from lxml import etree

            root = etree.fromstring(theme_part.blob)
            for kind in ("minor", "major"):
                latin = root.find(f".//{{{A_NS}}}{kind}Font/{{{A_NS}}}latin")
                if latin is not None:
                    self.theme[kind] = latin.get("typeface")
        except Exception:
            pass
        self.defaults = _doc_defaults(doc, "rPr")

    def _rpr_chain(self, r_el, paragraph: Paragraph):
        rPr = r_el.find(qn("w:rPr"))
        if rPr is not None:
            yield rPr
        style_id = None if rPr is None or rPr.find(qn("w:rStyle")) is None else \
            rPr.find(qn("w:rStyle")).get(qn("w:val"))
        if style_id:
            try:
                style = self.doc.styles.element.get_by_id(style_id)
                from docx.styles.style import StyleFactory

                for st in _style_chain(StyleFactory(style)):
                    if st.element.rPr is not None:
                        yield st.element.rPr
            except Exception:
                pass
        try:
            pstyle = paragraph.style
        except Exception:
            pstyle = None
        for st in _style_chain(pstyle):
            if st.element.rPr is not None:
                yield st.element.rPr
        if self.defaults is not None:
            yield self.defaults

    def font(self, r_el, paragraph) -> str:
        for rPr in self._rpr_chain(r_el, paragraph):
            rf = rPr.find(qn("w:rFonts"))
            if rf is None:
                continue
            theme = rf.get(qn("w:asciiTheme"))
            if theme:
                return self.theme["major" if theme.startswith("major") else "minor"] or f"тема ({theme})"
            if rf.get(qn("w:ascii")):
                return rf.get(qn("w:ascii"))
        return "Times New Roman"

    def size_pt(self, r_el, paragraph) -> float:
        for rPr in self._rpr_chain(r_el, paragraph):
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                return int(sz.get(qn("w:val"))) / 2
        return 10.0

    def color(self, r_el, paragraph) -> str:
        for rPr in self._rpr_chain(r_el, paragraph):
            c = rPr.find(qn("w:color"))
            if c is not None:
                return (c.get(qn("w:val")) or "auto").upper()
        return "AUTO"

    def flag(self, r_el, paragraph, tag: str) -> bool:
        for rPr in self._rpr_chain(r_el, paragraph):
            el = rPr.find(qn(f"w:{tag}"))
            if el is not None:
                return _on(el)
        return False


def is_black(color: str) -> bool:
    return color in ("AUTO", "000000")


def text_runs(block: Block) -> list:
    """Прогоны с непустым видимым текстом."""
    return [r for r in visible_runs(block.el) if run_text(r).strip()]


# ---------------------------------------------------------------- подписи

CAPTION_RE = {
    "figure": re.compile(
        r"^(?P<label>Рисунок|Рис\.?)(?P<sp>\s*)(?P<num>[А-ЯA-Z]?\.?\d+(?:[.\-]\d+)*)"
        r"(?P<sep>\s*[.\-–—:]*\s*)(?P<title>.*)$", re.S),
    "table": re.compile(
        r"^(?P<label>Таблица|Табл\.?)(?P<sp>\s*)(?P<num>[А-ЯA-Z]?\.?\d+(?:[.\-]\d+)*)"
        r"(?P<sep>\s*[.\-–—:]*\s*)(?P<title>.*)$", re.S),
}
CANON_LABEL = {"figure": "Рисунок", "table": "Таблица"}


def match_caption(text: str, kind: str):
    m = CAPTION_RE[kind].match(text.strip()) if text.strip() else None
    if not m:
        return None
    # «Рисунок 1 показывает…» — это текст, а не подпись
    if not re.search(r"[.\-–—:]", m.group("sep")) and m.group("title")[:1].islower():
        return None
    return m


# ---------------------------------------------------------------- роли абзацев


def normalize_heading(text: str) -> str:
    return " ".join(text.replace(" ", " ").split()).rstrip(".").upper()


def classify(doc, blocks: list[Block], profile: dict) -> None:
    """Проставляет block.role.

    Роли: title, toc, empty, figure, formula, caption_figure, caption_table,
    heading_struct, heading_appendix, appendix_title, heading_section,
    list, bib_entry, table_text, body.
    """
    names = {normalize_heading(n) for n in prof(profile, "headings.structural_elements.names", [])}
    appendix_word = (prof(profile, "appendices.heading_word") or "ПРИЛОЖЕНИЕ").upper()
    bib_heading = normalize_heading(prof(profile, "bibliography.heading", ""))
    bib_aliases = {normalize_heading(a) for a in prof(profile, "bibliography.heading_aliases", [])}
    toc_names = {"СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"}

    paras = [b for b in blocks if b.kind == "p"]

    def is_struct(b: Block) -> bool:
        if b.in_table or b.in_toc_sdt:
            return False
        t = normalize_heading(b.text)
        return bool(t) and (t in names or t in bib_aliases or t == bib_heading or t in toc_names)

    def is_appendix(b: Block) -> bool:
        if b.in_table or b.in_toc_sdt:
            return False
        t = normalize_heading(b.text)
        return bool(re.match(rf"^{appendix_word}(\s+[А-ЯA-Z0-9]{{1,2}})?$", t))

    first_struct = next((i for i, b in enumerate(paras) if is_struct(b) or is_appendix(b)), None)

    state = "start" if first_struct else "body"
    appendix_tail = 0
    for i, b in enumerate(paras):
        style = " ".join(style_names(b.obj)).lower()
        t = b.text.strip()
        heading_lvl = heading_level(doc, b.obj)

        if first_struct is not None and i < first_struct:
            b.role = "title"
            continue
        if is_struct(b):
            b.role = "heading_struct"
            nt = normalize_heading(b.text)
            state = "toc" if nt in toc_names else ("bib" if nt == bib_heading or nt in bib_aliases else "body")
            appendix_tail = 0
            continue
        if is_appendix(b):
            b.role = "heading_appendix"
            state, appendix_tail = "body", 2
            continue
        if b.in_toc_sdt or style.startswith("toc") or "оглавление" in style or style.startswith("содержание"):
            b.role = "toc"
            continue
        if not t and not has_drawing(b.el):
            b.role = "empty"
            continue
        if state == "toc" and not heading_lvl:
            b.role = "toc"
            continue
        if appendix_tail and not b.in_table and len(t) < 200 and not t.endswith("."):
            b.role = "appendix_title"
            appendix_tail -= 1
            continue
        appendix_tail = 0
        if has_drawing(b.el) and len(t) < 5:
            b.role = "figure"
            continue
        if has_math(b.el) and len(re.sub(r"[\s()\d.,]", "", t)) < 3:
            b.role = "formula"
            continue
        if match_caption(t, "figure") and not b.in_table:
            b.role = "caption_figure"
            continue
        if match_caption(t, "table") and not b.in_table:
            b.role = "caption_table"
            continue
        if b.in_table:
            b.role = "table_text"
            continue
        if heading_lvl:
            b.role = "heading_section"
            b.extra["level"] = heading_lvl
            continue
        if state == "bib":
            b.role = "bib_entry"
            continue
        if has_numbering(b.obj):
            b.role = "list"
            continue
        b.role = "body"


def prev_block(blocks: list[Block], i: int):
    """Предыдущий непустой блок верхнего уровня (абзацы внутри таблиц пропускаются)."""
    j = i - 1
    while j >= 0:
        b = blocks[j]
        if b.kind == "tbl":
            return b
        if not b.in_table and b.role != "empty":
            return b
        j -= 1
    return None


def next_block(blocks: list[Block], i: int):
    j = i + 1
    while j < len(blocks):
        b = blocks[j]
        if b.kind == "tbl":
            return b
        if not b.in_table and b.role != "empty":
            return b
        j += 1
    return None


def starts_new_page(doc, blocks: list[Block], i: int) -> bool:
    """Начинается ли абзац blocks[i] с новой страницы."""
    b = blocks[i]
    if eff_para(doc, b.obj, "page_break_before"):
        return True
    if _break_positions(b.el)[0]:
        return True
    j = i - 1
    while j >= 0:
        pb = blocks[j]
        if pb.kind == "tbl":
            return False
        if pb.in_table:
            j -= 1
            continue
        if has_page_break(pb.el):
            return True
        if pb.text.strip() or has_drawing(pb.el):
            return False
        j -= 1
    return True


def open_docx(path) -> "Document":
    return Document(str(path))


def mm(length) -> float:
    return Emu(length).mm if length is not None else 0.0


__all__ = [name for name in dir() if not name.startswith("_")] + ["Pt"]
