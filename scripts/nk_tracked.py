"""Правки в режиме рецензирования (track changes).

Изменения оформления пишутся как w:pPrChange / w:rPrChange / w:sectPrChange:
в Word автор видит «Форматировано: …» и может принять или отклонить правку.
Изменения текста (тире в подписи, точка в конце заголовка) — как w:del + w:ins.
"""

from __future__ import annotations

import copy
from contextlib import contextmanager
from datetime import datetime, timezone

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from nk_common import run_text, visible_runs

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_CHANGE_TAGS = {
    qn("w:pPr"): qn("w:pPrChange"),
    qn("w:rPr"): qn("w:rPrChange"),
    qn("w:sectPr"): qn("w:sectPrChange"),
}
# что не копируем в «старое» состояние
_EXCLUDE = {
    qn("w:pPr"): {qn("w:rPr"), qn("w:sectPr"), qn("w:pPrChange")},
    qn("w:rPr"): {qn("w:rPrChange")},
    qn("w:sectPr"): {qn("w:headerReference"), qn("w:footerReference"), qn("w:sectPrChange")},
}


class Revisions:
    def __init__(self, doc, author: str = "Нормоконтроль", enabled: bool = True):
        self.doc = doc
        self.author = author
        self.enabled = enabled
        self.date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        self.next_id = self._max_id() + 1
        self.count = 0

    def _max_id(self) -> int:
        best = 0
        parts = [self.doc.part] + [rel.target_part for rel in self.doc.part.rels.values()
                                   if not rel.is_external and hasattr(rel.target_part, "element")]
        for part in parts:
            for el in part.element.iter():
                val = el.get(qn("w:id"))
                if val and val.lstrip("-").isdigit():
                    best = max(best, int(val))
        return best

    def mark(self, el):
        el.set(qn("w:id"), str(self.next_id))
        el.set(qn("w:author"), self.author)
        el.set(qn("w:date"), self.date)
        self.next_id += 1
        self.count += 1
        return el

    # --- оформление

    @staticmethod
    def _snapshot(props):
        if props is None:
            return []
        excl = _EXCLUDE[props.tag]
        return [copy.deepcopy(c) for c in props if c.tag not in excl]

    @staticmethod
    def _serialize(children) -> bytes:
        from lxml import etree

        return b"".join(etree.tostring(c) for c in children)

    @contextmanager
    def props(self, owner, props_tag: str, get_or_add):
        """Отслеживает изменение pPr / rPr / sectPr.

        owner — элемент, в котором лежат свойства (w:p, w:r) или сам w:sectPr;
        get_or_add — функция, возвращающая (и при нужде создающая) элемент свойств.
        """
        tag = qn(props_tag)
        before_el = owner if owner.tag == tag else owner.find(tag)
        before = self._snapshot(before_el)
        yield
        after_el = owner if owner.tag == tag else owner.find(tag)
        after = self._snapshot(after_el)
        if self._serialize(before) == self._serialize(after):
            return
        if not self.enabled:
            return
        props_el = after_el if after_el is not None else get_or_add()
        change_tag = _CHANGE_TAGS[tag]
        change = props_el.find(change_tag)
        if change is None:
            change = self.mark(OxmlElement(props_tag + "Change"))
            old = OxmlElement(props_tag)
            for c in before:
                old.append(c)
            change.append(old)
        props_el.append(change)  # элемент изменения должен быть последним

    def para(self, paragraph):
        return self.props(paragraph._p, "w:pPr", paragraph._p.get_or_add_pPr)

    def run(self, r_el):
        return self.props(r_el, "w:rPr", r_el.get_or_add_rPr)

    def section(self, section):
        return self.props(section._sectPr, "w:sectPr", lambda: section._sectPr)

    # --- текст

    def _wrap(self, tag: str):
        return self.mark(OxmlElement(tag))

    def replace_span(self, p_el, start: int, end: int, new_text: str) -> bool:
        """Заменяет символы [start, end) видимого текста абзаца на new_text.

        Возвращает False, если затронут «сложный» прогон (поле, вложенный
        элемент, несколько w:t) — такую правку оставляем Claude или автору.
        """
        runs, pos = [], 0
        for r in visible_runs(p_el):
            t = run_text(r)
            runs.append([r, pos, pos + len(t)])
            pos += len(t)
        if not (0 <= start <= end <= pos):
            return False

        def simple(r):
            kids = [c for c in r if c.tag != qn("w:rPr")]
            return r.getparent() is p_el and len(kids) == 1 and kids[0].tag == qn("w:t")

        if start == end:
            host = next((x for x in runs if x[1] <= start < x[2]), None) or \
                   next((x for x in reversed(runs) if x[2] == start), None)
            if host is None or not simple(host[0]):
                return False
            left = self._split(host[0], start - host[1])
            anchor = left if start > host[1] else None
            ins = self._ins_run(host[0], new_text)
            if anchor is None:
                host[0].addprevious(ins)
            else:
                anchor.addnext(ins)
            return True

        touched = [x for x in runs if x[1] < end and x[2] > start]
        if not touched or not all(simple(x[0]) for x in touched):
            return False
        if not self.enabled:
            return self._replace_plain(touched, start, end, new_text)

        first, last = touched[0], touched[-1]
        # отрезаем хвост у последнего и голову у первого
        if end < last[2]:
            self._split(last[0], end - last[1])
        span_runs = [x[0] for x in touched]
        if start > first[1]:
            left = first[0]
            self._split(left, start - first[1])
            span_runs[0] = left.getnext()
        rpr_source = span_runs[0]
        d = self._wrap("w:del")
        span_runs[0].addprevious(d)
        for r in span_runs:
            d.append(r)
            for t in r.findall(qn("w:t")):
                t.tag = qn("w:delText")
                t.set(XML_SPACE, "preserve")
        if new_text:
            d.addnext(self._ins_run(rpr_source, new_text))
        return True

    def _replace_plain(self, touched, start, end, new_text) -> bool:
        first = touched[0]
        for i, (r, a, b) in enumerate(touched):
            t = r.find(qn("w:t"))
            text = t.text or ""
            lo, hi = max(start - a, 0), min(end - a, len(text))
            text = text[:lo] + (new_text if i == 0 else "") + text[hi:]
            t.text = text
            t.set(XML_SPACE, "preserve")
        return first is not None

    @staticmethod
    def _split(r, k: int):
        """Делит простой прогон по смещению k; возвращает левую часть (сам r)."""
        t = r.find(qn("w:t"))
        text = t.text or ""
        if k <= 0 or k >= len(text):
            return r
        right = copy.deepcopy(r)
        t.text = text[:k]
        t.set(XML_SPACE, "preserve")
        rt = right.find(qn("w:t"))
        rt.text = text[k:]
        rt.set(XML_SPACE, "preserve")
        r.addnext(right)
        return r

    def _ins_run(self, like_r, text: str):
        r = OxmlElement("w:r")
        rpr = like_r.find(qn("w:rPr"))
        if rpr is not None:
            rpr = copy.deepcopy(rpr)
            ch = rpr.find(qn("w:rPrChange"))
            if ch is not None:
                rpr.remove(ch)
            r.append(rpr)
        t = OxmlElement("w:t")
        t.text = text
        t.set(XML_SPACE, "preserve")
        r.append(t)
        if not self.enabled:
            return r
        ins = self._wrap("w:ins")
        ins.append(r)
        return ins

    def wrap_inserted(self, runs):
        """Помечает уже вставленные прогоны как вставку."""
        if not self.enabled or not runs:
            return
        ins = self._wrap("w:ins")
        runs[0].addprevious(ins)
        for r in runs:
            ins.append(r)


def ensure_change_last(root):
    """Страховка: w:*PrChange всегда последний дочерний элемент свойств."""
    for change_tag in _CHANGE_TAGS.values():
        for ch in list(root.iter(change_tag)):
            parent = ch.getparent()
            if parent is not None and parent[-1] is not ch:
                parent.append(ch)


def accept_all(root):
    """Принимает все исправления в XML-дереве (для просмотра итогового вида)."""
    for tag in ("w:del", "w:moveFrom"):
        for el in list(root.iter(qn(tag))):
            el.getparent().remove(el)
    for tag in ("w:ins", "w:moveTo"):
        for el in list(root.iter(qn(tag))):
            parent = el.getparent()
            idx = parent.index(el)
            for child in list(el):
                parent.insert(idx, child)
                idx += 1
            parent.remove(el)
    for change_tag in list(_CHANGE_TAGS.values()) + [qn("w:tblPrChange"), qn("w:trPrChange"),
                                                     qn("w:tcPrChange"), qn("w:numberingChange")]:
        for el in list(root.iter(change_tag)):
            el.getparent().remove(el)
