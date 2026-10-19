#!/usr/bin/env python3
"""Копия документа со всеми принятыми исправлениями — чтобы посмотреть итоговый вид.

    python3 scripts/accept_changes.py исправлено.docx -o итог.docx
"""

from __future__ import annotations

import argparse
import sys

import nk_common as nk
from nk_tracked import accept_all


def main(argv=None):
    ap = argparse.ArgumentParser(description="Принять все исправления в копии .docx")
    ap.add_argument("docx")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args(argv)
    doc = nk.open_docx(args.docx)
    for part in [doc.part] + [r.target_part for r in doc.part.rels.values()
                              if not r.is_external and hasattr(r.target_part, "element")]:
        accept_all(part.element)
    doc.save(args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
