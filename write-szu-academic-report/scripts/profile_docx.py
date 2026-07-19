#!/usr/bin/env python3
"""Extract sample-derived DOCX structure and table-format profiles."""

from __future__ import annotations

import argparse
import json
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def val(node: ET.Element | None, name: str = "val") -> str | None:
    return node.get(W + name) if node is not None else None


def profile(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        doc = ET.fromstring(zf.read("word/document.xml"))
        styles = ET.fromstring(zf.read("word/styles.xml"))

    body = doc.find(W + "body")
    tables = body.findall(".//" + W + "tbl") if body is not None else []
    table_profiles = []
    for table in tables:
        pr = table.find(W + "tblPr")
        borders = pr.find(W + "tblBorders") if pr is not None else None
        table_profiles.append({
            "style": val(pr.find(W + "tblStyle")) if pr is not None else None,
            "alignment": val(pr.find(W + "jc")) if pr is not None else None,
            "width": val(pr.find(W + "tblW"), "w") if pr is not None else None,
            "width_type": val(pr.find(W + "tblW"), "type") if pr is not None else None,
            "rows": len(table.findall(W + "tr")),
            "columns": max((len(row.findall(W + "tc")) for row in table.findall(W + "tr")), default=0),
            "borders": {
                edge: val(borders.find(W + edge)) if borders is not None else None
                for edge in ("top", "left", "bottom", "right", "insideH", "insideV")
            },
        })

    style_names = Counter()
    for paragraph in body.iter(W + "p") if body is not None else []:
        ppr = paragraph.find(W + "pPr")
        style_names[val(ppr.find(W + "pStyle")) if ppr is not None else None] += 1

    sect = body.find(W + "sectPr") if body is not None else None
    return {
        "file": str(path),
        "paragraphs": len(list(body.iter(W + "p"))) if body is not None else 0,
        "tables": len(tables),
        "drawings": len(list(body.iter(W + "drawing"))) if body is not None else 0,
        "page": {
            "size": dict(sect.find(W + "pgSz").attrib) if sect is not None and sect.find(W + "pgSz") is not None else {},
            "margins": dict(sect.find(W + "pgMar").attrib) if sect is not None and sect.find(W + "pgMar") is not None else {},
        },
        "paragraph_styles": dict(style_names),
        "table_profiles": table_profiles,
        "defined_styles": len(styles.findall(W + "style")),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path, nargs="+")
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    result = [profile(path) for path in args.docx]
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json:
        args.json.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
