#!/usr/bin/env python3
"""Warn when code-like names used in a DOCX are absent from supplied source code."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CODE_EXTENSIONS = {".java", ".py", ".js", ".ts", ".tsx", ".jsx", ".cpp", ".cc", ".c", ".h", ".hpp", ".cs", ".go"}
DECLARATION = re.compile(r"\b(?:class|interface|enum|record|struct|def|function|fun|type)\s+([A-Za-z_]\w*)")
DOC_IDENTIFIER = re.compile(r"\b[A-Z][A-Za-z0-9_]{2,}\b")
ALLOW = {"UML", "OO", "API", "HTTP", "JSON", "XML", "DOCX", "Times", "Roman", "Word", "FIRST", "FOLLOW", "First", "Follow", "ACTION", "GOTO", "CSV", "UTF", "Python", "Simple", "Grammar", "Dict", "List", "Tuple", "Sequence", "EPS", "EPSILON", "A_first_set", "B_follow_set", "C_ll1_expression", "D_lr_driver", "E_lr0_table", "X1X2"}


def document_text(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    return "\n".join("".join(t.text or "" for t in p.iter(W + "t")) for p in root.iter(W + "p"))


def source_declarations(root: Path) -> set[str]:
    names: set[str] = set()
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in CODE_EXTENSIONS:
            try:
                text = path.read_text(encoding="utf-8", errors="ignore")
            except OSError:
                continue
            names.update(DECLARATION.findall(text))
    return names


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("source_dir", type=Path)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    declared = source_declarations(args.source_dir)
    mentioned = set(DOC_IDENTIFIER.findall(document_text(args.docx))) - ALLOW
    unmatched = sorted(mentioned - declared)
    result = {
        "declared_identifiers": sorted(declared),
        "mentioned_code_like_identifiers": sorted(mentioned),
        "unmatched_mentions": unmatched,
        "note": "Review unmatched mentions; high strictness requires resolving false names or documenting intentional external terms.",
    }
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json:
        args.json.write_text(text, encoding="utf-8")
    print(text)
    return 1 if unmatched else 0


if __name__ == "__main__":
    raise SystemExit(main())
