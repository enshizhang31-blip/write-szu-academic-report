#!/usr/bin/env python3
"""Audit that supplied-template text and formatting remain protected.

This audit intentionally treats the supplied template as the baseline. Generated
content may be appended around it, but every original paragraph and table must
remain in order and keep its run/paragraph/table formatting.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def xml_hash(element) -> str:
    return hashlib.sha256(element.xml.encode("utf-8")).hexdigest()


def run_signature(run) -> dict:
    font = run.font
    rpr = run._r.rPr
    return {
        "text": run.text,
        "style": run.style.style_id if run.style is not None else None,
        "bold": font.bold,
        "italic": font.italic,
        "underline": str(font.underline),
        "name": font.name,
        "eastAsia": rpr.rFonts.get(qn("w:eastAsia")) if rpr is not None and rpr.rFonts is not None else None,
        "size": font.size.pt if font.size is not None else None,
        "color": font.color.rgb.__str__() if font.color is not None and font.color.rgb is not None else None,
        "rpr_hash": xml_hash(rpr) if rpr is not None else None,
    }


def paragraph_signature(paragraph, index: int) -> dict:
    ppr = paragraph._p.pPr
    return {
        "index": index,
        "text": paragraph.text,
        "style": paragraph.style.style_id if paragraph.style is not None else None,
        "ppr_hash": xml_hash(ppr) if ppr is not None else None,
        "runs": [run_signature(run) for run in paragraph.runs],
    }


def cell_signature(cell) -> dict:
    tc_pr = cell._tc.tcPr
    return {
        "text": "\n".join(p.text for p in cell.paragraphs),
        "tcpr_hash": xml_hash(tc_pr) if tc_pr is not None else None,
        "paragraphs": [paragraph_signature(p, i) for i, p in enumerate(cell.paragraphs)],
    }


def table_signature(table, index: int) -> dict:
    tbl_pr = table._tbl.tblPr
    rows = []
    for row in table.rows:
        rows.append([cell_signature(cell) for cell in row.cells])
    return {
        "index": index,
        "text": "\n".join(cell.text for row in table.rows for cell in row.cells),
        "tblpr_hash": xml_hash(tbl_pr) if tbl_pr is not None else None,
        "rows": rows,
    }


def document_signature(path: Path) -> dict:
    doc = Document(path)
    return {
        "paragraphs": [paragraph_signature(p, i) for i, p in enumerate(doc.paragraphs) if p.text.strip()],
        "tables": [table_signature(t, i) for i, t in enumerate(doc.tables)],
    }


def find_ordered_paragraphs(template: list[dict], report: list[dict]) -> tuple[list[tuple[int, int]], list[dict]]:
    matches: list[tuple[int, int]] = []
    missing: list[dict] = []
    position = 0
    for item in template:
        found = None
        for idx in range(position, len(report)):
            if report[idx]["text"] == item["text"]:
                found = idx
                break
        if found is None:
            missing.append(item)
        else:
            matches.append((item["index"], report[found]["index"]))
            position = found + 1
    return matches, missing


def find_ordered_tables(template: list[dict], report: list[dict]) -> tuple[list[tuple[int, int]], list[dict]]:
    matches: list[tuple[int, int]] = []
    missing: list[dict] = []
    position = 0
    for item in template:
        found = None
        for idx in range(position, len(report)):
            if report[idx]["text"] == item["text"]:
                found = idx
                break
        if found is None:
            missing.append(item)
        else:
            matches.append((item["index"], report[found]["index"]))
            position = found + 1
    return matches, missing


def compare_paragraph(template_item: dict, report_item: dict) -> list[dict]:
    issues: list[dict] = []
    keys = ["style", "ppr_hash"]
    for key in keys:
        if template_item[key] != report_item[key]:
            issues.append({"rule": f"paragraph_{key}_changed", "template": template_item["index"], "text": template_item["text"][:80]})
    if len(template_item["runs"]) != len(report_item["runs"]):
        issues.append({"rule": "paragraph_run_count_changed", "template": template_item["index"], "text": template_item["text"][:80]})
        return issues
    for run_index, (left, right) in enumerate(zip(template_item["runs"], report_item["runs"])):
        if left != right:
            changed = [key for key in left.keys() if left[key] != right.get(key)]
            issues.append({
                "rule": "run_format_changed",
                "template": template_item["index"],
                "run": run_index,
                "changed": changed,
                "text": left["text"][:80],
            })
    return issues


def compare_table(template_item: dict, report_item: dict) -> list[dict]:
    issues: list[dict] = []
    if template_item["tblpr_hash"] != report_item["tblpr_hash"]:
        issues.append({"rule": "table_properties_changed", "template": template_item["index"]})
    if len(template_item["rows"]) != len(report_item["rows"]):
        issues.append({"rule": "table_row_count_changed", "template": template_item["index"]})
        return issues
    for r_idx, (left_row, right_row) in enumerate(zip(template_item["rows"], report_item["rows"])):
        if len(left_row) != len(right_row):
            issues.append({"rule": "table_cell_count_changed", "template": template_item["index"], "row": r_idx})
            continue
        for c_idx, (left_cell, right_cell) in enumerate(zip(left_row, right_row)):
            if left_cell["text"] != right_cell["text"]:
                issues.append({"rule": "table_cell_text_changed", "template": template_item["index"], "row": r_idx, "cell": c_idx})
            if left_cell["tcpr_hash"] != right_cell["tcpr_hash"]:
                issues.append({"rule": "table_cell_properties_changed", "template": template_item["index"], "row": r_idx, "cell": c_idx})
    return issues


def audit(template_path: Path, report_path: Path) -> dict:
    template = document_signature(template_path)
    report = document_signature(report_path)
    template_paragraph_by_index = {item["index"]: item for item in template["paragraphs"]}
    report_paragraph_by_index = {item["index"]: item for item in report["paragraphs"]}
    template_table_by_index = {item["index"]: item for item in template["tables"]}
    report_table_by_index = {item["index"]: item for item in report["tables"]}
    paragraph_matches, missing_paragraphs = find_ordered_paragraphs(template["paragraphs"], report["paragraphs"])
    table_matches, missing_tables = find_ordered_tables(template["tables"], report["tables"])
    issues: list[dict] = []

    for template_idx, report_idx in paragraph_matches:
        issues.extend(compare_paragraph(template_paragraph_by_index[template_idx], report_paragraph_by_index[report_idx]))
    for template_idx, report_idx in table_matches:
        issues.extend(compare_table(template_table_by_index[template_idx], report_table_by_index[report_idx]))

    for item in missing_paragraphs:
        issues.append({"rule": "protected_paragraph_missing_or_reordered", "template": item["index"], "text": item["text"][:80]})
    for item in missing_tables:
        issues.append({"rule": "protected_table_missing_or_reordered", "template": item["index"]})

    return {
        "template": str(template_path),
        "report": str(report_path),
        "protected_paragraphs": len(template["paragraphs"]),
        "protected_tables": len(template["tables"]),
        "errors": len(issues),
        "issues": issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("template", type=Path)
    parser.add_argument("report", type=Path)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    result = audit(args.template, args.report)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json:
        args.json.write_text(text, encoding="utf-8")
    print(text)
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
