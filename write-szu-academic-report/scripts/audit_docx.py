#!/usr/bin/env python3
"""Audit the high-strictness SZU report contract directly from DOCX OOXML."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DRAWING = W + "drawing"
PICT = W + "pict"
LATIN_RE = re.compile(r"[A-Za-z0-9]")
CJK_RE = re.compile(r"[\u3400-\u9fff]")
CAPTION_RE = re.compile(r"^\s*图\s*\d+")
FORBIDDEN_BULLET_RE = re.compile(r"^\s*[•●○◆■□▪▫·\-]\s*")
CODE_LIKE_RE = re.compile(
    r"^\s*(?:function|class|interface|const|let|var|def|import|from|public|private|"
    r"protected|if|for|while|return|\.on)\b.*(?:[;{}]|\)$)"
)


def attr(node: ET.Element | None, name: str) -> str | None:
    return node.get(W + name) if node is not None else None


def child(node: ET.Element | None, name: str) -> ET.Element | None:
    return node.find(W + name) if node is not None else None


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iter(W + "t")).strip()


def direct_table_paragraphs(table: ET.Element) -> list[ET.Element]:
    return table.findall(f"./{W}tr/{W}tc/{W}p")


def run_props(rpr: ET.Element | None) -> dict:
    fonts = child(rpr, "rFonts")
    size = attr(child(rpr, "sz"), "val")
    return {
        "eastAsia": attr(fonts, "eastAsia"),
        "ascii": attr(fonts, "ascii"),
        "hAnsi": attr(fonts, "hAnsi"),
        "size_pt": float(size) / 2 if size and size.isdigit() else None,
        "color": attr(child(rpr, "color"), "val"),
    }


def merge_props(base: dict, direct: dict) -> dict:
    return {key: direct.get(key) if direct.get(key) is not None else base.get(key) for key in base}


def style_map(root: ET.Element) -> tuple[dict[str, dict], dict]:
    result: dict[str, dict] = {}
    defaults = {}
    doc_defaults = child(root, "docDefaults")
    rpr_default = child(doc_defaults, "rPrDefault")
    if rpr_default is not None:
        defaults = run_props(child(rpr_default, "rPr"))
    for style in root.findall(W + "style"):
        sid = attr(style, "styleId")
        if sid:
            result[sid] = {
                "name": attr(child(style, "name"), "val") or sid,
                "pPr": child(style, "pPr"),
                "rPr": child(style, "rPr"),
                "basedOn": attr(child(style, "basedOn"), "val"),
            }
    return result, defaults


def resolved_style_props(styles: dict[str, dict], defaults: dict, sid: str) -> dict:
    chain, seen = [], set()
    while sid and sid not in seen and sid in styles:
        seen.add(sid)
        chain.append(styles[sid])
        sid = styles[sid].get("basedOn") or ""
    props = defaults
    for style in reversed(chain):
        props = merge_props(props, run_props(style.get("rPr")))
    return props


def heading_level(text: str, style_name: str) -> int | None:
    for level in range(1, 5):
        if str(level) in style_name and ("heading" in style_name.lower() or "标题" in style_name):
            return level
    match = re.match(r"^(\d+(?:\.\d+){0,3})\s+", text)
    return min(4, match.group(1).count(".") + 1) if match else None


def classify(text: str, style_name: str, has_drawing: bool, in_table: bool, in_code_box: bool, before_heading: bool) -> str:
    lower = style_name.lower()
    if in_code_box or "codebox" in lower:
        return "code"
    if has_drawing:
        return "figure"
    if CAPTION_RE.match(text) or "caption" in lower or "题注" in style_name:
        return "caption"
    if "title" in lower or style_name == "标题":
        return "title"
    if "heading" in lower or "标题 " in style_name or heading_level(text, style_name):
        return "heading"
    if in_table:
        return "table"
    if before_heading:
        return "cover"
    return "body"


def expected_heading_size(text: str, style_name: str) -> float | None:
    level = heading_level(text, style_name)
    return {1: 14.0, 2: 12.0, 3: 10.5, 4: 9.0}.get(level)


def audit(path: Path) -> dict:
    issues: list[dict] = []
    counts = {"paragraphs": 0, "tables": 0, "code_boxes": 0, "figures": 0, "captions": 0, "headings": 0}
    with zipfile.ZipFile(path) as zf:
        doc = ET.fromstring(zf.read("word/document.xml"))
        styles_root = ET.fromstring(zf.read("word/styles.xml"))
        styles, defaults = style_map(styles_root)

    body = doc.find(W + "body")
    if body is None:
        raise ValueError("DOCX has no document body")

    all_tables = body.findall(".//" + W + "tbl")
    code_boxes = {
        table
        for table in all_tables
        if attr(child(child(table, "tblPr"), "tblDescription"), "val") == "SZU_CODE_BOX"
        or any(attr(child(child(p, "pPr"), "pStyle"), "val") == "CodeBox" for p in direct_table_paragraphs(table))
    }
    template_labels = ("实验目的与要求", "方法、步骤", "实验过程及内容", "实验结论", "心得体会", "指导教师批阅意见", "备注")
    template_frames = {
        table
        for table in all_tables
        if attr(child(child(table, "tblPr"), "tblDescription"), "val") == "SZU_TEMPLATE_FRAME"
        or (
            len(table.findall("./" + W + "tr/" + W + "tc")) == len(table.findall(W + "tr"))
            and len(table.findall(W + "tr")) >= 5
            and sum(label in paragraph_text(table) for label in template_labels) >= 5
        )
    } - code_boxes
    cover_labels = ("课程名称", "项目名称", "学    院", "专    业", "指导教师", "报 告 人", "学号")
    layout_tables = {
        table
        for table in all_tables
        if len(table.findall(".//" + W + "tc")) == 1
        and sum(label in paragraph_text(table) for label in cover_labels) >= 3
    } - template_frames - code_boxes
    counts["code_boxes"] = len(code_boxes)
    counts["tables"] = len(all_tables) - len(code_boxes) - len(template_frames) - len(layout_tables)

    for table_index, table in enumerate(all_tables, start=1):
        tbl_pr = child(table, "tblPr")
        if table in code_boxes:
            if child(tbl_pr, "tblBorders") is None:
                issues.append({"severity": "error", "table": table_index, "rule": "code_box_border_required"})
            continue
        if table in template_frames:
            if child(tbl_pr, "tblBorders") is None:
                issues.append({"severity": "error", "table": table_index, "rule": "template_frame_border_required"})
            for cell in table.findall("./" + W + "tr/" + W + "tc"):
                fill = attr(child(child(cell, "tcPr"), "shd"), "fill")
                if fill is not None:
                    issues.append({"severity": "error", "table": table_index, "rule": "template_frame_no_shading", "actual": fill})
            continue
        if table in layout_tables:
            continue
        rows = table.findall(W + "tr")
        if not rows:
            continue
        header_fills = [attr(child(child(cell, "tcPr"), "shd"), "fill") for cell in rows[0].findall(W + "tc")]
        if not header_fills or any(fill != "FFFFFF" for fill in header_fills):
            issues.append({"severity": "error", "table": table_index, "rule": "table_white_header_required"})
        for row_index, row in enumerate(rows[1:], start=1):
            expected = "FFFFFF" if row_index % 2 == 1 else "F2F2F2"
            fills = [attr(child(child(cell, "tcPr"), "shd"), "fill") for cell in row.findall(W + "tc")]
            if any(fill != expected for fill in fills):
                issues.append({"severity": "error", "table": table_index, "rule": "table_alternating_fill_required"})
        if child(tbl_pr, "tblBorders") is None:
            issues.append({"severity": "error", "table": table_index, "rule": "table_borders_required"})

    parents = {node: parent for parent in body.iter() for node in parent}
    seen_heading = False
    previous_heading: int | None = None
    for index, p in enumerate(body.iter(W + "p"), start=1):
        text = paragraph_text(p)
        has_visual = p.find(".//" + DRAWING) is not None or p.find(".//" + PICT) is not None
        if not text and not has_visual:
            continue
        counts["paragraphs"] += 1
        ppr = child(p, "pPr")
        sid = attr(child(ppr, "pStyle"), "val") or ""
        style = styles.get(sid, {})
        style_name = style.get("name", sid)
        ancestor = parents.get(p)
        in_table = in_code_box = in_template_frame = in_layout_table = False
        while ancestor is not None:
            if ancestor.tag == W + "tc":
                in_table = True
            if ancestor in code_boxes:
                in_code_box = True
                break
            if ancestor in template_frames:
                in_template_frame = True
                break
            if ancestor in layout_tables:
                in_layout_table = True
                break
            ancestor = parents.get(ancestor)

        role = "cover" if in_layout_table else classify(text, style_name, has_visual, in_table or in_template_frame, in_code_box, not seen_heading)
        if role == "figure":
            counts["figures"] += 1
        elif role == "caption":
            counts["captions"] += 1
        elif role == "heading":
            counts["headings"] += 1
            seen_heading = True
            level = heading_level(text, style_name)
            if level and previous_heading and level > previous_heading + 1:
                issues.append({"severity": "error", "paragraph": index, "rule": "heading_level_skip", "text": text[:80]})
            previous_heading = level or previous_heading

        jc = attr(child(ppr, "jc"), "val") or attr(child(style.get("pPr"), "jc"), "val")
        ind = child(ppr, "ind")
        if ind is None:
            ind = child(style.get("pPr"), "ind")
        first_line = attr(ind, "firstLine") if ind is not None else None
        if role in {"figure", "caption"} and jc != "center":
            issues.append({"severity": "error", "paragraph": index, "rule": "center_figure_caption", "text": text[:80]})
        if role == "body" and text and first_line not in {"420", "421"}:
            issues.append({"severity": "warning", "paragraph": index, "rule": "body_two_character_indent", "text": text[:80]})
        if role == "body" and FORBIDDEN_BULLET_RE.match(text):
            issues.append({"severity": "error", "paragraph": index, "rule": "parenthesized_numbering_required", "text": text[:80]})
        if role == "body" and CODE_LIKE_RE.match(text):
            issues.append({"severity": "error", "paragraph": index, "rule": "displayed_code_requires_code_box", "text": text[:80]})

        for run in p.findall(W + "r"):
            run_text = "".join(t.text or "" for t in run.iter(W + "t"))
            if not run_text.strip():
                continue
            props = merge_props(resolved_style_props(styles, defaults, sid), run_props(child(run, "rPr")))
            expected_size = {"title": 15.0, "code": 9.0}.get(role, 10.5)
            if role == "heading":
                expected_size = expected_heading_size(text, style_name)
            if role in {"figure", "cover"}:
                continue
            if in_template_frame and not in_code_box:
                # Protected template/layout-frame text may intentionally keep the supplied
                # template font size and color when the user requests exact preservation.
                continue
            if expected_size and props["size_pt"] is not None and props["size_pt"] != expected_size:
                issues.append({"severity": "error", "paragraph": index, "rule": "font_size", "expected": expected_size, "actual": props["size_pt"], "text": run_text[:50]})
            if CJK_RE.search(run_text) and props["eastAsia"] not in {"宋体", "SimSun"}:
                issues.append({"severity": "error", "paragraph": index, "rule": "chinese_font", "actual": props["eastAsia"], "text": run_text[:50]})
            if LATIN_RE.search(run_text) and props["ascii"] != "Times New Roman" and props["hAnsi"] != "Times New Roman":
                issues.append({"severity": "error", "paragraph": index, "rule": "latin_font", "actual": props["ascii"] or props["hAnsi"], "text": run_text[:50]})
            allowed_colors = {None, "auto", "000000"}
            if role == "table":
                allowed_colors.add("08090C")
            if props["color"] not in allowed_colors:
                issues.append({"severity": "error", "paragraph": index, "rule": "default_black_color", "actual": props["color"], "text": run_text[:50]})

    return {
        "file": str(path),
        "strictness": "high",
        "counts": counts,
        "errors": sum(issue["severity"] == "error" for issue in issues),
        "warnings": sum(issue["severity"] == "warning" for issue in issues),
        "issues": issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    result = audit(args.docx)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json:
        args.json.write_text(text, encoding="utf-8")
    print(text)
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
