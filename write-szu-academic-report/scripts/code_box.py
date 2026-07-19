#!/usr/bin/env python3
"""Reusable python-docx helper for framed SZU code boxes."""

from __future__ import annotations

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


def _set_attr(parent, tag: str, **attrs: str) -> None:
    node = parent.find(qn(f"w:{tag}"))
    if node is None:
        node = OxmlElement(f"w:{tag}")
        parent.append(node)
    for key, value in attrs.items():
        node.set(qn(f"w:{key}"), value)


def ensure_code_box_style(doc: Document):
    styles = doc.styles
    style = styles["CodeBox"] if "CodeBox" in [s.name for s in styles] else styles.add_style("CodeBox", WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = "Times New Roman"
    style.font.size = Pt(9)
    style._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "SimSun")
    style.font.color.rgb = None
    style.paragraph_format.first_line_indent = Pt(0)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing = 1.0
    return style


def add_code_box(doc: Document, code: str):
    """Add one framed, single-cell code container and return its table."""
    ensure_code_box_style(doc)
    table = doc.add_table(rows=1, cols=1)
    table.autofit = True
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "6")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    tbl_pr.append(borders)

    cell = table.cell(0, 0)
    tc_pr = cell._tc.get_or_add_tcPr()
    _set_attr(tc_pr, "shd", fill="F7F7F7")
    margins = OxmlElement("w:tcMar")
    for side, value in (("top", "100"), ("start", "120"), ("bottom", "100"), ("end", "120")):
        margin = OxmlElement(f"w:{side}")
        margin.set(qn("w:w"), value)
        margin.set(qn("w:type"), "dxa")
        margins.append(margin)
    tc_pr.append(margins)

    first = cell.paragraphs[0]
    lines = code.rstrip("\n").splitlines() or [""]
    for index, line in enumerate(lines):
        paragraph = first if index == 0 else cell.add_paragraph()
        paragraph.style = "CodeBox"
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(line or " ")
        run.font.name = "Times New Roman"
        run.font.size = Pt(9)
        run.font.color.rgb = None
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "SimSun")
    return table
