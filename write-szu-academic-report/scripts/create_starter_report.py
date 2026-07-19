#!/usr/bin/env python3
"""Create a strict-format SZU academic report starter DOCX."""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt


def set_font(target, size: float, bold: bool | None = None) -> None:
    target.font.name = "Times New Roman"
    target._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    target.font.size = Pt(size)
    target.font.color.rgb = None
    if bold is not None:
        target.font.bold = bold


def configure_style(style, size: float, bold: bool = False) -> None:
    style.font.name = "Times New Roman"
    style._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = None


def remove_paragraph_border(style) -> None:
    p_pr = style._element.pPr
    if p_pr is None:
        return
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is not None:
        p_pr.remove(p_bdr)


def set_cell_reference_format(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", 80), ("start", 100), ("bottom", 80), ("end", 100)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:color"), "000000")


def add_body(doc: Document, text: str):
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Pt(21)
    set_font(p.add_run(text), 10.5)
    return p


def build(output: Path) -> None:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(25.4)
    section.bottom_margin = Mm(25.4)
    section.left_margin = Mm(31.7)
    section.right_margin = Mm(31.7)

    configure_style(doc.styles["Normal"], 10.5)
    configure_style(doc.styles["Title"], 15.0)
    remove_paragraph_border(doc.styles["Title"])
    configure_style(doc.styles["Heading 1"], 14.0, True)
    configure_style(doc.styles["Heading 2"], 12.0, True)
    configure_style(doc.styles["Heading 3"], 10.5, True)
    configure_style(doc.styles["Heading 4"], 9.0, True)
    configure_style(doc.styles["Caption"], 10.5)

    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(title.add_run("深圳大学实验报告标题"), 15.0)

    doc.add_heading("1 实验目的", level=1)
    add_body(doc, "在此填写实验目的。正文使用宋体五号，英文与数字使用 Times New Roman 五号，并设置两字符首行缩进。")

    doc.add_heading("2 实验内容", level=1)
    doc.add_heading("2.1 问题分析", level=2)
    add_body(doc, "在此填写问题分析，并保持标题层级与代码模块、类和接口的组织结构一致。")

    doc.add_heading("2.2 设计与实现", level=2)
    doc.add_heading("2.2.1 核心模块", level=3)
    add_body(doc, "在此填写设计与实现内容。")
    add_body(doc, "（1）正文中需要区分多个项目时，必须使用中文括号序号。")
    add_body(doc, "（2）禁止使用点、横线或装饰性符号作为正文项目标记。")

    table = doc.add_table(rows=2, cols=3)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            fill = "FFFFFF" if row_index == 0 or row_index % 2 == 1 else "F2F2F2"
            set_cell_reference_format(cell, fill=fill)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_index == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = None
            text = ["项目", "说明", "结果"][col_index] if row_index == 0 else ["示例", "替换为实际内容", "通过"][col_index]
            set_font(p.add_run(text), 10.5, row_index == 0)

    caption = doc.add_paragraph(style="Caption")
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(caption.add_run("图 1 系统结构示意图（占位）"), 10.5)

    doc.add_heading("3 实验结果与分析", level=1)
    add_body(doc, "在此填写实验结果及分析。")
    doc.add_heading("4 总结", level=1)
    add_body(doc, "在此填写总结。")

    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    build(args.output)
    print(args.output)


if __name__ == "__main__":
    main()
