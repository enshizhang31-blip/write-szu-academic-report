---
name: write-szu-academic-report
description: Create, complete, rewrite, and audit Shenzhen University Chinese academic and experiment reports in DOCX with high-strictness formatting, including supplied-template preservation, black-only template text, framed small-five code boxes, Songti Chinese text, Times New Roman Latin text, semantic headings, centered figures, and reference-derived tables.
---

# Shenzhen University Academic Report

Apply the user's explicit rules first. Read `references/format-standard.md` before editing or drafting.

## Required Workflow

1. Inspect every supplied template, tutorial, source file, and screenshot requirement before writing.
2. When a report template is supplied, treat it as the authoritative document skeleton and formatting baseline. Do not rewrite, rename, delete, reorder, summarize, replace, recolor, restyle, resize, bold/unbold, re-indent, re-border, shade, or otherwise modify any existing template content.
3. Before drafting into a supplied template, decide whether it is a frame-required report. Inspect page 2 and later for large bordered fill-in frames or text boxes that contain labels such as `实验过程及内容：`, `数据分析：`, `实验结论：`, `指导教师批阅意见`, `成绩评定`, or `备注`. If such fill-in frames exist, or if the user says the report must be completed in a frame/text box and not outside it, use frame-required mode.
4. In frame-required mode, insert generated report content only inside the matching template frame cells or text boxes. Put process text, figures, captions, code boxes, and code explanations in `实验过程及内容：`; put data comparisons in `数据分析：`; put the summary in `实验结论：`; keep teacher review, grade, signature, and remarks areas blank. Allow the frame to extend across pages, but do not add body-report content outside the frame.
5. Convert a legacy `.doc` template to `.docx` as a working copy. Do not run red-to-black conversion on protected template content unless the user explicitly asks for template color normalization.
6. Add completed work at the designated location, inside the matching frame when frame-required mode applies, or immediately after the relevant preserved requirement when no frame is required. Apply formatting normalization only to newly generated content.
7. For a new report without a supplied template, run:
   `python scripts/create_starter_report.py report.docx`
8. Use `assets/table-format-reference.docx` for ordinary data tables and `assets/strict-starter.docx` as the clean report baseline when no user template is supplied.
9. Put every displayed code excerpt in a dedicated framed code box using `scripts/code_box.py`. Never place code as unframed body paragraphs.
10. Create or edit the DOCX using the Documents skill's deterministic workflow.
11. When a template was supplied, run:
   `python scripts/audit_template_preservation.py template.docx report.docx --json template-audit.json`
12. Run:
    `python scripts/audit_docx.py report.docx --json audit.json`
13. When source code is supplied, run:
    `python scripts/audit_code_consistency.py report.docx source-dir`
14. Fix every audit error. Render the DOCX to page images, inspect every page, and rerun the audits.

## Report Content Depth

- After choosing the normal or frame-required insertion mode, enrich the report content instead of delivering a screenshot-only or code-only document.
- For experiment reports, include installation/configuration context, data source and field descriptions, view construction steps, interaction logic, representative screenshots, key code excerpts, and explanations that connect code identifiers to the visible result.
- Keep the enrichment inside the designated template region. In frame-required mode, all added process detail, figures, code boxes, analysis, and conclusion text must remain inside the corresponding fill-in frames.
- Data analysis must compare at least two meaningful dimensions such as category, region, time, metric, filtered state, or interaction result. Conclusions should summarize both technical implementation gains and data-insight gains.
- Do not enrich protected template wording itself. Add detail only as generated content in permitted blank or fill-in regions.

## Supplied Template Preservation

- Preserve every existing template title and every paragraph stating experiment content, tasks, requirements, screenshot counts, or submission rules verbatim and in the original order.
- The first page of a supplied template is the highest-protection region. Preserve its text, fonts, sizes, color, bold/italic/underline, paragraph spacing, indentation, alignment, tables, borders, fills, and placement exactly as supplied.
- Protected template content includes all original paragraphs, tables, cells, frames, headers, footers, page breaks, and field placeholders. Never apply generated-content normalization to protected content.
- Do not replace template instructions with a newly invented report outline.
- Do not remove template wording after completing the corresponding section.
- Keep the template's original section order.
- Preserve original template colors. Convert red template text to black only when the user explicitly requests template color normalization.
- If the user explicitly requests complete template preservation including font size and color, use the original supplied template as the protected baseline; do not normalize protected template text to black, SimSun, Times New Roman, or body sizes.
- Fill genuine blank fields such as name, student number, class, and date when values are known; do not rewrite surrounding labels.
- Newly added headings must be semantically meaningful, but must not rename supplied template headings.
- Treat template outer frames, cover frames, and report fill-in frames as `SZU_TEMPLATE_FRAME` layout containers, not ordinary data tables.
- Preserve template frame borders and wording, but do not apply ordinary table header fills, alternating row fills, or data-table alignment rules to them.
- When frame-required mode applies, the report content belongs inside the existing fill-in frames. Do not create appendix-like body pages outside the frames to bypass available template space.

## Strict Formatting Contract

- These rules apply to newly generated content only. They do not apply to protected template content.
- Use automatic/default black text for newly added content.
- Set simplified Chinese text in newly added content to `SimSun`.
- Set Latin letters, numbers, and English text in newly added content to `Times New Roman`.
- Set newly added ordinary body text to exactly 10.5 pt / 五号.
- Use: title 15 pt / 小三, Heading 1 14 pt / 四号, Heading 2 12 pt / 小四, Heading 3 10.5 pt / 五号, Heading 4 9 pt / 小五.
- Apply a 21 pt two-character first-line indent to ordinary Chinese body paragraphs.
- Do not indent headings, captions, code boxes, ordinary data-table cell paragraphs, lists, references, or centered content. In frame-required report cells, ordinary generated prose still uses the required 21 pt first-line indent.
- Use Chinese parenthesized numbering such as `（1）` for body distinctions. Do not use decorative bullets.
- Use strict academic numbering for newly generated sections: first level `（一）`, second level `1.`, third level `（1）`, fourth level `①` only when necessary. Do not mix `1）`, loose bullets, or decorative symbols in generated academic prose.
- Center every figure and caption. Set captions to 10.5 pt.
- Format generated table and figure captions without internal blanks between the label and number, such as `表1实验结果统计` and `图1系统类图`. Do not write `表 1` or `图 1`.
- Put every displayed code excerpt in a dedicated bordered code box using the `CodeBox` paragraph style. Set all code-box text to exactly 9 pt / 小五.
- Keep exactly one normal blank line of visual separation before and after code boxes when possible. The explanation paragraph must immediately follow the code box after that single spacing gap.

## Code And Document Consistency

- Use the same module, class, interface, method, variable, and pattern-role names as the source code.
- Verify identifiers against supplied code before finalizing.
- Do not invent or silently translate identifiers.
- Use `scripts/code_box.py` for code excerpts. A code box is a framed single-cell container whose paragraphs use the `CodeBox` style, not an ordinary data table.
- Keep code boxes with their immediately preceding code heading and following explanation when pagination allows.
- Show only the key code needed to prove the implementation logic. Prefer concise excerpts of orchestration, validation, core algorithms, model training, and CLI usage over long imports, boilerplate, tests, or repeated setup.
- Add exactly one normal blank line of visual spacing before and after every code box when possible. Do not allow a table and code box to touch directly or drift apart.
- Put the detailed explanation immediately after each displayed code box, separated only by that single visual spacing gap. The explanation must be precise, content-rich, and tied to the code identifiers, data flow, design pattern role, input/output, and edge-case behavior shown in the excerpt.

## Tables And Figures

- Template/layout frames are not ordinary data tables. Do not isolate their rows with white/light-gray background fills; keep their cells unshaded unless the supplied template already requires otherwise.
- Apply the reference grid, white header, alternating white/light-gray rows, compact cell margins, and deliberate alignment to newly generated ordinary data tables.
- Do not apply ordinary table styling to `SZU_TEMPLATE_FRAME` or `SZU_CODE_BOX` containers.
- Keep tables centered and avoid fixed row heights.
- Add clear spacing before and after every generated table. When a table is followed by a code box, use the same single-line visual separation required for code boxes.
- Place figure captions immediately below figures and keep each figure-caption pair together.
- Keep table captions and figure captions compact: no blank line between the table/figure and its caption, and no extra spaces inside caption labels such as `表 1` or `图 1`.

## Delivery Gate

High strictness requires:

- No audit errors.
- Supplied template text passes `audit_template_preservation.py`.
- No red text remains in newly generated content. Protected template colors remain exactly as supplied unless the user explicitly requests template color normalization.
- Every displayed code excerpt is inside a framed `CodeBox` container and uses 9 pt text.
- Template/layout frames are marked or recognized as `SZU_TEMPLATE_FRAME` and are not shaded as ordinary tables.
- Heading levels are semantic and sequential for newly added content.
- Code identifiers match the source.
- All pages are rendered and visually inspected.
- No clipping, overlap, orphan headings, separated figures/captions, or accidental colored text.

Retain warnings only when they are demonstrably false positives, and explain them in the final response.
