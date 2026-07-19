# High-Strictness Format Standard

## Priority

1. User instructions.
2. Supplied template wording, order, and original formatting.
3. Academic semantic structure.
4. Non-conflicting bundled samples.
5. Word defaults only when no rule exists.

## Typography Tokens

| Role | Chinese font | Latin font | Size | Color | Alignment | First-line indent |
|---|---|---|---:|---|---|---|
| Document title | SimSun | Times New Roman | 15 pt / 小三 | automatic black | centered | none |
| Heading 1 | SimSun | Times New Roman | 14 pt / 四号 | automatic black | left | none |
| Heading 2 | SimSun | Times New Roman | 12 pt / 小四 | automatic black | left | none |
| Heading 3 | SimSun | Times New Roman | 10.5 pt / 五号 | automatic black | left | none |
| Heading 4 | SimSun | Times New Roman | 9 pt / 小五 | automatic black | left | none |
| Body | SimSun | Times New Roman | 10.5 pt / 五号 | automatic black | justified or left | 21 pt |
| Figure caption | SimSun | Times New Roman | 10.5 pt / 五号 | automatic black | centered | none |
| Ordinary table text | SimSun | Times New Roman | 10.5 pt / 五号 | black | deliberate | none |
| Code box | SimSun | Times New Roman | 9 pt / 小五 | automatic black | left | none |

## Supplied Template Protection

- Treat a supplied template as authoritative content, not merely a visual reference.
- Preserve existing template titles, experiment content, tasks, requirements, screenshot-count instructions, and submission instructions verbatim and in their original order.
- Preserve protected template formatting exactly: font family, size, bold, italic, underline, color, spacing, indentation, alignment, borders, fills, table geometry, page breaks, and relative placement.
- Treat the first page as the strictest protected region. Do not change any existing first-page text, run formatting, paragraph formatting, table formatting, frame formatting, or layout.
- Do not rename, rewrite, summarize, delete, reorder, recolor, resize, restyle, shade, or normalize protected template text or tables.
- Insert completed content beneath or around the preserved instructions.
- Preserve original template colors. Change red template text to black only if the user explicitly asks for template color normalization.
- When the user asks for complete preservation including size and color, protected template text keeps the original supplied-template run formatting exactly; apply typography normalization only to newly added report content.
- Run `scripts/prepare_template_black.py` only when the user explicitly requests template color normalization. Always run `scripts/audit_template_preservation.py` before delivery.

## Protected Template Region And Generated Content Region

- Protected Template Region: all content that existed in the supplied template before editing, including the first page, headings, requirements, blank placeholders, tables, frames, headers, footers, and grading sections.
- Generated Content Region: only paragraphs, tables, figures, captions, code boxes, and explanations newly inserted by Codex.
- Apply black text, SimSun/Times New Roman normalization, body size, ordinary-table styling, code-box styling, and spacing rules only to Generated Content Region.
- Do not apply `normalize_docx.py`, global font normalization, ordinary-table styling, red-to-black conversion, or heading restyling to Protected Template Region.

## Frame-Required Reports

- Before writing into a supplied report template, inspect page 2 and later for large bordered fill-in frames, table cells, or text boxes. If those regions contain labels such as `实验过程及内容：`, `数据分析：`, `实验结论：`, `指导教师批阅意见`, `成绩评定`, or `备注`, classify the template as frame-required.
- If the user explicitly says the report must be completed in a frame/text box, or must not be supplemented outside the frame, classify it as frame-required even if automated inspection is uncertain.
- In frame-required mode, insert generated report content only inside the existing matching frame cells or text boxes. Do not append body-report pages, prose, figures, code boxes, analysis, or conclusions outside the template frame.
- Use `实验过程及内容：` for experiment steps, screenshots, figure captions, key code boxes, and code explanations. Use `数据分析：` for data comparison and interpretation. Use `实验结论：` for the final summary.
- Keep `指导教师批阅意见`, `成绩评定`, signatures, dates, and remarks blank unless the user explicitly provides those values.
- Allow the template frame to expand or continue across pages to fit content. Preserve the frame labels, borders, order, and protected template formatting.
- If no fill-in frame is present, use the normal supplied-template insertion workflow and place generated content at the designated location or immediately after the relevant preserved requirement.
- Frame-required mode changes only the layout container. All generated content still follows the normal font, size, color, first-line indent, code-box, figure, caption, audit, and render-review rules.

## Body Paragraphs

- Apply a 21 pt first-line indent to ordinary Chinese prose.
- Do not simulate indentation with spaces or tabs.
- Do not indent headings, captions, lists, code boxes, references, centered paragraphs, or table cells.
- Use strict academic numbering in newly generated content: `（一）` for the first generated level under a preserved template heading, `1.` for the second level, `（1）` for the third level, and `①` only when a fourth level is unavoidable.
- Do not mix academic numbering with `1）`, `1)`, dot bullets, hyphen bullets, or decorative symbols in generated prose.
- Do not use dot bullets, hyphen bullets, or decorative bullet symbols for body distinctions.

## Academic Headings

- Choose levels by meaning, not appearance.
- Do not skip levels.
- Preserve supplied template headings verbatim, even when their wording differs from preferred headings.
- Use no more than four numbered heading levels for newly added content.
- In Chinese experiment reports, prefer `（一）`, `1.`, `（1）`, `①` as the generated heading ladder unless the supplied template already mandates another ladder.

## Code Boxes

- Put every displayed code excerpt inside one dedicated framed code box created with `scripts/code_box.py`.
- Use a single-cell container whose paragraphs use the `CodeBox` style; it is a visual text box, not an ordinary data table.
- Use the `CodeBox` paragraph style for every paragraph inside the box.
- Set all code-box text to exactly 9 pt / 小五, with SimSun for Chinese and Times New Roman for Latin text.
- Use automatic black text, a thin black border, compact internal margins, and a white or very light gray fill.
- Do not apply first-line indentation inside code boxes.
- Do not split a short code box across pages. Keep it near its heading and explanation.
- Keep displayed code concise. Include only the decisive lines needed to explain orchestration, validation, algorithms, model training, or CLI usage.
- Add exactly one normal blank line of visual spacing before and after every code box when possible. If a generated table precedes or follows a code box, keep that same single-line separation so the blocks do not touch or drift apart.
- Put a detailed explanation immediately after every displayed code box, separated only by that single visual spacing gap. The explanation must name the key identifiers, state the inputs and outputs, explain the control/data flow, connect the excerpt to the intended design pattern or requirement, and mention important edge-case handling.

## Template/Layout Frames

- Treat supplied template outer frames, cover frames, and report fill-in frames as layout containers, not ordinary data tables.
- Mark generated or edited template frames with `SZU_TEMPLATE_FRAME` when possible.
- Preserve template frame wording, row order, borders, width, and placement.
- Do not apply ordinary table styling to template frames: no white header rule, no alternating white/light-gray rows, and no data-table alignment changes.
- Keep template frame cells unshaded unless the original supplied template explicitly used shading that the user wants preserved.
- Continue to enforce black text, SimSun Chinese text, Times New Roman Latin text, and required sizes for newly added content inside template frames; protected template text keeps its original formatting when the user requests exact preservation.
- In frame-required reports, template frames are the mandatory placement region for generated report content, not decorative containers. Keep newly added ordinary prose inside those frames first-line indented, while headings, captions, images, code boxes, lists, references, and centered content remain unindented.

## Ordinary Tables

- Use `assets/table-format-reference.docx` as the mandatory ordinary-table reference.
- Use A4 portrait geometry and sample-derived margins.
- Use a white header row and alternating white/light-gray body rows.
- Use single-line outer and internal grid borders.
- Keep short values centered, narrative cells left-aligned, and cells vertically centered.
- Never use fixed row heights that clip content.
- Exclude `SZU_TEMPLATE_FRAME` and `CodeBox` containers from ordinary-table styling.
- Apply ordinary-table styling only to newly generated tables. Never restyle supplied-template tables just to satisfy ordinary-table rules.
- Add spacing before and after every generated table, especially near code boxes.

## Figures And Captions

- Center images and captions.
- Place captions immediately below images.
- Use 10.5 pt automatic black caption text.
- Keep each figure-caption pair together.
- Use compact generated captions with no internal blank between label and number: `表1...`, `表3-1...`, `图1...`. Do not use `表 1`, `图 1`, or insert blank lines between a table/figure and its caption.
- Put table captions immediately below generated tables unless the supplied template or discipline explicitly requires table captions above tables.

## Delivery-Blocking Failures

- Supplied template titles, experiment content, requirements, screenshot counts, or submission instructions were modified, removed, reordered, or replaced.
- Any existing first-page template text or formatting was modified.
- Protected template content was font-normalized, recolored, resized, re-bolded, re-indented, shaded, or restyled without explicit user instruction.
- Red text remains in generated content, or protected red template text was changed without explicit user instruction.
- Displayed code is outside a framed code box.
- Code-box text is not exactly 9 pt / 小五.
- Ordinary body text is not 10.5 pt.
- Chinese text is not SimSun or Latin text is not Times New Roman.
- Non-default colored text appears.
- Template/layout frames were styled as ordinary data tables or given row background fills.
- Ordinary body paragraphs lack the required first-line indent.
- Heading levels are skipped or semantically inconsistent.
- Generated academic numbering mixes incompatible forms or uses nonstandard forms such as `1）`.
- Ordinary tables do not follow the reference format.
- Figures or captions are not centered.
- Generated captions contain internal blanks such as `表 1` or `图 1`, or a blank line separates a generated table/figure from its caption.
- Report identifiers conflict with supplied source code.
