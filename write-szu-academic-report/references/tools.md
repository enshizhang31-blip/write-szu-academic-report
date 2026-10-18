# 工具调用

从依赖加载器获取Python绝对路径，以下`python`表示该解释器。脚本路径以本技能根目录为基准，可用绝对路径调用。需要python-docx、lxml；渲染另需numpy、Pillow、pdf2image、pypdf、Poppler及WPS／Word或捆绑LibreOffice。工具不安装依赖、不访问网络。所有命令提供 --help，标准输出为JSON；支持 --json 的命令可另存结果，其余命令的JSON可用终端重定向保存。

## 检查和准备

```powershell
python scripts/report_tools.py --help
python scripts/report_tools.py doctor --json work/doctor.json
python scripts/report_tools.py doctor --probe --output-dir work/editor-probe --json work/doctor-probe.json
python scripts/report_tools.py profile template.docx --json work/profile.json
python scripts/report_tools.py outline template.docx --json work/template-outline.json
python scripts/report_tools.py prepare template.docx --output work/working.docx --manifest work/regions.json --layout-table 0 --field 15
```

示例索引不是通用模板索引。`--layout-table`是profile中从0开始的表格索引，可以重复；明确授权该表格的填写和布局变化。封面和教师区域不要指定为可填写表格。如果正文、教师栏共用一个外表格，工具保护原段落，但作者必须仅在正文格内添加内容，不能在教师格添加签字等信息。

`--field`是outline中从0开始的段落索引，可以重复；填写需保留原run，再插入已知值的新run，不得整体替换标签。`--parent-heading 22:1`将原段落22作为语义一级标题上下文，保留原文字和样式；仅在确为上位标题时使用，表单标签不指定。

`--page-frame`在prepare时授权修改节的页边框，format时再次传入才应用；默认不改变模板的节边框。应用后必须检查封面是否适合、是否与已有表框重叠。工作副本、区域清单和原模板必须一起保留。区域清单不是自动从已写完报告反推的可信保护基线。

## 内容和标题

正文布局表格先转换为正常段落；真实表单继续在授权单元格中填写。新段落不要复制原段落的`w14:paraId`，否则会被认为修改或复制原文。原段落标识不得删除。先完成内容移动，再重新提取outline，避免索引过期。

```powershell
python scripts/report_tools.py reflow work/working.docx --manifest work/regions.json --layout-table 0 --output work/flow.docx --manifest-out work/flow-regions.json --json work/reflow-result.json
```

reflow仅接收prepare已授权的布局表索引，不自动选择；合并、多列、浮动定位、嵌套布局包装表先阻塞待审。内部真实数据表和代码表原样保留，教师行保留为表单。原保护表记录标记retired并保留；段落与内部表格记录保存原容器与新位置。转换前后均审计。新增正文插在SZU_FLOW命名书签的起止段落之间，保留书签及所有原文记录；不能移到封面或教师区域。未识别用途的单列表通过人工检查确认后再授权。

旧兼容模式小于14会阻塞prepare；只有明确授权时传入 `--upgrade-compatibility`，在副本升级到15并记录旧settings与升级结果。禁止静默转换；转换后重新渲染。`format --generated-all`仅用于完全新建样稿，不得用于绕过报告保护。

```powershell
python scripts/report_tools.py outline work/draft.docx --manifest work/regions.json --json work/outline.json
```

outline的`paragraphs`每项包含`index`、`id`、`text`、`level`、`observed_level`、`protected`、`question`、`evidence`、`parent_index`。按正文含义编辑：非标题`level=null`；标题1至5级，填写本节要回答的问题`question`及父标题段落索引`parent_index`，一级的父级为null。`evidence`记录支撑该节的材料标识。已有保护标题的level只能取prepare记录的语义上下文，不能重设其外观。

必要时用`replacement_text`显式给出不含旧序号的新标题文字；工具只对新增标题使用该字段。`text`必须仍与输入段落精确相同，以检测过期提纲。不提供replacement_text时保留段落文字。已有手写标题编号必须移除，避免自动编号叠加。

提纲可以只包含标题，但必须包括全部要改级的标题和必要上位上下文。模型须自行审查父子主题、同级划分和标题正文一致性，脚本不声称理解正文。

```powershell
python scripts/report_tools.py outline work/draft.docx --manifest work/regions.json --plan work/outline.json --json work/outline-check.json
python scripts/report_tools.py format work/draft.docx --manifest work/regions.json --plan work/outline.json --output work/formatted.docx
```

全新生成的文档可以使用`--generated-all`替代manifest；不得对带有原模板内容的文档使用它。`normalize_docx.py`同样要求明确范围。`create_starter_report.py`生成的是未完成的起始稿，需替换所有提示文字，并据实际正文重新组织提纲。

```python
from code_box import add_code_box
# cell必须是已授权的模板正文单元格；说明段落紧随代码框。
add_code_box(cell, 'result = model.predict(image)\nreturn result')
cell.add_paragraph('说明关键标识符、输入输出与实际效果。')
```

## 审计与渲染

```powershell
python scripts/report_tools.py audit work/formatted.docx --manifest work/regions.json --source-dir source --evidence work/evidence.json --json work/audit.json
python scripts/report_tools.py render work/formatted.docx --output-dir work/render-v1/wps --backend wps --dpi 200 --roundtrip
python scripts/report_tools.py render work/formatted.docx --output-dir work/render-v1/word --backend word --dpi 200 --roundtrip
```

`--source-dir`和`--evidence`按任务材料提供；无源码时不伪称已核验代码。标识符匹配是启发式，外部库名称可能产生待审提示。evidence结构示例：

```json
{"requirements":[{"id":"R1","description":"展示预标注结果","status":"complete","needs_result":true,"evidence":[{"kind":"screenshot","path":"result.png"}]}]}
```

status为complete、incomplete或not_applicable。kind为source、log、raw_image、result、screenshot或manual_review。相对路径按证据清单目录解析。证据文件存在不能证明其内容真实匹配，需模型核对。未满足要求需在交付说明中明确列出。

渲染目录必须为空，避免混用旧页图。Windows先确认自动化实例文档数为0，再只读打开本工具的临时副本；通过该文档窗口句柄核对新建进程归属，确认唯一文档后才隐藏实例并导出。空白Word的应用窗口句柄可能不可用，应检查文档窗口。无法证明隔离即阻塞，不修改已有窗口、全局设置或关闭已有文档。COM沙箱会话隔离失败可通过环境审批执行同一操作。`--roundtrip`另存临时副本并重开，对比实际五级编号与层级。其他系统显式传`--backend libreoffice --soffice <捆绑可执行文件绝对路径>`；其结果不能替代WPS／Word验收。

渲染产生`report.pdf`、默认200 DPI的`page-N.png`、`render.json`、软件版本、实际编号与诊断日志。疑似断线、重复边框、裁切保存页码、像素区域及`risk-页-序号.png`；留白和孤立标题另作提示。自动检查可能把图形线条误判为表格，必须逐页复核；重叠、文字完整与外框闭合仍靠视觉确认。编辑辅助网格不打印，正文布局表格必须通过结构检查识别，不能只看PDF。

## 完整验收

```powershell
python scripts/report_tools.py verify work/formatted.docx --manifest work/regions.json --output-dir work/render-v1 --primary-editor wps --compat-editor word --reuse-render --json work/verify.json
```

verify在软件可用且无结构错误、审查未完成时返回3并生成`review-template.json`；缺软件返回2，结构错误返回1。审查后填写全局reviewer、semantic_pass、semantic_notes，以及editors.wps与editors.word下各自的pages_reviewed、visual_pass、visual_notes、checks_passed。checks_passed须包含closed_frames、continuous_table_grid、readable_figures、no_orphan_headings、figure_caption_pairs、whitespace_reviewed、thin_single_borders、no_body_layout_wrapper、clipping_overlap_reviewed。语义说明应包含具体章节关系；保护原文和新内容的编号重启、导航、换行缩进与正文完整均要确认。自动检测不能代填通过。

```powershell
python scripts/report_tools.py verify work/formatted.docx --manifest work/regions.json --output-dir work/render-v1 --reuse-render --review work/render-v1/review-template.json --json work/verify-final.json
```

review绑定DOCX及各软件版本、PDF、页图、编号和风险记录的render_sha256；修改后必须重渲染和重审。各软件的`warning_rules_reviewed`与`render_risk_rules_reviewed`分别列出已复核的结构警告和图像风险，并在notes解释处理结果。工具不会自动填通过。`delivery_ready`仅表示对应检查完成，实验要求不完整项仍须披露。缺任一必需软件实际检查时`dual_editor_verified=false`；`--compat-editor none`只用于用户明确选择单软件检查，不能称双软件通过。

退出码：0为当前命令完成（仅verify且delivery_ready=true代表完整验收）；1为结构错误；2为依赖、输入或运行阻塞；3为人工审查未完成。普通audit输出的semantic_review与visual_review保持pending，不代表最终交付合格。

兼容入口`audit_docx.py`转发到audit；`audit_template_preservation.py template.docx report.docx --manifest regions.json`使用区域基线，不再采用整表文字相等。`prepare_template_black.py`仅在用户明确要求改色时使用，不加入默认流程。
