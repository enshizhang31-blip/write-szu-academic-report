# write-szu-academic-report

一个用于生成、补全、改写和审查深圳大学中文学术报告与实验报告的 Codex Skill。

它强调严格保留用户提供的 Word 模板，并为新增内容统一应用深大报告常见的字体、字号、代码框、图表和章节规范。仓库同时提供 DOCX 审计脚本、通用起始模板和表格格式参考。

## 主要能力

- 保留模板原文、结构、颜色、边框、文本框和填空区域
- 支持普通报告与框内填写报告两种模式
- 新增中文使用宋体，英文与数字使用 Times New Roman
- 代码片段使用带边框的 9 pt 代码框
- 审计模板保真度、报告格式和源码标识符一致性
- 提供可复用的起始 DOCX 与数据表样式参考

## 安装

将技能目录复制到 Codex 技能目录：

```powershell
Copy-Item -Recurse .\write-szu-academic-report "$HOME\.codex\skills\write-szu-academic-report"
```

也可以使用支持 GitHub 技能仓库的安装工具，从本仓库选择 `write-szu-academic-report` 子目录安装。

## 使用

在 Codex 中显式调用：

```text
$write-szu-academic-report 请严格保留这个实验报告模板，并根据项目源码完成报告。
```

技能会要求先检查模板、教程、源码和截图要求，再生成内容、执行审计并渲染检查最终 DOCX。

## 目录

```text
write-szu-academic-report/
├── SKILL.md
├── agents/openai.yaml
├── assets/
├── references/
└── scripts/
```

## 依赖

- Python 3.10+
- `python-docx`
- Codex 的 Documents 技能及其 DOCX 渲染工具（用于完整工作流）
- Microsoft Word 或 LibreOffice（处理旧版 `.doc` 模板时建议安装）

## 隐私说明

公开版本不包含带真实姓名、学号或个人作者元数据的示例报告。使用本技能处理报告前，也建议检查模板、截图和文档属性中的个人信息。

## 许可证

[MIT](LICENSE)

