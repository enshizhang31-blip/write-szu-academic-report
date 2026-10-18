# write-szu-academic-report

用于生成、补全、重构和审查深圳大学中文学术报告与实验报告的 Codex Skill。默认 **WPS 为主、Microsoft Word 兼容**，正文按学术逻辑组织，模板原文和实验事实受到保护。

## 主要能力

- 按正文内容建立最多五级标题，使用真实多级编号和 Word 大纲级别，下级随上级重启。
- 在授权区域填写或转换布局表，保留保护清单、原文、图片、代码和教师表单；正文不再由每段一行的大表格承载。
- 数据表使用0.5磅黑色单实线、白色表头及白／浅灰交替正文行；代码采用独立9磅代码框。
- 探测两款软件的实际导出能力，分别生成200 DPI页图，核验实际编号及保存重开稳定性。
- 结构审计、渲染风险提示与逐页审查共同验收；审查绑定当前文档和渲染哈希，文件修改后失效。

| 标题级别 | 序号 | 字体字号 | 加粗 | 段落左缩进 |
|---|---|---|---|---|
| 一级 | 一、二、三、 | 黑体14磅 | 是 | 0 |
| 二级 | (一)(二)(三) | 宋体12磅 | 否 | 0 |
| 三级 | 1.、2.、3. | 宋体12磅 | 否 | 24磅／两字符 |
| 四级 | (1)(2)(3) | 宋体12磅 | 否 | 24磅／两字符 |
| 五级 | ①②③ | 宋体12磅 | 否 | 24磅／两字符 |

不强制每篇报告使用全部五级；步骤、正文列举和图题不会因为带数字就自动升级为标题。

## 安装与调用

将仓库中的 `write-szu-academic-report` 子目录复制到 Codex 技能目录。更新已有安装前先备份。

```powershell
Copy-Item -Recurse .\write-szu-academic-report "$HOME\.codex\skills\write-szu-academic-report"
```

在 Codex 中调用：

```text
$write-szu-academic-report 请根据实验要求和真实材料重构这份报告，使用WPS为主、Word兼容，保留模板原文并完成逐页检查。
```

首次任务确认主要软件，同一任务已回答则沿用。未提供的个人信息、实验结果和教师签名不会猜填。

## 工具与依赖

统一入口为 `write-szu-academic-report/scripts/report_tools.py`。所有子命令提供 `--help` 和JSON结果。

| 命令 | 用途 |
|---|---|
| `doctor` | 探测Word／WPS注册、版本和实际自动化导出能力 |
| `profile` | 提取格式与布局特征，旧标题仅供诊断 |
| `prepare` | 创建工作副本、保护清单和授权填写区域 |
| `reflow` | 受控展开明确授权的正文布局表 |
| `outline` | 提取正文、建立并检查语义标题树 |
| `format` | 在允许区域应用标题、正文、表格和代码格式 |
| `audit` | 检查内容保护、样式、编号和结构风险 |
| `render` | 使用指定软件生成PDF、页图、编号及哈希 |
| `verify` | 汇总结构、双软件渲染和逐页审查结果 |

Python 3.10+；Python依赖见 [requirements.txt](requirements.txt)。渲染还需要Poppler（`pdftoppm`和`pdfinfo`可执行文件位于PATH中）。Windows的WPS／Word渲染依赖已安装、可用的COM自动化；其他环境可显式指定捆绑LibreOffice，但其结果不能替代WPS／Word验收。字体需在渲染机器可用。

```powershell
python -m pip install -r requirements.txt
python write-szu-academic-report/scripts/report_tools.py --help
python write-szu-academic-report/scripts/report_tools.py doctor --probe --output-dir work/probe --json work/doctor.json
python write-szu-academic-report/scripts/report_tools.py render report.docx --backend wps --dpi 200 --roundtrip --output-dir work/render/wps
python write-szu-academic-report/scripts/report_tools.py render report.docx --backend word --dpi 200 --roundtrip --output-dir work/render/word
```

渲染目录须为空。完整的授权区域、布局转换、提纲、审查记录和验收示例见 [工具说明](write-szu-academic-report/references/tools.md)，排版要求见 [格式规范](write-szu-academic-report/references/format-standard.md)。

退出码：0为当前命令完成；1为结构错误；2为依赖、输入或运行阻塞；3为审查未完成。仅 `verify` 的 `delivery_ready=true` 表示相应检查完成；实验材料缺项仍需披露。缺少必需软件实际检查时不能宣称双软件通过。

## 验证

```powershell
python -m unittest discover -s write-szu-academic-report/scripts -p "test_*.py" -v
```

本次发布运行31项测试全部通过，覆盖五级样式与重启定义、模板保护、布局转换、边框冲突、缺失软件、失效审查记录及WPS相关回归。

本机实际报告验证使用WPS 12.1.0.24034和Word 16.0.20326.20140：两者均为12页，逐页检查完成，报告实际使用的21个一级至三级标题在保存重开后编号稳定。这是特定报告与版本的实测，不代表所有版本或全部文档均已验证；第四、第五级的覆盖来自结构回归测试。自动图像检测只生成提示，不能自行填写“视觉通过”。

## 公开内容与许可证

公开版本只包含技能、工具、测试和通用DOCX模板，不包含个人实验报告、截图、保护清单或本机渲染产物。通用模板未填写真实姓名、学号或个人作者元数据。

[MIT](LICENSE)
