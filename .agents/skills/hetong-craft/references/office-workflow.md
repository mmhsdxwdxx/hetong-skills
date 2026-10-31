# Office 精修与无头预览

本流程用于合同、审查报告、谈判汇报等 Office 交付物。核心原则：
**Python 负责精确修改，LibreOffice 只负责批量转换/渲染，Word/WPS 只在关键交付前做一次
目标软件兼容性确认，不参与每处结构校验。**

## 一、工具边界

| 文件/任务 | 首选工具 | 说明 |
|---|---|---|
| 含修订/批注 DOCX 的精细阅读 | `word_reader.py` | 原文/接受/标记视图、修订时间线、批注线程；只读 |
| DOCX 清洁编辑 | `docx_tool.py` | 段落/表格定位、批量预演、结构校验 |
| DOCX 留痕修订/范围批注 | `hetong_redline.py` | 真实 OOXML 修订和批注；历史审阅只读 |
| PPTX 既有稿精修 | `office_kit.py pptx` | 文本、表格单元格、字体、几何、备注 |
| DOCX/PPTX/XLSX/PDF 结构盘点 | `office_kit.py inspect` | 只读生成定位用的结构化清单 |
| DOCX/PPTX/XLSX 转 PDF/页面图 | `office_kit.py render` | 批量调用无头 LibreOffice，自动缓存 |
| XLSX 精修 | 专用表格流程 | 当前本工具只检查/渲染，不承诺公式、图表安全的写入 |

`.doc/.ppt/.xls` 不直接编辑：先另存为无宏标准 OOXML，核对转换差异后再进入本流程。
`.docm/.pptm/.xlsm` 等宏文档默认拒绝。

## 二、一次会话的标准编排

### 0. 环境检查，一次即可

```bash
O=<技能目录>/scripts/office_kit.py
python3 $O doctor --full
```

`doctor` 检查隔离 Python、LibreOffice、PDF 渲染和中文字体；`--full` 还会真实生成并渲染
一份虚构 DOCX/PPTX，并从 PDF 文本层确认中文没有丢失。失败时不要继续批量办公任务。

### Windows 11

Windows使用PowerShell，不要求WSL。首次安装、升级、只检查及非标准路径统一按[Windows安装检查](windows-setup.md)执行；脚本保留虚构Word、PDF和页面图供复核，不刷新已有线程。已安装技能后，在项目根目录运行：

```powershell
# 缺少Python或LibreOffice时才允许安装系统前置软件
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .agents\skills\hetong-craft\scripts\setup_windows.ps1 -InstallPrerequisites
```

Windows 的隔离虚拟环境位于 `%LOCALAPPDATA%\contact-office\venv`，LibreOffice 默认从
`C:\Program Files\LibreOffice\program\soffice.exe` 发现；Windows 使用系统字体，不生成
Unix fontconfig 配置。若 Word/WPS 正打开目标文件，先关闭或复制到独立路径再处理。
非标准LibreOffice位置可用当前进程的`CONTACT_LIBREOFFICE`指定完整可执行文件路径；无效路径报错，不隐式回退。`doctor --full --artifacts <新目录> --out <检查.json>`可保留检查证据；普通`doctor --full`仍使用临时样张。

### 1. 一次盘点全部输入

```bash
python3 $O inspect 合同.docx 审阅汇报.pptx 附件.xlsx
```

PPTX 清单给出每页 `shape_id`、段落、字体、表格、备注、位置和源文件 SHA-256。
后续计划只使用这些稳定定位，不靠截图坐标猜对象。

### 2. 生成一份完整修改计划

多个修改合并为一个计划，不按段落反复启动进程。PPTX 计划：

```json
{
  "schema": "contact-pptx-plan/1",
  "source": "原件/审阅汇报.pptx",
  "source_sha256": "<inspect 返回的 64 位哈希>",
  "output": "产出/审阅汇报_v2_修订稿.pptx",
  "operations": [
    {"id": "O-001", "op": "replace_text", "slide": 3, "shape_id": 12,
     "paragraph": 2, "old": "30%", "new": "40%"},
    {"id": "O-002", "op": "set_table_cell", "slide": 4, "shape_id": 7,
     "row": 2, "column": 3, "text": "待确认"},
    {"id": "O-003", "op": "set_notes", "slide": 3,
     "text": "依据 R-004：付款比例以已确认方案为准。"}
  ]
}
```

计划中的相对路径以**计划 JSON 所在目录**为基准。输出不得等于源文件。

### 3. 预演，不写目标

```bash
python3 $O pptx apply 修改计划.json --dry-run
```

预演在内存/临时文件完成，随后重新打开临时结果，逐项验证：
- 幻灯片数量和尺寸未变；
- 只改计划指向的形状/表格/备注；
- 未指向的形状、媒体文件集合与内容未变；
- 目标内容与预演快照一致。

预演失败不得加 `--commit` 硬闯。

### 4. 已获授权且预演通过后一次提交

```bash
python3 $O pptx apply 修改计划.json --commit \
  --result 产出/修订执行记录.json \
  --render-dir 产出/_预览
```

`--commit` 是技术提交开关，不是用户授权；已有明确授权直接沿用，缺少授权时按 [共用边界](task-authorization.md) 处理。工具原子另存、重新打开验证；验证失败会删除本次输出而不是留下
半成品。源文件哈希前后必须一致，不支持 `--in-place`。已有输出需显式 `--overwrite`。

### 5. 一次批量渲染

全部文件修改完成后统一渲染，而不是改一处开一次 Word：

```bash
python3 $O render 产出/*.docx 产出/*.pptx --out 产出/_预览 --dpi 120
```

同一次调用中多个文件只启动 **一个** LibreOffice 无头进程。输出包括：
- `source.pdf`；
- `page-NNN.png`；
- 最多 9 页一张的 `contact-sheet-NN.png`；
- `render.json` 和目录级 `render-manifest.json`。

缓存键包含输入 SHA、页范围、DPI、LibreOffice 版本、字体签名和工具版本。输入与设置未变时
直接复用；PDF/页面图被篡改或删除会触发重渲染，不会把旧图冒充新稿。

### 6. 视觉与内容双重验收

先看联系表定位异常，再打开可疑页的原尺寸 PNG：
- 中文是否缺字、乱码、字体替代异常；
- 文本是否溢出文本框/表格、元素重叠、边距过小；
- 分页、表格跨页、页眉页脚是否异常；
- 页码/图表/备注是否与结构清单一致。

视觉通过不等于内容正确。金额、比例、期限、责任上限、交叉引用仍按合同规则复算；
PPTX 文字改动还要用 `inspect` 或 `pptx verify` 复核精确文本。

### 7. 目标软件最终确认

LibreOffice 与 Word/PowerPoint/WPS 的字体和排版并非完全一致。关键交付物在全部自动
检查通过后，用目标软件打开一次，检查“审阅/所有标记”、批注选区和最终分页；有新增修订时在临时副本验证接受/拒绝，有批注时验证实际锚点，
不需要在每个编辑步骤反复开关 GUI。

Word批注和修订的原生显示另按[审阅配色检查](native-review-display.md)验证；不得以正文加色模拟高亮。无法独立设置或核验时保留原生显示、披露限制，不因配色中断已授权的文档工作。

## 三、PPTX 操作语义

| op | 必填 | 说明 |
|---|---|---|
| `replace_text` | `slide,shape_id,old,new` | 默认要求 `old` 在目标中只出现一次；可加 `paragraph`、`occurrence` |
| `set_paragraph_text` | `slide,shape_id,paragraph,text` | 整段改写；多字符格式/域/链接需 `allow_style_loss:true` |
| `set_table_cell` | `slide,shape_id,row,column,text` | 行列从 1 开始；多行文本用 `\n` |
| `set_font` | `slide,shape_id` + 字体属性 | 支持 `font_name,size_pt,bold,italic,color_hex`，可加 `match` |
| `set_geometry` | `slide,shape_id` + 位置/尺寸 | 英寸；默认拒绝超出幻灯片；仅顶层形状 |
| `set_notes` | `slide,text` | 备注，`\n` 分行 |

`replace_text` 默认拒绝跨 run 替换，因为可能改变局部粗体/颜色。确认格式影响后加
`allow_cross_run:true`，新文字沿用起始 run 的格式并在结果中产生警告。表格中的文本替换
需加 `"table": true` 以及 `row/column`。

不支持：复制/删除/重排幻灯片、动画、SmartArt/图表内部对象、组合内部几何、跨段整块改写。
遇到这些对象应缩小修改范围或改用模板重新生成，不能拆散 XML 硬改。

## 四、效率规则

1. **按需检查**：环境或字体变化、渲染失败时运行 `doctor --full`；已有同环境有效证据时不重复，每个文件也不重复。
2. **一次盘点**：批量 `inspect` 全部输入。
3. **一文件一计划**：同一文件的多处修改合并预演和提交。
4. **一次批量渲染**：多个文件一次 LibreOffice 调用；仅对改动文件重渲染。
5. **缓存优先**：相同 SHA/DPI/字体直接复用，并在交付记录标注缓存命中。
6. **视觉分级**：先联系表，再看异常页原图；不要每页都做高成本人工检查。
7. **不持久驻留 Office 进程**：每个渲染批次一个隔离无头进程，崩溃不污染其他任务；
   相比每文件启动 GUI 已显著减少开销，且比长期 UNO 守护进程更容易复现和清理。

## 五、安全与保密

- 输入原件只读；渲染先复制到临时目录，不把 LibreOffice 指向原件写回。
- 拒绝宏文档、ZIP 路径穿越、符号链接、异常压缩比和超大解压包。
- 所有文件留在本机；`doctor`、`inspect`、`pptx`、`render` 均不上传合同。
- 缓存与预览放在任务产出目录或本机缓存，不提交仓库；真实合同不进入技能目录。
- 外部关系只做提示，不主动更新链接；不得在未确认时让文档访问外网资源。
- 交付须保留计划、预演/提交结果、输入 SHA 和渲染清单，便于复核。

## 六、安装与修复

Python 依赖由锁定文件管理：

```bash
python3 <技能目录>/scripts/office_kit.py setup
# 网络需要镜像时：
python3 <技能目录>/scripts/office_kit.py setup \
  --index-url https://pypi.tuna.tsinghua.edu.cn/simple
```

LibreOffice须按实际操作系统安装：Windows版使用官方安装器或上述PowerShell入口；`brew install --cask libreoffice`仅适用于macOS。宿主提供bundled Python并不意味着同时提供LibreOffice，必须核对实际可执行文件。中文预览需要系统可用的CJK字体；
缺失时 `doctor --full` 会失败并明确提示。本工具自动生成**仅供 LibreOffice 子进程使用**的
fontconfig配置仅用于macOS/Linux的LibreOffice子进程，不覆盖用户全局字体设置。Windows直接使用系统字体，并移除该子进程继承的Unix字体配置变量。

任何环境变更后重新运行 `doctor --full`；版本、字体和缓存签名都会变化，旧渲染缓存自动失效。
