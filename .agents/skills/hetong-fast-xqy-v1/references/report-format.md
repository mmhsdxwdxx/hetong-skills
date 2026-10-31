# Word交付规格与输入

## 固定排版（详细审核报告）

- A4纵向；上下边距1.8cm、左右2.0cm；不加封面、目录、装饰线或彩色风险卡。
- 标题18pt黑体居左；一级标题13pt黑体，二级标题11pt黑体；黑色。
- **正文11pt宋体，包括中文、数字和英文**；左对齐，1.25倍行距，段后5pt；标签可加粗，不整段加粗。无首行缩进，靠小标题与段间距形成层次。
- 注记、日期、依据、页码9.5pt宋体；页脚居中“第X页／共Y页”。不另加反复页眉。
- 标题与下一段同页、孤行控制；正文允许自然续页，不将整个长章节锁在一页。默认不人为指定第几项必须在第几页。
- 一般目标3—5页，现由报告承担详细分析和完整文字建议，可按实际复杂度增加；不因页数删掉必要建议，也不靠过小字体压页。
- 每项先现状和位置，再意见，有必要才放“备注／建议写法”；已满足／不适用的项目可只有简短现状。第1—8项顺序固定。
- 命中本方固定条款时，在对应要点内增加“甲方／乙方固定条款重点〔编号；符合情况〕”段，仅标签加粗，论述仍为11pt宋体；紧接现状／意见，不另放全文条款对照表。写本方要求、原文差异和处置，不只报编号。
- 框架外问题只保留明显重大且不重复的事项。依据与限制合并放末尾，全文法条、精算底稿和完整覆盖台账留在核验目录。
- 页面和Word结构全部可编辑；不把报告转换成截图塞进Word。

宋体和黑体必须通过实际渲染核验。DOCX使用标准字体族名SimSun（宋体）和SimHei（黑体），减少对本地化别名的依赖；这不是替换字体。Windows使用系统字体，缺失时检查简体中文补充字体或有授权的字体来源，不生成或套用Unix fontconfig；macOS/Linux可对LibreOffice子进程使用任务／用户数据内的fontconfig。不改全局字体配置、不下载不明字体，不将替代字体称为宋体。按[Windows安装检查](../../hetong-craft/references/windows-setup.md)保留样张和PDF实际字体结果；缺失时不宣称版式验收通过。

## 输入格式

生成器接受UTF-8 JSON，结构以 [example-report.json](../assets/example-report.json) 为准：

| 字段 | 内容 |
|---|---|
| `schema` | 固定`hetong-fast-xqy-v1/report-1` |
| `title`, `review_date`, `perspective` | 报告标题、YYYY-MM-DD、实际审核立场 |
| `stance` | 新任务填写`party`（甲方／乙方／我方）、`role`（buyer／seller）；可填本次明确的`categories`，缺省按本方库发现相关项，空数组表示不用固定条款；旧无固定标记输入仍兼容 |
| `basics` | `product`、`quantity`、`price`、`delivery`，可加`conclusion`；都是审阅后的事实文字 |
| `sections` | 严格7项，`id`为2—8；`status`为已具备／部分具备／缺失／待核实／不适用；必须有`location`、`observation`；可有`recommendation`、`sample` |
| `extras` | 可选；每项`title`、`location`、`issue`、`recommendation`，可有`sample` |
| `verification` | `coverage`、`limitations`、`pending`文字；可有`sources`数组，每项`title`、`url`、`checked_on` |
| `input_files` | 可选；原件路径字符串数组，相对路径以输入JSON所在目录为准，用于禁止覆盖原件；原件SHA及核验日志另存任务记录 |
| `issues` | 默认双交付必填，统一R编号及详细建议、短批注、语义锚点，见[双文件规则](paired-delivery.md) |
| `bundle` | 默认双交付的DOCX基线、SHA、作者、两份文件名，见双文件规则 |

`basics`、每项`sections`、每项`extras`均可添加`fixed_clauses`，格式为`[{"id":"M03","status":"部分符合","analysis":"本方要求……；合同差异……；意见……。"}]`。采购方使用当前M编号并由生成器补A来源，销售方使用B编号；若存在固定条款标记，`stance`必填。生成器拒绝跨方、未知、重复或超出明确类别范围的编号，但不自动判断条款是否适用。报告首行立场优先从`stance`生成，旧数据没有`stance`时才用`perspective`文字。详细规则见[立场匹配](stance-and-fixed-clauses.md)。

所有正文都是纯文本。只在法源`url`字段建立HTTP(S)链接；尖括号、`&`等按文字保留，不执行HTML或Markdown。输入不负责批准商业条件；`sample`一律显示“建议，未确认”，已同意的参数仍须能回溯用户指令。

## 命令行为

默认使用`build_bundle.py`生成两份Word，并用同一`issues`数据建立报告、批注和映射。事项`section`决定其在八项或框架外问题下的位置；建议修订、履约注意及待核实分开标记。新任务的详细修订建议放`issues.suggestion`，不要只藏在无编号`sample`中。

以下`build_report.py`只用于用户明确只要报告或兼容旧输入：

```bash
"$PY" scripts/build_report.py --input /绝对路径/快审数据.json --out /绝对路径/产出/合同名称_v1_XQY快审报告.docx --dry-run
"$PY" scripts/build_report.py --input /绝对路径/快审数据.json --out /绝对路径/产出/合同名称_v1_XQY快审报告.docx
```

成功同时生成同名`.md`、`.build.json`；后者记输入／输出SHA和结构检查，**不代表法源／实质／视觉验收通过**。已有输出拒绝覆盖，修改后用新版本文件名。拒绝写入`原件/`及输入原件路径。`--dry-run`不创建任何目录或文件。

生成后用hetong现有Office工具渲染两份文件。Codex先查看workspace dependencies实际返回的能力，不假定存在bundled LibreOffice；可用时复用，否则采用已验证的hetong隔离环境及本机LibreOffice。Windows快审生成器使用隔离环境的Scripts/python.exe，LibreOffice采用本机Windows可执行文件。批注稿保留原稿排版，不应用本页报告字体模板；原生批注另验锚点和历史。保留渲染、双向编号核验及原件未变记录。内部PDF、PNG、虚构样张仅作验收，不作为真实合同报告；PDF通过不等于Word/WPS原生显示及往返已通过。
