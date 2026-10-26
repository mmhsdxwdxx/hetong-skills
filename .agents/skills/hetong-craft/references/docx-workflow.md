# docx 读写工作流

合同以 docx 为主。本技能的 `scripts/docx_tool.py`（纯 Python ≥3.8 标准库，
零依赖）用于普通无修订文档的读写；需要真实修订或批注时使用
[红线与批注流程](docx-redline.md)及独立的 `hetong_redline.py`。**禁止**用文本替换工具（sed 等）直接改 docx
二进制/XML——会踩 run 边界与命名空间的坑。

## 一、纪律（每次动手前默诵）

1. **先提取留档**：`extract` 存一份改动前文本，作为 diff 基准与回滚参考。
2. **备份**：`--in-place` 自动备份为 `文件名.bak.时间戳`；批量改动前也可
   手动 `cp` 整目录。
3. **先预演**：所有修改命令支持 `--dry-run`，先看不改。
4. **改后验证**：`verify`（zip/XML 结构）＋ 重新 `extract` 比对改动点 ＋
   `textutil -convert txt` 能正常解析（仅文本解析检查，不等于Word/WPS实机验证）；
   再用 `office_kit.py render` 做一次无头页面图/中文文本层检查，不要在每处修改后开 GUI。
5. **编号会漂移**：insert/delete/add-row 改变结构后，P/T 编号重排——
   结构改动后重新 `extract` 再定位。`batch` 例外：其编号以批处理开始
   前的状态统一解析。
6. **修订标记**：本工具不处理 w:ins/w:del，不得先接受或拒绝全部修订来绕过限制。
   先按 [红线与批注流程](docx-redline.md)检查已有修订；仅在用户明确确认形成新清洁基线后，
   才另存清洁副本交普通工具使用，历史版本不变。带修订文档的普通 extract 不作为完整视图。

## 二、读取

```bash
T=<技能目录>/scripts/docx_tool.py

python3 $T extract 合同.docx              # 打印全文（段落 P0001…，表格 T0001 R01 C01）
python3 $T extract 合同.docx -o 全文.txt  # 存文件
python3 $T extract 合同.docx --footers    # 连同页眉页脚
python3 $T search 合同.docx 关键词        # 定位段落（默认子串匹配）
python3 $T search 合同.docx '违约金|赔偿' -e   # 正则
python3 $T search 合同.docx --numbers     # 金额/税率/期限/违约金敏感段落
```

提取输出中 `P0012〔Heading1〕` 为段落样式名；表格单元格行同时给出 P 编号，
可直接对单元格段落做 `set`/`replace`。

## 三、修改（五把刀）

```bash
# 1) replace：跨 run 文本替换（措辞、数值、名称）
python3 $T replace 合同.docx 旧文本 新文本 --para P0012 --in-place
python3 $T replace 合同.docx 旧文本 新文本 --all --in-place   # 全文（先 dry-run！）
# 多段落命中而未加 --all/--para 时拒绝执行并列出命中段——这是防呆，照做。

# 2) set：整段改写（保留段落格式与字符格式）
python3 $T set 合同.docx P0012 --text "新文本" --in-place

# 3) insert-after：在 P0009 后插入新段（克隆 P0009 的格式；--clone 可另指格式来源）
python3 $T insert-after 合同.docx P0009 --text "4.　新增条款内容" --in-place

# 4) delete：删除段落（单元格内唯一段落不可删，用 set 清空代替）
python3 $T delete 合同.docx P0020 P0021 --in-place

# 5) add-row：表格增行（克隆参考行格式，--cells 用 | 分列）
python3 $T add-row 合同.docx T0002 --cells "6|备件|品牌X|2|SN-9|通过|无" --after R03 --in-place
```

**批量修改**（推荐多于 3 处时使用）——edits.json：

```json
[
  {"op": "replace", "old": "旧称", "new": "新称", "para": "P0012"},
  {"op": "replace", "old": "词A", "new": "词B", "all": true},
  {"op": "set", "para": "P0003", "text": "协议编号：XX-001"},
  {"op": "insert-after", "para": "P0015", "text": ["（新增款一）", "（新增款二）"]},
  {"op": "delete", "paras": ["P0004"]},
  {"op": "add-row", "table": "T0002", "cells": "…", "after": "R02"}
]
```

```bash
python3 $T batch 合同.docx edits.json --dry-run    # 预演
python3 $T batch 合同.docx edits.json --in-place   # 执行
```

批量语义：所有 P/T 编号按**执行前**的文档状态统一解析；replace 未命中
即整体中止（防呆）；任何一步失败不部分写盘。

## 四、新合同 docx 的生成路径

新起草默认先交付 Markdown 评审稿；定稿后转 docx：

1. 把 Markdown 写成带样式的 HTML（标题 h1/h2、正文 p、表格 table、
   加粗 strong、居中 style="text-align:center"）；
2. `textutil -convert docx 稿.html -output 合同.docx`；
3. `extract` 检查转换结果（表格、缩进、字体回退），必要时微调 HTML 重转。

若同族已有版式基准文件，更优路径：复制基准 docx 作为壳，用五把刀逐段
替换内容（保留其字体、页边距、页脚等版式）。

## 五、验证清单（改完必做）

```bash
python3 $T verify 合同.docx                       # 结构与统计
python3 $T extract 合同.docx -o 改后.txt && diff 改前.txt 改后.txt   # 差异即改动
textutil -convert txt 合同.docx -output /tmp/解析测试.txt            # 可解析性
```

diff 结果应**只含**预期改动；出现意外差异（尤其空段增减、表格行列变化）
必须查明原因再交付。

## 六、常见故障与回退

| 症状 | 原因与处置 |
|---|---|
| replace 报"未找到匹配文本" | 目标文本跨段落或被全角空格/特殊空格断开——先 `search` 看实际文本，缩小到单段逐段替换 |
| 多段命中被拒绝执行 | 防呆机制：逐段 `--para` 或确认后 `--all` |
| 改完 Word 报无法打开 | 立即从 `.bak.时间戳` 恢复：`cp 文件.bak.时间戳 文件.docx`；并把场景反馈给用户 |
| 段落格式丢失 | insert-after 未克隆对参照段；用 `--clone` 指定同格式段落 |
| 编号错乱 | 结构改动后用了旧编号——重新 `extract` 再定位 |
| 嵌套表格/图片/文本框 | 本工具不展开嵌套表格、不动图片与文本框；先停止批量改写，定位对象后按 [Office 精修与无头预览](office-workflow.md) 选择专用工具或请用户确认；复杂结构不得强行拆 XML |
