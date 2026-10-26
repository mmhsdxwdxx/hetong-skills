#!/usr/bin/env python3
"""Deterministic support for hetong-review: baseline, numbers, arithmetic and deliverable gates.

Read-only against hetong inputs. Writes only an explicitly requested JSON/Markdown output.
No third-party packages, network calls, OCR, or legal conclusions.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, cast
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
VERSION = 1

# Windows 控制台默认编码可能不是 UTF-8；命令行输出统一为 UTF-8。
for _stream in (sys.stdout, sys.stderr):
    try:
        cast(Any, _stream).reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_data(data: object, output: Path | None, fmt: str = "json") -> None:
    if fmt == "markdown":
        text = render_number_markdown(data)  # type: ignore[arg-type]
    else:
        text = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
    if output:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(text, encoding="utf-8")
    else:
        print(text, end="")


def protect_inputs(output: Path | None, inputs: list[Path]) -> None:
    if not output:
        return
    destination = output.expanduser().resolve()
    protected = {p.expanduser().resolve() for p in inputs}
    if destination in protected:
        raise ValueError(f"输出不得覆盖输入文件：{destination}")


def xml_count(root: ET.Element, name: str) -> int:
    return len(root.findall(f".//w:{name}", NS))


def docx_inventory(path: Path) -> dict:
    result: dict = {"valid_container": False, "warnings": [], "parts": {}}
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            result["valid_container"] = True
            result["has_main_document"] = "word/document.xml" in names
            result["header_parts"] = [n for n in names if re.fullmatch(r"word/header[^/]*\.xml", n)]
            result["footer_parts"] = [n for n in names if re.fullmatch(r"word/footer[^/]*\.xml", n)]
            result["media_files"] = [n for n in names if n.startswith("word/media/") and not n.endswith("/")]
            result["embedded_files"] = [n for n in names if n.startswith("word/embeddings/") and not n.endswith("/")]
            totals = {
                "insertions": 0,
                "deletions": 0,
                "moves_from": 0,
                "moves_to": 0,
                "comments": 0,
                "footnotes": 0,
                "endnotes": 0,
                "textboxes": 0,
                "hidden_text": 0,
                "nested_tables": 0,
                "alt_chunks": 0,
            }
            tag_names = {
                "insertions": "ins",
                "deletions": "del",
                "moves_from": "moveFrom",
                "moves_to": "moveTo",
                "comments": "comment",
                "footnotes": "footnote",
                "endnotes": "endnote",
                "textboxes": "txbxContent",
                "hidden_text": "vanish",
                "alt_chunks": "altChunk",
            }
            for name in names:
                if not (name.startswith("word/") and name.endswith(".xml")):
                    continue
                try:
                    root = ET.fromstring(archive.read(name))
                except (ET.ParseError, KeyError) as exc:
                    result["parts"].setdefault(name, {})["parse_error"] = str(exc)
                    continue
                counts = {}
                for key, tag in tag_names.items():
                    count = xml_count(root, tag)
                    if count:
                        counts[key] = count
                        totals[key] += count
                nested = 0
                for table in root.findall(".//w:tbl", NS):
                    nested += len(table.findall(".//w:tbl", NS))
                if nested:
                    counts["nested_tables"] = nested
                    totals["nested_tables"] += nested
                if counts:
                    result["parts"][name] = counts
            result["counts"] = totals
    except (OSError, zipfile.BadZipFile, zipfile.LargeZipFile) as exc:
        result["error"] = str(exc)
        result["warnings"].append("无法读取 docx 容器；不得声称已完成全文审核。")
        return result

    counts = result.get("counts", {})
    if not result.get("has_main_document"):
        result["warnings"].append("缺少 word/document.xml；不能把该压缩包视为已完整读取的 docx 合同。")
    if any(counts.get(k, 0) for k in ("insertions", "deletions", "moves_from", "moves_to")):
        result["warnings"].append("检测到修订痕迹；须确认插入/删除/移动的接受状态，普通 extract 可能不能代表签署文本。")
    if result.get("header_parts") or result.get("footer_parts"):
        result["warnings"].append("存在页眉或页脚；须确认其文本已纳入提取和审核范围。")
    if counts.get("comments"):
        result["warnings"].append("检测到批注；须区分谈判批注与拟签署正文。")
    if counts.get("footnotes") or counts.get("endnotes"):
        result["warnings"].append("检测到脚注或尾注；须确认其文本已纳入提取和审核范围。")
    if counts.get("hidden_text"):
        result["warnings"].append("检测到隐藏文字标记；须确认其是否属于签署文本及是否需要纳入审查。")
    if counts.get("textboxes"):
        result["warnings"].append("检测到文本框；须另行核对其中的文字是否已提取。")
    if counts.get("nested_tables"):
        result["warnings"].append("检测到嵌套表格；须另行展开并核对单元格，不能依赖普通段落提取。")
    if result.get("media_files"):
        result["warnings"].append("存在图片/扫描内容；未做 OCR 前不得视为已审，需用图像读取或 OCR 并抽检。")
    if result.get("embedded_files") or counts.get("alt_chunks"):
        result["warnings"].append("存在嵌入对象或 altChunk；须列出并决定是否作为合同附件另行提取。")
    return result


def fingerprint(files: list[Path]) -> dict:
    entries: list[dict[str, object]] = []
    for file in files:
        path = file.expanduser().resolve()
        entry: dict[str, object] = {"path": str(path), "name": path.name}
        try:
            stat = path.stat()
            entry.update({
                "exists": path.is_file(),
                "size": stat.st_size,
                "mtime_ns": stat.st_mtime_ns,
                "suffix": path.suffix.lower(),
                "sha256": sha256(path),
            })
            if path.suffix.lower() in (".docx", ".dotx"):
                entry["docx_inventory"] = docx_inventory(path)
        except OSError as exc:
            entry.update({"exists": False, "error": str(exc)})
        entries.append(entry)
    return {
        "schema_version": VERSION,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "files": entries,
    }


def command_fingerprint(args: argparse.Namespace) -> int:
    protect_inputs(args.output, args.files)
    data = fingerprint(args.files)
    write_data(data, args.output)
    return 0 if all(x.get("exists") and x.get("sha256") for x in data["files"]) else 1


def command_readonly_check(args: argparse.Namespace) -> int:
    before = json.loads(args.manifest.read_text(encoding="utf-8"))
    old = {entry["path"]: entry for entry in before.get("files", [])}
    paths = [str(p.expanduser().resolve()) for p in (args.files or [Path(x) for x in old])]
    protect_inputs(args.output, [args.manifest, *(Path(name) for name in paths)])
    changed, missing, added, errors = [], [], [], []
    for name in paths:
        path = Path(name)
        if name not in old:
            added.append(name)
            continue
        if not path.is_file():
            missing.append(name)
            continue
        try:
            current = sha256(path)
            if current != old[name].get("sha256"):
                changed.append({"path": name, "before": old[name].get("sha256"), "after": current})
        except OSError as exc:
            errors.append({"path": name, "error": str(exc)})
    omitted = sorted(set(old) - set(paths))
    passed = not any((changed, missing, added, errors, omitted))
    result = {
        "passed": passed,
        "checked": len(paths),
        "changed": changed,
        "missing": missing,
        "added": added,
        "omitted_from_check": omitted,
        "errors": errors,
    }
    write_data(result, args.output)
    return 0 if passed else 1


DIGIT = re.compile(r"[0-9０-９]")
ZH_NUMBER = re.compile(
    r"(?:百分之|千分之|万分之)|"
    r"(?:[零〇一二三四五六七八九十百千万亿两壹贰叁肆伍陆柒捌玖拾佰仟萬億]+(?:元|万元|角|分|个月|年|日|天|%))"
)
FINANCIAL_CONTEXT = re.compile(
    r"合计|总价|单价|金额|税率|征收率|费率|比例|付款|价款|税额|价税|违约金|定金|"
    r"保证金|质保金|赔偿|限额|上限|期限|质保期|保证期间|工作日|自然日|第.{1,8}期"
)
LOCATOR = re.compile(r"\b(?:P\d{4,}|T\d{4,}(?:\s+R\d+\s+C\d+)?)\b")


def number_ledger(files: list[Path]) -> dict:
    records = []
    for file in files:
        path = file.expanduser().resolve()
        text = path.read_text(encoding="utf-8", errors="replace")
        for number, line in enumerate(text.splitlines(), 1):
            reasons = []
            if DIGIT.search(line):
                reasons.append("digit")
            if ZH_NUMBER.search(line):
                reasons.append("chinese-number")
            if FINANCIAL_CONTEXT.search(line):
                reasons.append("financial-context")
            if not reasons:
                continue
            locator = LOCATOR.search(line)
            records.append({
                "file": str(path),
                "line": number,
                "locator": locator.group(0) if locator else None,
                "reasons": reasons,
                "text": line,
            })
    return {
        "schema_version": VERSION,
        "notice": "高召回候选台账，不是数字结论。须逐条判断语义、独立复算并与大小写/附件/关联合同勾稽。",
        "files": [str(p.expanduser().resolve()) for p in files],
        "count": len(records),
        "records": records,
    }


def render_number_markdown(data: dict) -> str:
    lines = [
        "# 数字候选台账",
        "",
        f"> {data.get('notice', '')}",
        "",
        "| 文件 | 行号 | 提取定位 | 命中原因 | 原文 |",
        "|---|---:|---|---|---|",
    ]
    for row in data.get("records", []):
        text = str(row["text"]).replace("|", "\\|").replace("\n", " ")
        lines.append(f"| {Path(row['file']).name} | {row['line']} | {row.get('locator') or ''} | {', '.join(row['reasons'])} | {text} |")
    return "\n".join(lines) + "\n"


def command_numbers(args: argparse.Namespace) -> int:
    protect_inputs(args.output, args.files)
    data = number_ledger(args.files)
    write_data(data, args.output, "markdown" if args.markdown else "json")
    return 0


def decimal_value(raw: str) -> Decimal:
    try:
        value = Decimal(raw.replace(",", "").strip())
        if not value.is_finite():
            raise ValueError
        return value
    except (InvalidOperation, ValueError) as exc:
        raise argparse.ArgumentTypeError(f"不是有效有限十进制数字：{raw}") from exc


def decimal_rate(raw: str) -> Decimal:
    raw = raw.strip()
    try:
        if raw.endswith("%"):
            rate = Decimal(raw[:-1].strip()) / Decimal(100)
        else:
            rate = Decimal(raw)
        if not (Decimal(0) <= rate <= Decimal(1)):
            raise ValueError
        return rate
    except (InvalidOperation, ValueError) as exc:
        raise argparse.ArgumentTypeError("税率须为 0–1 的小数或 0%–100% 的百分比，例如 0.06 或 6%") from exc


def money(value: Decimal, places: int) -> Decimal:
    if not (0 <= places <= 6):
        raise ValueError("小数位数须在 0–6 之间")
    quantum = Decimal(1).scaleb(-places)
    return value.quantize(quantum, rounding=ROUND_HALF_UP)


def command_calc_tax(args: argparse.Namespace) -> int:
    rate = args.rate
    if (args.net is None) == (args.gross is None):
        raise ValueError("--net 与 --gross 必须且只能提供一个")
    if args.net is not None:
        net = args.net
        tax = net * rate
        gross = net + tax
        formula = "税额=不含税价×税率；含税价=不含税价+税额"
    else:
        gross = args.gross
        net = gross / (Decimal(1) + rate)
        tax = gross - net
        formula = "不含税价=含税价÷(1+税率)；税额=含税价-不含税价"
    result = {
        "formula": formula,
        "rounding": f"ROUND_HALF_UP to {args.places} decimal places",
        "rate": str(rate),
        "net": str(money(net, args.places)),
        "tax": str(money(tax, args.places)),
        "gross": str(money(gross, args.places)),
    }
    write_data(result, args.output)
    return 0


def command_sum(args: argparse.Namespace) -> int:
    total = sum(args.values, Decimal(0))
    result = {
        "values": [str(x) for x in args.values],
        "total": str(money(total, args.places)),
        "rounding": f"ROUND_HALF_UP to {args.places} decimal places",
    }
    if args.compare is not None:
        result["compare"] = str(args.compare)
        result["difference_total_minus_compare"] = str(money(total - args.compare, args.places))
    write_data(result, args.output)
    return 0


def command_tax_check(args: argparse.Namespace) -> int:
    protect_inputs(args.output, [args.ledger])
    data = json.loads(args.ledger.read_text(encoding="utf-8"))
    source_items = data.get("items")
    if not isinstance(source_items, list) or not source_items:
        raise ValueError("价税台账必须含非空 items 数组")
    items, discrepancies = [], []
    net_total = Decimal(0)
    tax_total = Decimal(0)
    gross_total = Decimal(0)
    for index, item in enumerate(source_items, 1):
        item_id = str(item.get("id", f"ITEM-{index}"))
        try:
            net = decimal_value(str(item["net"]))
            rate = decimal_rate(str(item["rate"]))
        except (KeyError, argparse.ArgumentTypeError) as exc:
            raise ValueError(f"{item_id} 必须提供合法 net 与 rate") from exc
        tax = net * rate
        gross = net + tax
        row = {
            "id": item_id,
            "net": str(money(net, args.places)),
            "rate": str(rate),
            "tax": str(money(tax, args.places)),
            "gross": str(money(gross, args.places)),
        }
        for key, calculated in (("tax", tax), ("gross", gross)):
            if key in item:
                try:
                    stated = decimal_value(str(item[key]))
                except argparse.ArgumentTypeError as exc:
                    raise ValueError(f"{item_id} 的 {key} 不是合法数字") from exc
                row[f"stated_{key}"] = str(stated)
                if money(stated, args.places) != money(calculated, args.places):
                    discrepancies.append(f"{item_id} 的 {key} 与 net×rate 复算不一致")
        items.append(row)
        net_total += net
        tax_total += tax
        gross_total += gross
    totals = {
        "net": str(money(net_total, args.places)),
        "tax": str(money(tax_total, args.places)),
        "gross": str(money(gross_total, args.places)),
    }
    result = {
        "passed": True,
        "items": items,
        "totals": totals,
        "rounding": f"ROUND_HALF_UP to {args.places} decimal places",
        "errors": [],
        "discrepancies": discrepancies,
        "items_match_stated": not discrepancies,
        "note": "只验证价税算术并报告差异；税率适用、交易拆分与税务合规仍须另行核验。合同或分项不一致是 finding，不自动令本工具失败。",
    }
    if "stated_total" in data:
        try:
            stated_total = decimal_value(str(data["stated_total"]))
        except argparse.ArgumentTypeError as exc:
            raise ValueError("stated_total 不是合法数字") from exc
        difference = gross_total - stated_total
        result["stated_total"] = str(money(stated_total, args.places))
        result["difference_gross_minus_stated"] = str(money(difference, args.places))
        result["hetong_total_matches"] = money(difference, args.places) == 0
    write_data(result, args.output)
    return 0


def parse_id_blocks(text: str, prefix: str) -> tuple[dict[str, str], list[str]]:
    pattern = re.compile(rf"^###\s+({re.escape(prefix)}-\d+)[^\n]*$", re.M)
    matches = list(pattern.finditer(text))
    blocks, duplicates = {}, []
    seen = set()
    for index, match in enumerate(matches):
        item_id = match.group(1)
        if item_id in seen:
            duplicates.append(item_id)
        seen.add(item_id)
        end_candidates = [m.start() for m in matches[index + 1:index + 2]]
        next_heading = re.search(r"^#{2,3}\s+", text[match.end():], re.M)
        if next_heading:
            end_candidates.append(match.end() + next_heading.start())
        end = min(end_candidates) if end_candidates else len(text)
        blocks[item_id] = text[match.start():end]
    return blocks, duplicates


def gate_result(
    kind: str,
    path: Path,
    errors: list[str],
    warnings: list[str],
    count: int,
    output: Path | None,
    extra: dict | None = None,
) -> int:
    result = {
        "type": kind,
        "file": str(path),
        "passed": not errors,
        "item_count": count,
        "errors": errors,
        "warnings": warnings,
    }
    if extra:
        result.update(extra)
    write_data(result, output)
    return 0 if not errors else 1


def finding_level(block: str) -> str | None:
    match = re.search(r"风险等级[^\n]*?S([1-4])", block)
    return f"S{match.group(1)}" if match else None


def check_level_summary(text: str, blocks: dict[str, str], final: bool) -> tuple[list[str], list[str], dict]:
    errors, warnings = [], []
    counts = {f"S{i}": 0 for i in range(1, 5)}
    ids = {f"S{i}": [] for i in range(1, 5)}
    for item_id, block in blocks.items():
        level = finding_level(block)
        if level:
            counts[level] += 1
            ids[level].append(item_id)
    heading = re.search(r"^###\s+风险分级统计[^\n]*$", text, re.M)
    if not heading:
        message = "缺少‘风险分级统计’章节，无法核对汇总与明细"
        (errors if final else warnings).append(message)
        return errors, warnings, {"level_counts": counts, "level_ids": ids}
    remainder = text[heading.end():]
    next_heading = re.search(r"^#{2,3}\s+", remainder, re.M)
    section = remainder[:next_heading.start()] if next_heading else remainder
    rows = []
    for line in section.splitlines():
        if line.lstrip().startswith("|"):
            rows.append([cell.strip() for cell in line.strip().strip("|").split("|")])
    for level in counts:
        row = next((cells for cells in rows if cells and cells[0].startswith(level)), None)
        if not row or len(row) < 2:
            message = f"{level} 缺少分级统计行"
            (errors if final else warnings).append(message)
            continue
        raw_count = row[1]
        if not raw_count.isdigit():
            message = f"{level} 统计数量未填写：{raw_count or '空'}"
            (errors if final else warnings).append(message)
            continue
        declared = int(raw_count)
        if declared != counts[level]:
            errors.append(f"{level} 统计不一致：汇总 {declared}，R-ID 明细 {counts[level]}")
        if len(row) >= 3:
            declared_ids = set(re.findall(r"R-\d+", row[2]))
            actual_ids = set(ids[level])
            if declared_ids and declared_ids != actual_ids:
                errors.append(f"{level} 统计 R-ID 列表与明细不一致：汇总 {sorted(declared_ids)}，明细 {sorted(actual_ids)}")
            elif declared > 0 and not declared_ids:
                warnings.append(f"{level} 有数量但未列对应 R-ID，无法核对名单")
    return errors, warnings, {"level_counts": counts, "level_ids": ids}


def amount_variants(raw: object) -> set[str]:
    value = Decimal(str(raw))
    variants = {str(raw), f"{value:.2f}", f"{value:,.2f}"}
    if value == value.to_integral_value():
        variants.update({str(int(value)), f"{int(value):,}"})
    return variants


def check_tax_summary(path: Path, text: str) -> tuple[list[str], dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    totals = data.get("totals", {})
    errors = []
    requirements = {
        "net": ("不含税",),
        "tax": ("税额",),
        "gross": ("含税",),
    }
    labels = {"net": "合计不含税金额", "tax": "合计税额", "gross": "合计含税金额"}
    for key, keywords in requirements.items():
        if key not in totals:
            errors.append(f"价税复算结果缺少 totals.{key}")
            continue
        variants = amount_variants(totals[key])
        found = False
        for line in text.splitlines():
            if not any(scope in line for scope in ("合计", "总额", "总计")):
                continue
            if all(keyword in line for keyword in keywords) and any(variant in line for variant in variants):
                found = True
                break
        if not found:
            errors.append(f"报告未明确列示{labels[key]}：{totals[key]}")
    if not data.get("passed"):
        errors.append("价税复算工具本身未通过，须先修正台账或计算")
    return errors, {"tax_totals": totals, "tax_discrepancies": data.get("discrepancies", [])}


def command_check_report(args: argparse.Namespace) -> int:
    protected = [args.file, args.tax_check] if args.tax_check else [args.file]
    protect_inputs(args.output, protected)
    text = args.file.read_text(encoding="utf-8", errors="replace")
    errors, warnings = [], []
    for heading in ("结论摘要", "审核边界", "详细审核发现", "法律依据与核验记录", "覆盖矩阵"):
        if not re.search(rf"^##+\s+.*{re.escape(heading)}", text, re.M):
            errors.append(f"缺少报告章节：{heading}")
    blocks, duplicates = parse_id_blocks(text, "R")
    if not blocks:
        errors.append("未找到以 '### R-编号' 开头的详细发现")
    for item in duplicates:
        errors.append(f"重复 R-ID：{item}")
    required = ("风险等级", "结论状态", "定位", "原文或缺口", "发现与推理", "对我方的具体影响", "依据", "复核/反方解释", "修订方向", "保护性条款需求")
    for item_id, block in blocks.items():
        for field in required:
            if field not in block:
                errors.append(f"{item_id} 缺少字段：{field}")
        level = finding_level(block)
        if not level:
            errors.append(f"{item_id} 未标 S1–S4 风险等级")
        if level in ("S1", "S2") and "第二路径" not in block:
            errors.append(f"{item_id} 为 {level}，但未记录第二路径复核")
        if args.final:
            risk_line = re.search(r"风险等级[^\n]*", block)
            if risk_line and len(set(re.findall(r"S[1-4]", risk_line.group(0)))) > 1:
                errors.append(f"{item_id} 风险等级未唯一确定")
    summary_errors, summary_warnings, extra = check_level_summary(text, blocks, args.final)
    errors.extend(summary_errors)
    warnings.extend(summary_warnings)
    if args.tax_check:
        tax_errors, tax_extra = check_tax_summary(args.tax_check, text)
        errors.extend(tax_errors)
        extra.update(tax_extra)
    if args.final:
        if re.search(r"- \[ \]", text):
            errors.append("仍有未完成的质量检查框")
        for marker in ("[合同名称及版本]", "[YYYY-MM-DD]", "[一句话结论]"):
            if marker in text:
                errors.append(f"仍有模板占位符：{marker}")
    return gate_result("review-report", args.file, errors, warnings, len(blocks), args.output, extra)


def suggested_text(block: str) -> str:
    match = re.search(r"-\s*\*\*建议文本\*\*\s*[：:]\s*", block)
    if not match:
        return ""
    remainder = block[match.end():]
    next_field = re.search(r"\n-\s*\*\*待确认变量\*\*", remainder)
    return remainder[:next_field.start()] if next_field else remainder


def check_clause_anchors(path: Path, blocks: dict[str, str]) -> tuple[list[str], list[dict]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    items = data.get("items")
    if not isinstance(items, list) or not items:
        raise ValueError("锚点文件必须含非空 items 数组")
    errors, records = [], []
    seen = set()
    for index, item in enumerate(items, 1):
        item_id = str(item.get("c_id", ""))
        if not item_id:
            raise ValueError(f"锚点第 {index} 项缺少 c_id")
        if item_id in seen:
            errors.append(f"锚点 C-ID 重复：{item_id}")
        seen.add(item_id)
        required = item.get("required", [])
        forbidden = item.get("forbidden", [])
        if not isinstance(required, list) or not isinstance(forbidden, list) or not all(isinstance(x, str) and x for x in [*required, *forbidden]):
            raise ValueError(f"{item_id} 的 required/forbidden 必须是非空字符串组成的数组")
        body = suggested_text(blocks.get(item_id, ""))
        missing = [anchor for anchor in required if anchor not in body]
        present = [anchor for anchor in forbidden if anchor in body]
        if item_id not in blocks:
            errors.append(f"锚点 {item_id} 在条款提案中不存在")
        elif not body:
            errors.append(f"{item_id} 无法提取建议文本用于锚点核验")
        for anchor in missing:
            errors.append(f"{item_id} 建议文本缺少确认锚点：{anchor}")
        for anchor in present:
            errors.append(f"{item_id} 建议文本出现禁止锚点：{anchor}")
        records.append({"c_id": item_id, "passed": item_id in blocks and not missing and not present and bool(body), "missing_required": missing, "present_forbidden": present})
    return errors, records


def command_check_clauses(args: argparse.Namespace) -> int:
    protected = [args.file, args.anchors] if args.anchors else [args.file]
    protect_inputs(args.output, protected)
    text = args.file.read_text(encoding="utf-8", errors="replace")
    errors, warnings = [], []
    blocks, duplicates = parse_id_blocks(text, "C")
    for heading in ("确认方案一致性", "逐句反查"):
        if not re.search(rf"^##+\s+.*{heading}", text, re.M):
            message = f"缺少‘{heading}’记录，无法核对已确认方案与正文"
            (errors if args.final else warnings).append(message)
    if not blocks:
        errors.append("未找到以 '### C-编号' 开头的条款提案")
    for item in duplicates:
        errors.append(f"重复 C-ID：{item}")
    required = ("保护目标", "建议文本", "待确认变量", "联动", "法律核验")
    for item_id, block in blocks.items():
        if not re.search(r"R-\d+", block):
            errors.append(f"{item_id} 未关联 R-ID")
        for field in required:
            if field not in block:
                errors.append(f"{item_id} 缺少字段：{field}")
    if args.final:
        if re.search(r"- \[ \]", text):
            errors.append("仍有未完成的质量检查框")
        if "[完整中文条款" in text:
            errors.append("仍有条款正文模板占位符")
        if "[逐项抄写已确认要素]" in text or "[逐项抄写确认原句]" in text:
            errors.append("确认方案一致性表仍为模板占位符")
        for marker in ("确认方案原句", "正文逐字对应"):
            if marker not in text:
                errors.append(f"确认方案一致性表缺少逐字核对列：{marker}")
    anchor_records = []
    if args.anchors:
        anchor_errors, anchor_records = check_clause_anchors(args.anchors, blocks)
        errors.extend(anchor_errors)
    return gate_result("clause-proposals", args.file, errors, warnings, len(blocks), args.output, {"anchor_checks": anchor_records})


def extract_quote_segments(block: str) -> tuple[list[str], bool]:
    lines = block.splitlines()
    for index, line in enumerate(lines):
        if "原文或缺口" not in line:
            continue
        remainder = "\n".join(lines[index + 1:])
        field = re.split(r"\n(?=\s*-\s*\*\*)", remainder, maxsplit=1)[0]
        groups, current = [], []
        for candidate in field.splitlines():
            stripped = candidate.lstrip()
            if stripped.startswith(">"):
                current.append(re.sub(r"^\s*>\s?", "", candidate))
            elif current:
                groups.append("\n".join(current).strip())
                current = []
        if current:
            groups.append("\n".join(current).strip())
        missing_claim = "缺失" in line or "未见" in line or "未约定" in line
        return [x for x in groups if x], missing_claim
    return [], False


def command_verify_quotes(args: argparse.Namespace) -> int:
    protect_inputs(args.output, [args.report, *args.sources])
    sources = {str(p.expanduser().resolve()): p.read_text(encoding="utf-8", errors="replace").replace("\r\n", "\n").replace("\r", "\n") for p in args.sources}
    report = args.report.read_text(encoding="utf-8", errors="replace").replace("\r\n", "\n").replace("\r", "\n")
    blocks, _ = parse_id_blocks(report, "R")
    records, missing_claims, errors = [], 0, []
    for item_id, block in blocks.items():
        quotes, missing_claim = extract_quote_segments(block)
        if not quotes:
            if missing_claim:
                missing_claims += 1
                records.append({"r_id": item_id, "kind": "missing-clause-claim", "passed": True})
            else:
                errors.append(f"{item_id} 没有可校验的逐字引文或明确缺失声明")
                records.append({"r_id": item_id, "kind": "quote", "passed": False, "error": "missing quote"})
            continue
        for quote in quotes:
            record = {"r_id": item_id, "kind": "quote", "quote": quote, "passed": True}
            if "…" in quote or "..." in quote:
                record.update({"passed": False, "error": "引文含省略号，无法逐字校验"})
            else:
                found = [name for name, text in sources.items() if quote in text]
                record["sources"] = found
                if not found:
                    record.update({"passed": False, "error": "未在任何冻结底稿中找到逐字引文"})
            if not record["passed"]:
                errors.append(f"{item_id}：{record['error']}")
            records.append(record)
    if not records:
        errors.append("报告中没有 R-ID 引文记录")
    result = {
        "passed": not errors,
        "report": str(args.report.expanduser().resolve()),
        "source_count": len(sources),
        "quote_count": sum(1 for x in records if x["kind"] == "quote"),
        "missing_clause_claims": missing_claims,
        "records": records,
        "errors": errors,
    }
    write_data(result, args.output)
    return 0 if not errors else 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    fp = sub.add_parser("fingerprint", help="文件哈希与 docx 覆盖风险清单")
    fp.add_argument("files", nargs="+", type=Path)
    fp.add_argument("-o", "--output", type=Path)
    fp.set_defaults(func=command_fingerprint)

    ro = sub.add_parser("readonly-check", help="以先前 manifest 验证原件未变")
    ro.add_argument("manifest", type=Path)
    ro.add_argument("files", nargs="*", type=Path)
    ro.add_argument("-o", "--output", type=Path)
    ro.set_defaults(func=command_readonly_check)

    num = sub.add_parser("numbers", help="高召回数字/财务上下文候选台账")
    num.add_argument("files", nargs="+", type=Path)
    num.add_argument("-o", "--output", type=Path)
    num.add_argument("--markdown", action="store_true")
    num.set_defaults(func=command_numbers)

    tax = sub.add_parser("calc-tax", help="用 Decimal 复算价税")
    tax.add_argument("--net", type=decimal_value)
    tax.add_argument("--gross", type=decimal_value)
    tax.add_argument("--rate", required=True, type=decimal_rate)
    tax.add_argument("--places", type=int, default=2)
    tax.add_argument("-o", "--output", type=Path)
    tax.set_defaults(func=command_calc_tax)

    ledger = sub.add_parser("tax-check", help="复算分项及合计不含税、税额、含税金额")
    ledger.add_argument("ledger", type=Path, help="含 items[{id,net,rate,tax?,gross?}] 与可选 stated_total 的 JSON")
    ledger.add_argument("--places", type=int, default=2)
    ledger.add_argument("-o", "--output", type=Path)
    ledger.set_defaults(func=command_tax_check)

    total = sub.add_parser("sum", help="用 Decimal 复算分项/付款总额及差异")
    total.add_argument("--values", nargs="+", required=True, type=decimal_value)
    total.add_argument("--compare", type=decimal_value)
    total.add_argument("--places", type=int, default=2)
    total.add_argument("-o", "--output", type=Path)
    total.set_defaults(func=command_sum)

    report = sub.add_parser("check-report", help="审核报告结构与 findings 字段门禁")
    report.add_argument("file", type=Path)
    report.add_argument("--final", action="store_true")
    report.add_argument("--tax-check", type=Path, help="tax-check 输出；同时核对报告是否明示三项价税合计")
    report.add_argument("-o", "--output", type=Path)
    report.set_defaults(func=command_check_report)

    clauses = sub.add_parser("check-clauses", help="保护条款提案结构与追溯门禁")
    clauses.add_argument("file", type=Path)
    clauses.add_argument("--final", action="store_true")
    clauses.add_argument("--anchors", type=Path, help="确认方案锚点 JSON；required/forbidden 逐字检查建议文本")
    clauses.add_argument("-o", "--output", type=Path)
    clauses.set_defaults(func=command_check_clauses)

    quotes = sub.add_parser("verify-quotes", help="报告引文与冻结底稿逐字核验")
    quotes.add_argument("report", type=Path)
    quotes.add_argument("sources", nargs="+", type=Path)
    quotes.add_argument("-o", "--output", type=Path)
    quotes.set_defaults(func=command_verify_quotes)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return args.func(args)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"review_kit: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
