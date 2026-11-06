#!/usr/bin/env python3
"""两份近似 DOCX 的暗改检测：区分留痕修改与直接改正文。

检测口径：
- 后一版“拒绝全部修订”的文本 vs 前一版基线文本 → 未被后一版修订记录的改动；
- 后一版“接受全部修订” vs 其“拒绝全部修订” → 可由后一版修订记录解释的改动；
- 前一版修订/批注在后一版消失 → 审阅历史清除或重写；
- 图片/嵌入对象/隐藏文字/设置变化 → 文字修订不能解释的非文本变化。

本工具只读，不自动接受或拒绝修订，不修改任何输入文件。
"""

from __future__ import annotations

import argparse
import csv
import difflib
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any, cast

SCRIPT_DIR = Path(__file__).resolve().parent
CRAFT_SCRIPTS = SCRIPT_DIR
if str(CRAFT_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(CRAFT_SCRIPTS))

try:
    import importlib.util as _importlib_util

    _spec = _importlib_util.spec_from_file_location(
        "hetong_word_reader", CRAFT_SCRIPTS / "word_reader.py"
    )
    if _spec is None or _spec.loader is None:
        raise ImportError("spec is empty")
    wr = _importlib_util.module_from_spec(_spec)
    sys.modules[_spec.name] = wr
    _spec.loader.exec_module(wr)
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"无法加载 word_reader.py：{exc}") from exc

SCHEMA = "contact-hidden-change/1"

# Windows 控制台默认编码可能不是 UTF-8；命令行输出统一为 UTF-8。
for _stream in (sys.stdout, sys.stderr):
    try:
        cast(Any, _stream).reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass
DEFAULT_SIMILARITY = 0.55
MEDIA_PREFIXES = ("word/media/", "word/embeddings/", "word/charts/")


class HiddenChangeError(RuntimeError):
    pass


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    return text.strip()


def group_for_flow(flow: str) -> str:
    parts = (flow or "").split("|")
    story = parts[0] if parts else "word/document.xml"
    if "table" in parts:
        index = parts.index("table")
        if index + 1 < len(parts):
            return f"{story}|table|{parts[index + 1]}"
    if "textbox" in parts:
        index = parts.index("textbox")
        if index + 1 < len(parts):
            return f"{story}|textbox|{parts[index + 1]}"
    return f"{story}|body"


def blocks_by_group(blocks: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    result: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for block in blocks:
        if not block.get("text"):
            continue
        result[group_for_flow(block.get("flow", ""))].append(block)
    return result


def diff_text_groups(
    before_blocks: list[dict[str, Any]],
    after_blocks: list[dict[str, Any]],
    side: str,
) -> list[dict[str, Any]]:
    before_groups = blocks_by_group(before_blocks)
    after_groups = blocks_by_group(after_blocks)
    records: list[dict[str, Any]] = []
    for group in sorted(set(before_groups) | set(after_groups)):
        before = before_groups.get(group, [])
        after = after_groups.get(group, [])
        before_texts = [normalize(block["text"]) for block in before]
        after_texts = [normalize(block["text"]) for block in after]
        matcher = difflib.SequenceMatcher(None, before_texts, after_texts, autojunk=False)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                continue
            old = "\n".join(before_texts[i1:i2])
            new = "\n".join(after_texts[j1:j2])
            locations = []
            for block in after[j1:j2] or before[i1:i2]:
                if block.get("locations"):
                    locations.extend(block["locations"])
                elif block.get("location"):
                    locations.append(block["location"])
            paragraph_ids = []
            for block in after[j1:j2] or before[i1:i2]:
                paragraph_ids.extend(block.get("paragraph_ids", []))
            clauses = [
                block.get("clause_path") or block.get("clause_id")
                for block in (after[j1:j2] or before[i1:i2])
                if block.get("clause_path") or block.get("clause_id")
            ]
            records.append({
                "side": side,
                "kind": tag,
                "group": group,
                "location": locations[0] if locations else group,
                "locations": list(dict.fromkeys(locations)),
                "paragraph_ids": list(dict.fromkeys(paragraph_ids)),
                "clause": clauses[0] if clauses else None,
                "before": old,
                "after": new,
            })
    return records


def revision_signature(revision: dict[str, Any]) -> tuple[Any, ...]:
    return (
        revision.get("type"),
        revision.get("author"),
        revision.get("date_raw"),
        normalize(revision.get("text") or ""),
        revision.get("location"),
    )


def comment_signature(comment: dict[str, Any]) -> tuple[Any, ...]:
    return (
        comment.get("author"),
        comment.get("date_raw"),
        normalize(comment.get("accepted_text") or ""),
        comment.get("clause_id"),
    )


def package_media(path: Path) -> dict[str, str]:
    hashes: dict[str, str] = {}
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith(MEDIA_PREFIXES):
                hashes[name] = hashlib.sha256(archive.read(name)).hexdigest()
    return hashes


def hidden_blocks(reader: Any) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for paragraph in reader.paragraphs:
        if not paragraph.get("hidden_text"):
            continue
        result.append({
            "text": paragraph["hidden_text"],
            "location": paragraph.get("location"),
            "flow": paragraph.get("flow"),
            "paragraph_ids": [paragraph.get("id")],
            "clause_path": paragraph.get("clause_path"),
        })
    return result


def load_document(path: Path, include_hidden: bool) -> tuple[Any, dict[str, Any]]:
    if path.suffix.lower() != ".docx":
        raise HiddenChangeError("仅接受 .docx；旧版 .doc/.wps 请先另存并核对")
    if not path.is_file():
        raise HiddenChangeError(f"文件不存在：{path}")
    before_hash = sha256_file(path)
    reader = wr.WordReader(path, include_hidden=include_hidden)
    reader.load()
    if sha256_file(path) != before_hash:
        raise HiddenChangeError(f"分析期间文件发生变化：{path.name}")
    return reader, {
        "path": str(path.resolve()),
        "sha256": before_hash,
        "size": path.stat().st_size,
        "core": reader.core,
        "settings": reader.settings,
        "counts": {
            "revisions": len(reader.revisions),
            "comments": len(reader.comments),
            "paragraphs": len(reader.paragraphs),
        },
    }


def add_finding(
    findings: list[dict[str, Any]],
    category: str,
    severity: str,
    location: str | None,
    before: str | None,
    after: str | None,
    evidence: str,
    action: str,
    confidence: str = "高",
) -> None:
    findings.append({
        "id": f"HC-{len(findings) + 1:03d}",
        "category": category,
        "severity": severity,
        "location": location,
        "before": before,
        "after": after,
        "evidence": evidence,
        "action": action,
        "confidence": confidence,
    })


def substantive(text: str) -> bool:
    return bool(re.search(r"[0-9０-９一二三四五六七八九十百千万亿%％元日天月年台套个]", text or ""))


def compare_documents(
    before_path: Path,
    after_path: Path,
    output_dir: Path | None = None,
    *,
    baseline_view: str = "accepted",
    include_hidden: bool = False,
    similarity_threshold: float = DEFAULT_SIMILARITY,
    force: bool = False,
) -> dict[str, Any]:
    if baseline_view not in {"accepted", "original"}:
        raise HiddenChangeError("baseline_view 只能是 accepted 或 original")
    if not 0 <= similarity_threshold <= 1:
        raise HiddenChangeError("similarity_threshold 必须在 0—1 之间")
    before_reader, before_info = load_document(before_path, include_hidden)
    after_reader, after_info = load_document(after_path, include_hidden)
    before_blocks = before_reader.views[baseline_view]
    after_accepted = after_reader.views["accepted"]
    after_original = after_reader.views["original"]
    before_visible = "\n".join(normalize(block["text"]) for block in before_blocks)
    after_visible = "\n".join(normalize(block["text"]) for block in after_accepted)
    similarity = difflib.SequenceMatcher(
        None, before_visible, after_visible, autojunk=False
    ).ratio()

    total_changes = diff_text_groups(before_blocks, after_accepted, "total")
    untracked_changes = diff_text_groups(before_blocks, after_original, "untracked")
    tracked_changes = diff_text_groups(after_original, after_accepted, "tracked")

    findings: list[dict[str, Any]] = []
    for change in untracked_changes:
        severity = "高" if substantive(change["before"] + change["after"]) else "中"
        add_finding(
            findings,
            "untracked_text_change",
            severity,
            change["location"],
            change["before"],
            change["after"],
            "后一版拒绝全部修订后的文本与前一版基线不一致，说明该变化未由后一版修订记录解释。",
            "逐字核对业务含义；要求对方说明修改人、时间和依据，或在下一版中用真实修订补记。",
            "高",
        )
    if not untracked_changes:
        add_finding(
            findings,
            "no_untracked_text_change",
            "低",
            None,
            None,
            None,
            "未检测到拒绝修订文本与前一版基线的文字差异。",
            "仍需核对修订历史、批注、对象和版式；本结果不等于无风险。",
            "中",
        )

    before_revisions = [revision_signature(item) for item in before_reader.revisions]
    after_revisions = [revision_signature(item) for item in after_reader.revisions]
    removed_revisions = [item for item in before_revisions if item not in after_revisions]
    if removed_revisions:
        add_finding(
            findings,
            "review_history_removed",
            "中",
            None,
            None,
            None,
            f"前一版有 {len(removed_revisions)} 条修订在后一版中找不到相同记录；可能已接受、清理或被重写。",
            "核对是否为正常接受修订；若不是，要求提供原修订稿和接受/删除记录。",
            "中",
        )
    before_comments = [comment_signature(item) for item in before_reader.comments.values()]
    after_comments = [comment_signature(item) for item in after_reader.comments.values()]
    removed_comments = [item for item in before_comments if item not in after_comments]
    if removed_comments:
        add_finding(
            findings,
            "comments_removed",
            "中",
            None,
            None,
            None,
            f"前一版有 {len(removed_comments)} 条批注在后一版中找不到相同记录。",
            "确认批注是已解决、正常删除，还是审阅历史被清理；未解决的问题应保留/重建。",
            "中",
        )

    before_media = package_media(before_path)
    after_media = package_media(after_path)
    if before_media != after_media:
        changed = sorted(set(before_media) | set(after_media))
        add_finding(
            findings,
            "non_text_media_change",
            "中",
            None,
            None,
            None,
            f"图片、嵌入对象或图表部件发生变化（{len(changed)} 个媒体路径）；文字修订不能解释此类变更。",
            "逐项核对替换前后的图片/对象及图注；确认是否经授权。",
            "高",
        )

    before_hidden = hidden_blocks(before_reader)
    after_hidden = hidden_blocks(after_reader)
    hidden_diff = diff_text_groups(before_hidden, after_hidden, "hidden")
    for change in hidden_diff:
        add_finding(
            findings,
            "hidden_text_change",
            "高",
            change["location"],
            change["before"],
            change["after"],
            "隐藏文字或样式隐藏内容发生变化；默认阅读视图不会显示该内容。",
            "在 Word/WPS 中开启隐藏文字/显示格式复核；确认是否属于有效约定。",
            "高",
        )

    if after_info["settings"].get("track_changes") is not None and not after_info["settings"].get("track_changes"):
        add_finding(
            findings,
            "track_changes_off",
            "低" if not untracked_changes else "中",
            None,
            None,
            None,
            "后一版当前跟踪修改开关为关闭；不能据此判断历史改动是否都留有痕迹。",
            "结合本报告的未留痕差异和修订历史判断，不把开关关闭本身当作恶意证据。",
            "高",
        )

    if similarity < similarity_threshold:
        add_finding(
            findings,
            "low_similarity",
            "中",
            None,
            None,
            None,
            f"两版可见文本相似度仅 {similarity:.3f}，可能不是同一合同族或版本。",
            "先确认文件关系和版本基线；不要用本报告的差异项直接判断暗改。",
            "中",
        )

    if untracked_changes and similarity >= similarity_threshold:
        severity = "高"
    else:
        severity = "中"
    for change in total_changes:
        if any(
            change["group"] == untracked["group"]
            and change["location"] == untracked["location"]
            for untracked in untracked_changes
        ):
            continue
        # 总差异中未被未留痕差异覆盖的部分，通常可由修订记录解释；
        # 仍保留在 JSON 中供对照，不单独作为暗改 finding。
    result = {
        "schema": SCHEMA,
        "ok": True,
        "before": before_info,
        "after": after_info,
        "baseline_view": baseline_view,
        "similarity": round(similarity, 6),
        "similarity_threshold": similarity_threshold,
        "counts": {
            "total_text_changes": len(total_changes),
            "untracked_text_changes": len(untracked_changes),
            "tracked_text_changes": len(tracked_changes),
            "history_removed": len(removed_revisions),
            "comments_removed": len(removed_comments),
            "hidden_text_changes": len(hidden_diff),
            "media_changes": 1 if before_media != after_media else 0,
            "findings": len(findings),
        },
        "severity": severity if untracked_changes else "低",
        "findings": findings,
        "untracked_changes": untracked_changes,
        "tracked_changes": tracked_changes,
        "total_changes": total_changes,
        "hidden_changes": hidden_diff,
        "media_changes": {
            "before": before_media,
            "after": after_media,
        },
    }
    if output_dir is not None:
        result["outputs"] = write_outputs(result, output_dir, force=force)
    return result


def render_markdown(result: dict[str, Any]) -> str:
    before = result["before"]
    after = result["after"]
    counts = result["counts"]
    lines = [
        "# 合同版本暗改检测报告",
        "",
        "> 本报告只读生成。“未留痕改动”指后一版拒绝全部修订后的文本与前一版基线不一致，",
        "> 不等于已经证明存在恶意；必须结合修改权限、邮件/审批记录和业务背景确认。",
        "> 正常的接受全部修订、清理历史批注也可能造成历史减少，不能单凭数量认定为暗改。",
        "",
        "## 一、版本与基线",
        "",
        f"- 前一版：`{Path(before['path']).name}`",
        f"  - SHA-256：`{before['sha256']}`",
        f"  - 修订/批注：{before['counts']['revisions']} / {before['counts']['comments']}",
        f"- 后一版：`{Path(after['path']).name}`",
        f"  - SHA-256：`{after['sha256']}`",
        f"  - 修订/批注：{after['counts']['revisions']} / {after['counts']['comments']}",
        f"- 基线视图：前一版 `{result['baseline_view']}`；后一版以“拒绝全部修订”作为未留痕基线。",
        f"- 可见文本相似度：`{result['similarity']:.3f}`（阈值 {result['similarity_threshold']}）。",
        "",
        "## 二、结论摘要",
        "",
        f"- 总文字变化：**{counts['total_text_changes']}** 处。",
        f"- 其中未留痕文字改动：**{counts['untracked_text_changes']}** 处。",
        f"- 可由后一版修订记录解释的变化：**{counts['tracked_text_changes']}** 处。",
        f"- 前一版修订记录减少：**{counts['history_removed']}** 条；批注减少：**{counts['comments_removed']}** 条。",
        f"- 隐藏文字变化：**{counts['hidden_text_changes']}**；媒体/对象变化：**{counts['media_changes']}**。",
        f"- 总体提示等级：**{result['severity']}**。",
        "",
        "## 三、未留痕文字改动（重点核对）",
        "",
        "| 编号 | 等级 | 位置 | 前一版 | 后一版 | 证据 |",
        "|---|---|---|---|---|---|",
    ]
    untracked_findings = [
        item for item in result["findings"] if item["category"] == "untracked_text_change"
    ]
    for item in untracked_findings[:200]:
        lines.append(
            "| {id} | {severity} | {location} | {before} | {after} | {evidence} |".format(
                id=item["id"], severity=item["severity"],
                location=(item.get("location") or "—").replace("|", "\\|"),
                before=(item.get("before") or "—").replace("|", "\\|").replace("\n", " ")[:600],
                after=(item.get("after") or "—").replace("|", "\\|").replace("\n", " ")[:600],
                evidence=item["evidence"].replace("|", "\\|"),
            )
        )
    if not untracked_findings:
        lines.append("| — | — | — | — | — | 未检测到与基线不一致的未留痕文字改动。 |")
    if len(untracked_findings) > 200:
        lines.append("")
        lines.append(f"> 仅展示前 200 条；完整结果见 `hidden-change.json`（共 {len(untracked_findings)} 条）。")
    lines.extend([
        "",
        "## 四、审阅历史、批注与非文本变化",
        "",
        "| 编号 | 类别 | 等级 | 证据 | 建议 |",
        "|---|---|---|---|---|",
    ])
    for item in result["findings"]:
        if item["category"] == "untracked_text_change":
            continue
        lines.append(
            f"| {item['id']} | {item['category']} | {item['severity']} | "
            f"{item['evidence'].replace('|', chr(92) + '|')} | "
            f"{item['action'].replace('|', chr(92) + '|')} |"
        )
    if not [item for item in result["findings"] if item["category"] != "untracked_text_change"]:
        lines.append("| — | — | — | 未检测到历史/非文本变化。 | — |")
    lines.extend([
        "",
        "## 五、可由后一版修订记录解释的变化",
        "",
        "以下变化来自后一版“拒绝全部修订”与“接受全部修订”的差异，属于后一版修订记录可解释的范围；",
        "仍应逐条核对条文语义、金额和联动，不可因有修订痕迹就认定内容正确。",
        "",
        "| 位置 | 修订前 | 修订后 |",
        "|---|---|---|",
    ])
    for change in result["tracked_changes"][:200]:
        lines.append(
            f"| {(change.get('location') or '—').replace('|', chr(92) + '|')} | "
            f"{(change.get('before') or '—').replace('|', chr(92) + '|').replace(chr(10), ' ')[:600]} | "
            f"{(change.get('after') or '—').replace('|', chr(92) + '|').replace(chr(10), ' ')[:600]} |"
        )
    if not result["tracked_changes"]:
        lines.append("| — | — | — |")
    lines.extend([
        "",
        "## 六、检测方法与局限",
        "",
        "- 以前一版选定视图（默认接受视图）为基线，以后一版拒绝全部修订后的文本作为“未留痕基线”。",
        "- 以后一版拒绝全部修订与接受全部修订的差异识别留痕改动。",
        "- 比较修订/批注签名、隐藏文字、媒体/嵌入对象和跟踪修改开关。",
        "- 文本比较按正文/表格/文本框分组进行；表格行号变化、复杂域、自动编号、页面布局和图片内文字",
        "  仍可能导致误报或漏报，须在 Word/WPS 中复核。",
        "- 未提供前一版、只提供后一版，或两版不是同族时，本报告不适用。",
        "- 结论不替代邮件、审批记录、签署权限和实际交付证据；发现未留痕改动应先询问，不直接指控。",
    ])
    return "\n".join(lines) + "\n"


def write_outputs(result: dict[str, Any], output_dir: Path, *, force: bool) -> dict[str, Any]:
    output_dir = output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    names = ["hidden-change.json", "hidden-change-report.md", "content-diff.csv"]
    existing = [name for name in names if (output_dir / name).exists()]
    if existing and not force:
        raise HiddenChangeError(f"输出文件已存在：{', '.join(existing)}；确认后用 --force 覆盖")
    temporary = Path(tempfile.mkdtemp(prefix=".hidden-change-", dir=output_dir))
    try:
        (temporary / "hidden-change.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        (temporary / "hidden-change-report.md").write_text(render_markdown(result), encoding="utf-8")
        with (temporary / "content-diff.csv").open("w", newline="", encoding="utf-8-sig") as stream:
            writer = csv.writer(stream)
            writer.writerow(["side", "kind", "location", "clause", "before", "after"])
            for side in ("untracked_changes", "tracked_changes", "total_changes"):
                for change in result[side]:
                    writer.writerow([
                        change.get("side", side), change.get("kind"), change.get("location"),
                        change.get("clause"), change.get("before"), change.get("after"),
                    ])
        for name in names:
            os.replace(temporary / name, output_dir / name)
    finally:
        shutil.rmtree(temporary, ignore_errors=True)
    return {
        "dir": str(output_dir),
        "files": names,
        "before_sha256": result["before"]["sha256"],
        "after_sha256": result["after"]["sha256"],
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="两份 DOCX 的未留痕改动检测（只读）")
    sub = parser.add_subparsers(dest="command", required=True)
    compare = sub.add_parser("compare", help="比较前一版和后一版，检测未留痕改动")
    compare.add_argument("--before", required=True, help="前一版/基准 DOCX")
    compare.add_argument("--after", required=True, help="后一版/待核对 DOCX")
    compare.add_argument("--out", required=True, help="任务产出目录")
    compare.add_argument("--baseline-view", choices=["accepted", "original"], default="accepted")
    compare.add_argument("--include-hidden", action="store_true",
                         help="将隐藏文字纳入主阅读视图；仍会单独报告隐藏文字变化")
    compare.add_argument("--similarity-threshold", type=float, default=DEFAULT_SIMILARITY)
    compare.add_argument("--force", action="store_true", help="覆盖本工具已有的输出文件")

    def handler(args: argparse.Namespace) -> int:
        result = compare_documents(
            Path(args.before).expanduser(),
            Path(args.after).expanduser(),
            Path(args.out).expanduser(),
            baseline_view=args.baseline_view,
            include_hidden=args.include_hidden,
            similarity_threshold=args.similarity_threshold,
            force=args.force,
        )
        print(json.dumps({
            "schema": SCHEMA,
            "ok": True,
            "counts": result["counts"],
            "similarity": result["similarity"],
            "severity": result["severity"],
            "outputs": result["outputs"],
        }, ensure_ascii=False, indent=2))
        return 0 if not result["counts"]["untracked_text_changes"] else 1

    compare.set_defaults(handler=handler)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.handler(args)
    except (HiddenChangeError, wr.ReaderError, OSError, ValueError, zipfile.BadZipFile) as exc:
        print(json.dumps({
            "schema": SCHEMA, "ok": False, "error": str(exc),
        }, ensure_ascii=False, indent=2))
        return 2
    except Exception as exc:
        print(json.dumps({
            "schema": SCHEMA, "ok": False,
            "error": f"内部错误（{type(exc).__name__}）：{exc}",
        }, ensure_ascii=False, indent=2))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
