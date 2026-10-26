#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Word 合同精细阅读器：从 OOXML 重建阅读视图、修订时间线和批注树。

只读，不修改输入文件。读取的是 OOXML 中真实存储的修订、批注和文档结构；
“接受全部”/“拒绝全部”是供阅读的分析视图，不自动写回 Word，也不代表
用户在 Word/WPS 中当前选中的显示模式。
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import difflib
import hashlib
import io
import json
import os
import re
import sys
import tempfile
from pathlib import Path
from typing import Any

try:
    import hetong_redline as redline
except ImportError as exc:  # pragma: no cover - 入口被人为移动时的可读错误
    raise SystemExit(f"无法导入 hetong_redline.py：{exc}") from exc

SCHEMA = "contact-word-read/1"
REV_TEXT_TAGS = {"ins", "del", "moveFrom", "moveTo"}
FORMAT_TAGS = {
    "rPrChange",
    "pPrChange",
    "tblPrChange",
    "trPrChange",
    "tcPrChange",
    "tblGridChange",
    "sectPrChange",
    "numberingChange",
}
STRUCTURAL_TAGS = {"cellIns", "cellDel", "cellMerge", "trIns", "trDel"}
HIDDEN_TAGS = {"vanish", "specVanish", "webHidden"}
OFF_VALUES = {"0", "false", "off", "none"}


class ReaderError(RuntimeError):
    pass


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def local_attr(node: Any, name: str) -> str | None:
    return redline.attr_local(node, name)


def local(node: Any) -> str:
    return redline.local_name(node) or ""


def child(node: Any, name: str) -> Any | None:
    for item in redline.element_children(node):
        if local(item) == name:
            return item
    return None


def children(node: Any, name: str) -> list[Any]:
    return [item for item in redline.element_children(node) if local(item) == name]


def toggle(rpr: Any | None, name: str) -> bool | None:
    if rpr is None:
        return None
    for item in redline.element_children(rpr):
        if local(item) != name:
            continue
        value = (local_attr(item, "val") or "true").lower()
        return value not in OFF_VALUES
    return None


def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    return text.strip()


def excerpt(text: str, limit: int = 500) -> str:
    text = normalize_text(text)
    if len(text) <= limit:
        return text
    return text[:limit] + "…"


def resolve_tag_text(tag: str) -> str:
    return {
        "ins": "插入", "del": "删除", "moveFrom": "移出", "moveTo": "移入",
        "rPrChange": "字符格式变更", "pPrChange": "段落格式变更",
        "tblPrChange": "表格属性变更", "trPrChange": "行属性变更",
        "tcPrChange": "单元格属性变更", "tblGridChange": "表格网格变更",
        "sectPrChange": "节属性变更", "numberingChange": "编号变更",
        "cellIns": "单元格插入", "cellDel": "单元格删除", "cellMerge": "单元格合并",
        "trIns": "行插入", "trDel": "行删除",
    }.get(tag, tag)


def parse_datetime(raw: str | None) -> tuple[dt.datetime | None, str | None, bool]:
    """返回 (datetime, ISO字符串, 是否带时区)。不把无时区时间强行改成北京时间。"""
    if not raw or not str(raw).strip():
        return None, None, False
    value = str(raw).strip()
    try:
        parsed = dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None, None, False
    aware = parsed.tzinfo is not None
    if aware:
        normalized = parsed.astimezone(dt.timezone.utc).isoformat()
    else:
        normalized = parsed.isoformat()
    return parsed, normalized, aware


def sorting_ts(parsed: dt.datetime | None) -> float | None:
    if parsed is None:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=dt.timezone.utc)
    return parsed.timestamp()


def format_summary(rpr: Any | None) -> dict[str, Any]:
    if rpr is None:
        return {}
    result: dict[str, Any] = {}
    fonts = child(rpr, "rFonts")
    if fonts is not None:
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            value = local_attr(fonts, attr)
            if value:
                result[attr] = value
    sz = child(rpr, "sz")
    if sz is not None:
        raw = local_attr(sz, "val")
        if raw and raw.isdigit():
            result["size_pt"] = int(raw) / 2
    color = child(rpr, "color")
    if color is not None:
        for attr in ("val", "themeColor"):
            value = local_attr(color, attr)
            if value:
                result[attr] = value
    for name in ("b", "i", "u", "strike", "caps", "vanish"):
        value = toggle(rpr, name)
        if value is not None:
            result[name] = value
    return result


def paragraph_summary(ppr: Any | None) -> dict[str, Any]:
    if ppr is None:
        return {}
    result: dict[str, Any] = {}
    style = child(ppr, "pStyle")
    if style is not None:
        result["style"] = local_attr(style, "val")
    outline = child(ppr, "outlineLvl")
    if outline is not None:
        raw = local_attr(outline, "val")
        if raw and raw.isdigit():
            result["outline_level"] = int(raw)
    jc = child(ppr, "jc")
    if jc is not None:
        result["alignment"] = local_attr(jc, "val")
    num = child(ppr, "numPr")
    if num is not None:
        ilvl = child(num, "ilvl")
        num_id = child(num, "numId")
        result["numbering"] = {
            "level": local_attr(ilvl, "val") if ilvl is not None else None,
            "id": local_attr(num_id, "val") if num_id is not None else None,
        }
    spacing = child(ppr, "spacing")
    if spacing is not None:
        before = local_attr(spacing, "before")
        after = local_attr(spacing, "after")
        if before or after:
            result["spacing"] = {"before": before, "after": after}
    return result


class CommentTracker:
    def __init__(self) -> None:
        self.active: set[str] = set()
        self.anchors: dict[str, dict[str, Any]] = {}
        self.anomalies: list[str] = []

    def _anchor(self, cid: str) -> dict[str, Any]:
        return self.anchors.setdefault(cid, {
            "raw_id": cid,
            "parts": [],
            "paragraph_ids": [],
            "accepted_segments": [],
            "original_segments": [],
            "hidden_segments": [],
            "point_paragraph_ids": [],
        })

    def start(self, cid: str, part: str, para_id: str) -> None:
        if cid in self.active:
            self.anomalies.append(f"批注 {cid} 在 {part}/{para_id} 重复开始")
        self.active.add(cid)
        anchor = self._anchor(cid)
        for key, value in (("parts", part), ("paragraph_ids", para_id)):
            if value not in anchor[key]:
                anchor[key].append(value)

    def end(self, cid: str, part: str, para_id: str) -> None:
        if cid not in self.active:
            self.anomalies.append(f"批注 {cid} 在 {part}/{para_id} 结束但未开始")
        self.active.discard(cid)
        anchor = self._anchor(cid)
        for key, value in (("parts", part), ("paragraph_ids", para_id)):
            if value not in anchor[key]:
                anchor[key].append(value)

    def reference(self, cid: str, part: str, para_id: str) -> None:
        anchor = self._anchor(cid)
        for key, value in (("parts", part), ("point_paragraph_ids", para_id)):
            if value not in anchor[key]:
                anchor[key].append(value)

    def record(self, accepted: str, original: str, active: set[str],
               part: str, para_id: str, hidden: bool) -> None:
        for cid in active:
            anchor = self._anchor(cid)
            for key, value in (("parts", part), ("paragraph_ids", para_id)):
                if value not in anchor[key]:
                    anchor[key].append(value)
            if accepted:
                anchor["accepted_segments"].append(accepted)
            if original:
                anchor["original_segments"].append(original)
            if hidden:
                anchor["hidden_segments"].append(accepted or original)

    def close_story(self, part: str) -> None:
        for cid in sorted(self.active):
            self.anomalies.append(f"批注 {cid} 在 {part} 结束时仍未闭合")
        self.active.clear()


class WordReader:
    def __init__(
        self,
        path: Path,
        *,
        include_hidden: bool = False,
        session_gap_minutes: int = 30,
        max_report_events: int = 500,
    ):
        self.path = path
        self.include_hidden = include_hidden
        self.session_gap_minutes = session_gap_minutes
        self.max_report_events = max_report_events
        self.pkg: Any = None
        self.doc: Any = None
        self.security: dict[str, Any] = {}
        self.settings: dict[str, Any] = {}
        self.core: dict[str, Any] = {}
        self.app: dict[str, Any] = {}
        self.styles: dict[str, dict[str, Any]] = {}
        self.style_by_name: dict[str, str] = {}
        self.default_hidden: bool | None = None
        self.comments: dict[str, dict[str, Any]] = {}
        self.comment_meta: dict[str, dict[str, Any]] = {}
        self.comments_extended: dict[str, dict[str, Any]] = {}
        self.comments_ids: dict[str, str] = {}
        self.comments_extensible: dict[str, str] = {}
        self.people: list[str] = []
        self.modern_comment_parts: list[str] = []
        self.paragraphs: list[dict[str, Any]] = []
        self.para_by_id: dict[str, dict[str, Any]] = {}
        self.para_by_elem: dict[int, dict[str, Any]] = {}
        self.part_maps: dict[str, dict[str, Any]] = {}
        self.story_parts: list[str] = []
        self.tracker = CommentTracker()
        self.anchors: dict[str, dict[str, Any]] = {}
        self.revisions: list[dict[str, Any]] = []
        self.events: list[dict[str, Any]] = []
        self.sessions: list[dict[str, Any]] = []
        self.clauses: list[dict[str, Any]] = []
        self.views: dict[str, list[dict[str, Any]]] = {}
        self.anomalies: list[dict[str, Any]] = []
        self.objects: dict[str, int] = {}
        self.moves: list[dict[str, Any]] = []
        self.original_sha256 = ""
        self.file_size = 0
        self.file_mtime = ""
        self.new_revision_ids: list[str] = []

    def anomaly(self, severity: str, kind: str, message: str, evidence: Any = None) -> None:
        self.anomalies.append({
            "severity": severity,
            "kind": kind,
            "message": message,
            "evidence": evidence,
        })

    def load(self) -> None:
        if self.path.suffix.lower() != ".docx":
            raise ReaderError("仅接受 .docx；旧版 .doc、加密文档或 .wps 请先另存并核对转换结果")
        if not self.path.is_file():
            raise ReaderError(f"文件不存在：{self.path}")
        self.file_size = self.path.stat().st_size
        self.file_mtime = dt.datetime.fromtimestamp(
            self.path.stat().st_mtime, tz=dt.timezone.utc
        ).isoformat()
        self.original_sha256 = sha256_file(self.path)
        self.pkg = redline.load_package(self.path)
        self.doc = redline.parse_xml(self.pkg.get("word/document.xml"), "word/document.xml")
        redline.resolve_w_prefix(self.doc.documentElement)
        self.security = redline.analyze_package(self.pkg, self.doc)
        self.story_parts = [
            part for part in redline.story_parts(self.pkg)
            if part != "word/comments.xml"
        ]
        self.story_parts.sort(key=self._story_order)
        self._parse_styles()
        self._parse_core()
        self._parse_settings()
        self._parse_comment_definitions()
        self._parse_stories()
        self._parse_revisions()
        self._assign_clauses()
        self._finalize_comments()
        self.views = self._build_views()
        self._build_clauses()
        self._build_events()
        self._build_anomalies()

    def _story_order(self, part: str) -> tuple[int, str]:
        if part == "word/document.xml":
            return 0, part
        if part.startswith("word/header"):
            return 1, part
        if part.startswith("word/footer"):
            return 2, part
        if part == "word/footnotes.xml":
            return 3, part
        if part == "word/endnotes.xml":
            return 4, part
        return 5, part

    def _story_label(self, part: str) -> str:
        if part == "word/document.xml":
            return "正文"
        match = re.match(r"word/header(\d*)\.xml", part)
        if match:
            return f"页眉{match.group(1) or ''}"
        match = re.match(r"word/footer(\d*)\.xml", part)
        if match:
            return f"页脚{match.group(1) or ''}"
        if part == "word/footnotes.xml":
            return "脚注"
        if part == "word/endnotes.xml":
            return "尾注"
        return part

    def _parse_styles(self) -> None:
        if not self.pkg.has("word/styles.xml"):
            return
        root = redline.parse_xml(self.pkg.get("word/styles.xml"), "word/styles.xml").documentElement

        def find_default_rpr() -> Any | None:
            for item in redline.iter_elements_in_order(root):
                if local(item) == "docDefaults":
                    rpr_default = child(item, "rPrDefault")
                    if rpr_default is not None:
                        return child(rpr_default, "rPr")
            return None

        default_rpr = find_default_rpr()
        for name in HIDDEN_TAGS:
            value = toggle(default_rpr, name)
            if value is not None:
                self.default_hidden = value
                break

        raw_styles: dict[str, dict[str, Any]] = {}
        for item in redline.iter_elements_in_order(root):
            if local(item) != "style":
                continue
            style_id = local_attr(item, "styleId")
            style_type = local_attr(item, "type")
            if not style_id:
                continue
            name_element = child(item, "name")
            name = local_attr(name_element, "val") if name_element is not None else None
            based_element = child(item, "basedOn")
            based = local_attr(based_element, "val") if based_element is not None else None
            rpr = child(item, "rPr")
            hidden = None
            for tag in HIDDEN_TAGS:
                value = toggle(rpr, tag)
                if value is not None:
                    hidden = value
                    break
            raw_styles[style_id] = {
                "id": style_id, "name": name or style_id, "type": style_type,
                "based_on": based, "hidden": hidden,
            }
            if name:
                self.style_by_name[name.lower()] = style_id
        self.styles = raw_styles

    def style_hidden(self, style_id: str | None, seen: set[str] | None = None) -> bool | None:
        if not style_id:
            return None
        seen = set() if seen is None else seen
        if style_id in seen:
            return None
        seen.add(style_id)
        style = self.styles.get(style_id)
        if not style:
            return None
        parent_value = self.style_hidden(style.get("based_on"), seen)
        own = style.get("hidden")
        return own if own is not None else parent_value

    def style_name(self, style_id: str | None) -> str | None:
        style = self.styles.get(style_id or "")
        return style.get("name") if style else style_id

    def _parse_core(self) -> None:
        if self.pkg.has("docProps/core.xml"):
            root = redline.parse_xml(self.pkg.get("docProps/core.xml"), "docProps/core.xml").documentElement
            for item in redline.iter_elements_in_order(root):
                name = local(item)
                value = (redline.text_of(item) or "").strip()
                if not value:
                    continue
                if name == "title" and "title" not in self.core:
                    self.core["title"] = value
                elif name == "subject" and "subject" not in self.core:
                    self.core["subject"] = value
                elif name == "creator" and "creator" not in self.core:
                    self.core["creator"] = value
                elif name == "lastModifiedBy" and "last_modified_by" not in self.core:
                    self.core["last_modified_by"] = value
                elif name == "created" and "created" not in self.core:
                    self.core["created"] = value
                elif name == "modified" and "modified" not in self.core:
                    self.core["modified"] = value
                elif name == "revision" and "revision" not in self.core:
                    self.core["revision"] = value
        if self.pkg.has("docProps/app.xml"):
            root = redline.parse_xml(self.pkg.get("docProps/app.xml"), "docProps/app.xml").documentElement
            for item in redline.iter_elements_in_order(root):
                name = local(item)
                value = (redline.text_of(item) or "").strip()
                if value and name in {"Pages", "Words", "Characters", "Paragraphs", "Company"}:
                    self.app[name] = value

    def _parse_settings(self) -> None:
        self.settings = {
            "track_changes": None,
            "document_protection": self.security.get("document_protection"),
            "write_protection": self.security.get("write_protection"),
        }
        if not self.pkg.has("word/settings.xml"):
            return
        root = redline.parse_xml(self.pkg.get("word/settings.xml"), "word/settings.xml").documentElement
        for item in redline.iter_elements_in_order(root):
            name = local(item)
            if name == "trackChanges":
                value = (local_attr(item, "val") or "true").lower()
                self.settings["track_changes"] = value not in OFF_VALUES
            elif name == "revisionView":
                self.settings["revision_view"] = {
                    re.sub(r"Markup$", "", local_attr(item, key) or key): local_attr(item, key)
                    for key in ("insMarkup", "delMarkup", "formatMarkup", "commentMarkup")
                    if local_attr(item, key) is not None
                }
            elif name == "doNotTrackMoves":
                self.settings["do_not_track_moves"] = True
            elif name == "doNotTrackFormatting":
                self.settings["do_not_track_formatting"] = True

    def _plain_comment_text(self, comment_element: Any, mode: str) -> str:
        lines: list[str] = []

        def walk(node: Any, rev_tags: list[str]) -> None:
            if node.nodeType == node.TEXT_NODE:
                return
            name = local(node)
            if name in {"pPr", "rPr"}:
                return
            if name in {"t", "delText"}:
                text = redline.text_of(node)
                inserted = any(tag in {"ins", "moveTo"} for tag in rev_tags)
                deleted = any(tag in {"del", "moveFrom"} for tag in rev_tags)
                if mode == "accepted" and deleted:
                    return
                if mode == "original" and inserted:
                    return
                if text:
                    lines.append(text)
                return
            if name in {"tab"}:
                lines.append("\t")
                return
            if name in {"br", "cr"}:
                lines.append("\n")
                return
            if name == "p" and lines and lines[-1] != "\n":
                lines.append("\n")
            next_stack = [*rev_tags, name] if name in REV_TEXT_TAGS else rev_tags
            for ch in redline.element_children(node):
                walk(ch, next_stack)
            if name == "p" and lines and lines[-1] != "\n":
                lines.append("\n")

        walk(comment_element, [])
        return normalize_text("".join(lines))

    def _parse_comment_definitions(self) -> None:
        if not self.pkg.has("word/comments.xml"):
            return
        root = redline.parse_xml(self.pkg.get("word/comments.xml"), "word/comments.xml").documentElement
        definition_elements = [
            item for item in redline.iter_elements_in_order(root) if local(item) == "comment"
        ]
        for item in definition_elements:
            raw_id = local_attr(item, "id")
            if raw_id is None:
                continue
            paragraphs = [
                sub for sub in redline.iter_elements_in_order(item) if local(sub) == "p"
            ]
            para_ids = [local_attr(p, "paraId") for p in paragraphs if local_attr(p, "paraId")]
            self.comments[str(raw_id)] = {
                "raw_id": str(raw_id),
                "author": local_attr(item, "author"),
                "initials": local_attr(item, "initials"),
                "date_raw": local_attr(item, "date"),
                "accepted_text": self._plain_comment_text(item, "accepted"),
                "original_text": self._plain_comment_text(item, "original"),
                "paragraph_ids": para_ids,
            }
        modern = ["word/commentsExtended.xml", "word/commentsIds.xml",
                  "word/commentsExtensible.xml", "word/people.xml"]
        self.modern_comment_parts = [part for part in modern if self.pkg.has(part)]
        if self.pkg.has("word/commentsExtended.xml"):
            root = redline.parse_xml(
                self.pkg.get("word/commentsExtended.xml"), "word/commentsExtended.xml"
            ).documentElement
            for item in redline.iter_elements_in_order(root):
                if local(item) != "commentEx":
                    continue
                para_id = local_attr(item, "paraId")
                if para_id:
                    done_raw = local_attr(item, "done")
                    self.comments_extended[para_id] = {
                        "parent_para_id": local_attr(item, "paraIdParent"),
                        "done": None if done_raw is None else done_raw.lower() in {"1", "true", "on"},
                    }
        if self.pkg.has("word/commentsIds.xml"):
            root = redline.parse_xml(self.pkg.get("word/commentsIds.xml"), "word/commentsIds.xml").documentElement
            for item in redline.iter_elements_in_order(root):
                if local(item) == "commentId":
                    para_id = local_attr(item, "paraId")
                    durable = local_attr(item, "durableId")
                    if para_id and durable:
                        self.comments_ids[para_id] = durable
        if self.pkg.has("word/commentsExtensible.xml"):
            root = redline.parse_xml(
                self.pkg.get("word/commentsExtensible.xml"), "word/commentsExtensible.xml"
            ).documentElement
            for item in redline.iter_elements_in_order(root):
                if local(item) == "commentExtensible":
                    durable = local_attr(item, "durableId")
                    date = local_attr(item, "dateUtc")
                    if durable and date:
                        self.comments_extensible[durable] = date
        if self.pkg.has("word/people.xml"):
            root = redline.parse_xml(self.pkg.get("word/people.xml"), "word/people.xml").documentElement
            for item in redline.iter_elements_in_order(root):
                if local(item) == "person":
                    author = local_attr(item, "author")
                    if author and author not in self.people:
                        self.people.append(author)

    def _parent_map(self, doc: Any) -> dict[int, Any]:
        result: dict[int, Any] = {}
        root = doc.documentElement if hasattr(doc, "documentElement") else doc
        stack = [root]
        while stack:
            node = stack.pop()
            for ch in redline.element_children(node):
                result[id(ch)] = node
                stack.append(ch)
        return result

    def _table_ids(self, root: Any) -> dict[int, str]:
        result: dict[int, str] = {}
        for item in redline.iter_elements_in_order(root):
            if local(item) == "tbl":
                result[id(item)] = f"T{len(result) + 1}"
        return result

    def _ancestors(self, node: Any, parent_map: dict[int, Any]) -> list[Any]:
        result = []
        current = parent_map.get(id(node))
        while current is not None:
            result.append(current)
            current = parent_map.get(id(current))
        return result

    def _element_index(self, element: Any, parent_map: dict[int, Any], name: str) -> int | None:
        parent = parent_map.get(id(element))
        if parent is None:
            return None
        peers = [item for item in redline.element_children(parent) if local(item) == name]
        for index, item in enumerate(peers, 1):
            if item is element:
                return index
        return None

    def _revision_state(self, element: Any | None) -> dict[str, Any] | None:
        if element is None:
            return None
        tag = local(element)
        return {
            "tag": tag,
            "id": local_attr(element, "id"),
            "author": local_attr(element, "author"),
            "date": local_attr(element, "date"),
            "name": local_attr(element, "name"),
        }

    def _container_info(self, para: Any, part: str, parent_map: dict[int, Any],
                        table_ids: dict[int, str], textbox_ids: dict[int, str]) -> dict[str, Any]:
        chain = self._ancestors(para, parent_map)
        tables = [item for item in chain if local(item) == "tbl"]
        rows = [item for item in chain if local(item) == "tr"]
        cells = [item for item in chain if local(item) == "tc"]
        textboxes = [item for item in chain if local(item) == "txbxContent"]
        table = tables[0] if tables else None
        row = rows[0] if rows else None
        cell = cells[0] if cells else None
        textbox = textboxes[0] if textboxes else None
        row_states = []
        for item in rows:
            trpr = child(item, "trPr")
            if trpr is not None:
                for rev in redline.element_children(trpr):
                    if local(rev) in {"ins", "del", "trIns", "trDel"}:
                        state = self._revision_state(rev)
                        if state:
                            row_states.append(state)
        cell_states = []
        for item in cells:
            tcpr = child(item, "tcPr")
            if tcpr is not None:
                for rev in redline.element_children(tcpr):
                    if local(rev) in {"cellIns", "cellDel", "cellMerge"}:
                        state = self._revision_state(rev)
                        if state:
                            cell_states.append(state)
        accepted_excluded = any(item["tag"] in {"del", "trDel", "cellDel"} for item in row_states + cell_states)
        original_excluded = any(item["tag"] in {"ins", "trIns", "cellIns"} for item in row_states + cell_states)
        bits = [self._story_label(part)]
        flow_bits = [part]
        if table is not None:
            table_id = table_ids.get(id(table), "T?")
            row_index = self._element_index(row, parent_map, "tr") if row is not None else None
            cell_index = self._element_index(cell, parent_map, "tc") if cell is not None else None
            bits.append(f"{table_id} R{row_index or '?'}C{cell_index or '?'}")
            flow_bits.extend(["table", table_id, str(row_index), str(cell_index)])
        if textbox is not None:
            textbox_id = textbox_ids.get(id(textbox), "TB?")
            bits.append(textbox_id)
            flow_bits.extend(["textbox", textbox_id])
        if len(bits) == 1:
            flow_bits.append("body")
        return {
            "story": part,
            "story_label": self._story_label(part),
            "location": " / ".join(bits),
            "flow": "|".join(flow_bits),
            "row_states": row_states,
            "cell_states": cell_states,
            "accepted_excluded": accepted_excluded,
            "original_excluded": original_excluded,
            "table_id": table_ids.get(id(table)) if table is not None else None,
            "row_index": self._element_index(row, parent_map, "tr") if row is not None else None,
            "cell_index": self._element_index(cell, parent_map, "tc") if cell is not None else None,
            "textbox_id": textbox_ids.get(id(textbox)) if textbox is not None else None,
        }

    def _paragraph_mark(self, ppr: Any | None) -> dict[str, Any]:
        result = {"inserted": False, "deleted": False, "ids": [], "author": None, "date": None}
        if ppr is None:
            return result
        rpr = child(ppr, "rPr")
        if rpr is None:
            return result
        for item in redline.element_children(rpr):
            if local(item) == "ins":
                result["inserted"] = True
                result["ids"].append(local_attr(item, "id"))
                result["author"] = local_attr(item, "author")
                result["date"] = local_attr(item, "date")
            elif local(item) == "del":
                result["deleted"] = True
                result["ids"].append(local_attr(item, "id"))
                result["author"] = local_attr(item, "author")
                result["date"] = local_attr(item, "date")
        return result

    def _run_hidden(self, run: Any, para_style: str | None) -> tuple[bool, str]:
        rpr = child(run, "rPr")
        style_id = None
        if rpr is not None:
            rstyle = child(rpr, "rStyle")
            style_id = local_attr(rstyle, "val") if rstyle is not None else None
        hidden = bool(self.default_hidden)
        confidence = "default"
        para_hidden = self.style_hidden(para_style)
        if para_hidden is not None:
            hidden = para_hidden
            confidence = "paragraph_style"
        run_hidden = self.style_hidden(style_id)
        if run_hidden is not None:
            hidden = run_hidden
            confidence = "character_style"
        for tag in HIDDEN_TAGS:
            direct = toggle(rpr, tag)
            if direct is not None:
                hidden = direct
                confidence = "direct"
                break
        return hidden, confidence

    def _parse_paragraph(self, para: Any, part: str, para_id: str,
                         parent_map: dict[int, Any], table_ids: dict[int, str],
                         textbox_ids: dict[int, str]) -> dict[str, Any]:
        ppr = child(para, "pPr")
        style_element = child(ppr, "pStyle") if ppr is not None else None
        style_id = local_attr(style_element, "val") if style_element is not None else None
        outline_element = child(ppr, "outlineLvl") if ppr is not None else None
        outline_raw = local_attr(outline_element, "val") if outline_element is not None else None
        outline_level = int(outline_raw) if outline_raw and outline_raw.isdigit() else None
        container = self._container_info(para, part, parent_map, table_ids, textbox_ids)
        mark = self._paragraph_mark(ppr)
        segments: list[dict[str, Any]] = []
        hidden_segments: list[str] = []
        point_comments: set[str] = set()

        def add_text(text: str, rev_stack: list[dict[str, Any]], hidden: bool) -> None:
            if not text:
                return
            tags = [item["tag"] for item in rev_stack]
            inserted = any(tag in {"ins", "moveTo"} for tag in tags)
            deleted = any(tag in {"del", "moveFrom"} for tag in tags)
            conflict = inserted and deleted
            # 同时处于插入和删除中的嵌套结构无法可靠解释，主视图均不收录。
            accepted = False if conflict else not deleted
            original = False if conflict else not inserted
            active = set(self.tracker.active)
            self.tracker.record(
                text if accepted else "",
                text if original else "",
                active, part, para_id, hidden,
            )
            if hidden:
                hidden_segments.append(text)
            markup = text
            if hidden:
                markup = f"⟦隐藏:{markup}⟧"
            if conflict:
                markup = f"⟦嵌套冲突修订:{markup}⟧"
            elif rev_stack:
                innermost = rev_stack[-1]
                label = "插入" if innermost["tag"] in {"ins", "moveTo"} else "删除"
                if innermost["tag"] in {"moveFrom", "moveTo"}:
                    label = "移出" if innermost["tag"] == "moveFrom" else "移入"
                markup = (
                    f"⟦{label}#{innermost.get('id') or '?'}"
                    f"|{innermost.get('author') or '未知'}"
                    f"|{innermost.get('date') or '无日期'}:{markup}⟧"
                )
            segments.append({
                "text": text, "accepted": accepted, "original": original,
                "hidden": hidden, "comments": sorted(active),
                "rev_tags": tags, "markup": markup, "conflict": conflict,
            })

        def process_item(item: Any, rev_stack: list[dict[str, Any]], hidden: bool) -> None:
            name = local(item)
            if name == "commentRangeStart":
                cid = local_attr(item, "id")
                if cid is not None:
                    self.tracker.start(cid, part, para_id)
                return
            if name == "commentRangeEnd":
                cid = local_attr(item, "id")
                if cid is not None:
                    self.tracker.end(cid, part, para_id)
                return
            if name == "commentReference":
                cid = local_attr(item, "id")
                if cid is not None:
                    point_comments.add(cid)
                    self.tracker.reference(cid, part, para_id)
                return
            if name in {"pPr", "rPr", "instrText"}:
                return
            if name == "r":
                run_hidden, _ = self._run_hidden(item, style_id)
                for ch in redline.element_children(item):
                    process_item(ch, rev_stack, run_hidden)
                return
            if name in REV_TEXT_TAGS:
                state = self._revision_state(item)
                if state:
                    for ch in redline.element_children(item):
                        process_item(ch, [*rev_stack, state], hidden)
                else:
                    for ch in redline.element_children(item):
                        process_item(ch, rev_stack, hidden)
                return
            if name in {"t", "delText"}:
                add_text(redline.text_of(item), rev_stack, hidden)
                return
            if name in {"br", "cr"}:
                add_text("\n", rev_stack, hidden)
                return
            if name == "tab":
                add_text("\t", rev_stack, hidden)
                return
            if name == "noBreakHyphen":
                add_text("-", rev_stack, hidden)
                return
            if name == "sym":
                add_text("�", rev_stack, hidden)
                return
            for ch in redline.element_children(item):
                process_item(ch, rev_stack, hidden)

        for node in redline.element_children(para):
            process_item(node, [], False)
        accepted_text = normalize_text("".join(
            item["text"] for item in segments
            if item["accepted"] and (self.include_hidden or not item["hidden"])
        ))
        original_text = normalize_text("".join(
            item["text"] for item in segments
            if item["original"] and (self.include_hidden or not item["hidden"])
        ))
        markup_text = normalize_text("".join(item["markup"] for item in segments))
        comment_ids = set(point_comments)
        for item in segments:
            comment_ids.update(item["comments"])
        if not accepted_text and not original_text and not markup_text:
            markup_text = ""
        hidden_text = normalize_text("".join(hidden_segments))
        paragraph = {
            "id": para_id,
            "story": part,
            "story_label": container["story_label"],
            "location": container["location"],
            "flow": container["flow"],
            "table_id": container["table_id"],
            "row_index": container["row_index"],
            "cell_index": container["cell_index"],
            "textbox_id": container["textbox_id"],
            "style_id": style_id,
            "style_name": self.style_name(style_id),
            "outline_level": outline_level,
            "clause_id": None,
            "clause_path": None,
            "heading": None,
            "accepted_text": accepted_text,
            "original_text": original_text,
            "markup_text": markup_text,
            "hidden_text": hidden_text,
            "comment_ids": sorted(comment_ids),
            "paragraph_mark": mark,
            "row_states": container["row_states"],
            "cell_states": container["cell_states"],
            "accepted_excluded": container["accepted_excluded"],
            "original_excluded": container["original_excluded"],
        }
        if container["row_states"] or container["cell_states"]:
            self.objects["structural_table_markers"] = self.objects.get("structural_table_markers", 0) + 1
        return paragraph

    def _parse_stories(self) -> None:
        self.tracker = CommentTracker()
        textbox_counter = 0
        for part in self.story_parts:
            root = redline.parse_xml(self.pkg.get(part), part).documentElement
            parent_map = self._parent_map(root)
            table_ids = self._table_ids(root)
            textbox_ids: dict[int, str] = {}
            for item in redline.iter_elements_in_order(root):
                if local(item) == "txbxContent":
                    textbox_counter += 1
                    key = f"TB{textbox_counter}"
                    textbox_ids[id(item)] = key
            # 对象和结构覆盖统计
            for item in redline.iter_elements_in_order(root):
                name = local(item)
                if name in {"drawing", "pict", "object", "altChunk", "sdt", "fldSimple", "instrText"}:
                    key = {
                        "drawing": "drawings", "pict": "pictures", "object": "embedded_objects",
                        "altChunk": "alt_chunks", "sdt": "content_controls",
                        "fldSimple": "fields", "instrText": "field_instructions",
                    }[name]
                    self.objects[key] = self.objects.get(key, 0) + 1
            paragraphs = [
                item for item in redline.iter_elements_in_order(root) if local(item) == "p"
            ]
            for paragraph_element in paragraphs:
                para_id = f"WR{len(self.paragraphs) + 1:04d}"
                paragraph = self._parse_paragraph(
                    paragraph_element, part, para_id, parent_map, table_ids, textbox_ids
                )
                self.paragraphs.append(paragraph)
                self.para_by_id[para_id] = paragraph
                self.para_by_elem[id(paragraph_element)] = paragraph
            self.tracker.close_story(part)
            self.part_maps[part] = {
                "root": root,
                "parent_map": parent_map,
                "table_ids": table_ids,
                "textbox_ids": textbox_ids,
            }
        for message in self.tracker.anomalies:
            self.anomaly("medium", "comment_anchor", message)

    def _nearest_paragraph(self, element: Any, parent_map: dict[int, Any]) -> dict[str, Any] | None:
        current = element
        while current is not None:
            paragraph = self.para_by_elem.get(id(current))
            if paragraph is not None:
                return paragraph
            current = parent_map.get(id(current))
        return None

    def _element_location(self, element: Any, part: str, parent_map: dict[int, Any]) -> tuple[str, dict[str, Any] | None]:
        paragraph = self._nearest_paragraph(element, parent_map)
        if paragraph is not None:
            return paragraph["location"], paragraph
        info = self.part_maps.get(part, {})
        table_ids = info.get("table_ids", {})
        chain = self._ancestors(element, parent_map)
        table = next((item for item in chain if local(item) == "tbl"), None)
        row = next((item for item in chain if local(item) == "tr"), None)
        if table is not None:
            table_id = table_ids.get(id(table), "T?")
            row_index = self._element_index(row, parent_map, "tr") if row is not None else None
            return f"{self._story_label(part)} / {table_id} R{row_index or '?'}", None
        return self._story_label(part), None

    def _revision_text(self, element: Any) -> str:
        texts = []
        for item in redline.iter_elements_in_order(element):
            name = local(item)
            if name in {"t", "delText"}:
                texts.append(redline.text_of(item))
            elif name in {"br", "cr"}:
                texts.append("\n")
            elif name == "tab":
                texts.append("\t")
        return normalize_text("".join(texts))

    def _classify_revision(self, element: Any, parent_map: dict[int, Any]) -> str:
        tag = local(element)
        parent = parent_map.get(id(element))
        parent_name = local(parent) if parent is not None else ""
        if tag in {"ins", "del"} and parent_name == "rPr":
            grandparent = parent_map.get(id(parent)) if parent is not None else None
            if grandparent is not None and local(grandparent) == "pPr":
                return "paragraphMarkInserted" if tag == "ins" else "paragraphMarkDeleted"
        if tag in {"ins", "del"} and parent_name == "trPr":
            return "tableRowInserted" if tag == "ins" else "tableRowDeleted"
        if tag in {"ins", "del"} and parent_name == "tcPr":
            return "tableCellInserted" if tag == "ins" else "tableCellDeleted"
        if tag == "ins":
            return "insertedText"
        if tag == "del":
            return "deletedText"
        if tag == "moveFrom":
            return "moveFrom"
        if tag == "moveTo":
            return "moveTo"
        return {
            "rPrChange": "runFormatting",
            "pPrChange": "paragraphFormatting",
            "tblPrChange": "tableFormatting",
            "trPrChange": "rowFormatting",
            "tcPrChange": "cellFormatting",
            "tblGridChange": "tableGridFormatting",
            "sectPrChange": "sectionFormatting",
            "numberingChange": "numberingFormatting",
            "cellIns": "tableCellInserted",
            "cellDel": "tableCellDeleted",
            "cellMerge": "tableCellMerged",
            "trIns": "tableRowInserted",
            "trDel": "tableRowDeleted",
        }.get(tag, tag)

    def _format_change(self, element: Any, parent_map: dict[int, Any]) -> dict[str, Any]:
        tag = local(element)
        old = child(element, "rPr") if tag in {"rPrChange", "pPrChange"} else None
        if tag == "rPrChange":
            parent = parent_map.get(id(element))
            current = parent if parent is not None and local(parent) == "rPr" else None
            return {"old": format_summary(old), "new": format_summary(current)}
        if tag == "pPrChange":
            parent = parent_map.get(id(element))
            current = parent if parent is not None and local(parent) == "pPr" else None
            return {"old": paragraph_summary(old), "new": paragraph_summary(current)}
        return {"old": {}, "new": {}}

    def _parse_revisions(self) -> None:
        duplicate_ids: dict[tuple[str, str], int] = {}
        source_order = 0
        for part in self.story_parts:
            info = self.part_maps[part]
            parent_map = info["parent_map"]
            root = info["root"]
            for element in redline.iter_elements_in_order(root):
                tag = local(element)
                if tag not in REV_TEXT_TAGS and tag not in FORMAT_TAGS and tag not in STRUCTURAL_TAGS:
                    continue
                if tag in {"ins", "del"}:
                    parent = parent_map.get(id(element))
                    if parent is not None and local(parent) in {"rPr", "trPr", "tcPr", "tblPr"}:
                        # 结构/段落标记修订可能在父级检查中再次出现；此处只记录一次。
                        pass
                source_order += 1
                paragraph = self._nearest_paragraph(element, parent_map)
                location, _ = self._element_location(element, part, parent_map)
                revision_id = local_attr(element, "id")
                if revision_id is not None:
                    key = (part, revision_id)
                    duplicate_ids[key] = duplicate_ids.get(key, 0) + 1
                date_raw = local_attr(element, "date")
                parsed, normalized, aware = parse_datetime(date_raw)
                record = {
                    "source_order": source_order,
                    "timeline_id": None,
                    "revision_id": revision_id,
                    "type": self._classify_revision(element, parent_map),
                    "tag": tag,
                    "author": local_attr(element, "author"),
                    "date_raw": date_raw,
                    "date_normalized": normalized,
                    "date_utc": normalized if aware else None,
                    "date_timezone_aware": aware,
                    "part": part,
                    "location": location,
                    "paragraph_id": paragraph["id"] if paragraph else None,
                    "clause_id": None,
                    "clause_path": None,
                    "text": self._revision_text(element),
                    "context_accepted": excerpt(paragraph["accepted_text"] if paragraph else ""),
                    "context_original": excerpt(paragraph["original_text"] if paragraph else ""),
                    "move_name": local_attr(element, "name"),
                    "move_group": None,
                    "move_paired": False,
                    "format_change": self._format_change(element, parent_map) if tag in FORMAT_TAGS else None,
                    "_parsed": parsed,
                    "_sort_ts": sorting_ts(parsed),
                }
                self.revisions.append(record)
        for (part, revision_id), count in duplicate_ids.items():
            if count > 1:
                self.anomaly(
                    "high", "duplicate_revision_id",
                    f"{part} 中修订 id={revision_id} 出现 {count} 次",
                    {"part": part, "revision_id": revision_id, "count": count},
                )
        self._pair_moves()
        self.revisions.sort(key=lambda item: (
            item["_sort_ts"] is None,
            item["_sort_ts"] if item["_sort_ts"] is not None else 0,
            item["source_order"],
        ))
        for index, revision in enumerate(self.revisions, 1):
            revision["timeline_id"] = f"TV{index:04d}"
        self.new_revision_ids = [item["timeline_id"] for item in self.revisions]

    def _pair_moves(self) -> None:
        groups: dict[str, dict[str, list[dict[str, Any]]]] = {}
        for revision in self.revisions:
            if revision["tag"] in {"moveFrom", "moveTo"} and revision.get("move_name"):
                groups.setdefault(revision["move_name"], {"moveFrom": [], "moveTo": []})
                groups[revision["move_name"]][revision["tag"]].append(revision)
        group_index = 0
        for name, group in groups.items():
            from_items = group["moveFrom"]
            to_items = group["moveTo"]
            if len(from_items) == 1 and len(to_items) == 1:
                group_index += 1
                code = f"MV{group_index:04d}"
                from_items[0]["move_group"] = code
                to_items[0]["move_group"] = code
                from_items[0]["move_paired"] = True
                to_items[0]["move_paired"] = True
                self.moves.append({"code": code, "name": name, "from": from_items[0]["source_order"], "to": to_items[0]["source_order"]})
            else:
                self.anomaly(
                    "medium", "move_unpaired",
                    f"移动修订组 {name!r} 无法唯一配对（moveFrom={len(from_items)}, moveTo={len(to_items)}）",
                    {"name": name, "moveFrom": len(from_items), "moveTo": len(to_items)},
                )

    def _heading_info(self, paragraph: dict[str, Any]) -> tuple[int, str] | None:
        text = normalize_text(paragraph["accepted_text"] or paragraph["original_text"])
        if not text:
            return None
        style = (paragraph.get("style_name") or "").lower()
        outline = paragraph.get("outline_level")
        level = None
        match = re.match(r"^(?:heading|标题)\s*([1-9])", style)
        if match:
            level = int(match.group(1))
        elif outline is not None:
            level = min(9, outline + 1)
        chapter = re.match(r"^第[一二三四五六七八九十百千0-9]+[编章节]\s*", text)
        article = re.match(r"^第[一二三四五六七八九十百千0-9]+条\s*", text)
        attachment = re.match(r"^附件[一二三四五六七八九十0-9]*\s*", text)
        if chapter and len(text) < 120:
            level = 1
        elif article and len(text) < 120:
            level = 2
        elif attachment and len(text) < 120:
            level = 1
        if level is None:
            return None
        return level, text

    def _assign_clauses(self) -> None:
        stack: list[tuple[int, str, str]] = []
        current = ("PRE", "前言")
        clause_index = 0
        for paragraph in self.paragraphs:
            if paragraph["story"] != "word/document.xml":
                paragraph["clause_id"] = None
                paragraph["clause_path"] = None
                continue
            heading = self._heading_info(paragraph)
            if heading is not None:
                level, title = heading
                clause_index += 1
                clause_id = f"H{clause_index:04d}"
                while stack and stack[-1][0] >= level:
                    stack.pop()
                stack.append((level, clause_id, title))
                current = (clause_id, " > ".join(item[2] for item in stack))
                paragraph["heading"] = {"level": level, "title": title}
            paragraph["clause_id"] = current[0]
            paragraph["clause_path"] = current[1]
        for revision in self.revisions:
            paragraph = self.para_by_id.get(revision.get("paragraph_id") or "")
            if paragraph:
                revision["clause_id"] = paragraph["clause_id"]
                revision["clause_path"] = paragraph["clause_path"]

    def _finalize_comments(self) -> None:
        for cid, anchor in self.tracker.anchors.items():
            self.anchors[cid] = {
                "parts": anchor["parts"],
                "paragraph_ids": anchor["paragraph_ids"],
                "point_paragraph_ids": anchor["point_paragraph_ids"],
                "accepted_text": normalize_text("".join(anchor["accepted_segments"])),
                "original_text": normalize_text("".join(anchor["original_segments"])),
                "hidden_text": normalize_text("".join(anchor["hidden_segments"])),
            }
        raw_ids = set(self.comments) | set(self.anchors)
        for raw_id in raw_ids:
            definition = self.comments.get(raw_id)
            anchor = self.anchors.get(raw_id, {})
            if definition is None:
                self.comment_meta[raw_id] = {
                    "raw_id": raw_id,
                    "author": None,
                    "initials": None,
                    "date_raw": None,
                    "date_normalized": None,
                    "accepted_text": "",
                    "original_text": "",
                    "paragraph_ids": anchor.get("paragraph_ids", []),
                    "parent_raw": None,
                    "resolved": None,
                    "orphan_anchor": True,
                }
                continue
            last_para = definition["paragraph_ids"][-1] if definition["paragraph_ids"] else None
            extended = self.comments_extended.get(last_para or "") if last_para else None
            if extended is None:
                for para_id in definition["paragraph_ids"]:
                    if para_id in self.comments_extended:
                        extended = self.comments_extended[para_id]
                        break
            durable = self.comments_ids.get(last_para or "") if last_para else None
            modified = self.comments_extensible.get(durable or "") if durable else None
            parsed, normalized, aware = parse_datetime(definition["date_raw"])
            parent_raw = None
            if extended and extended.get("parent_para_id"):
                parent_para = extended["parent_para_id"]
                for candidate_id, candidate in self.comments.items():
                    if parent_para in (candidate.get("paragraph_ids") or []):
                        parent_raw = candidate_id
                        break
            point_ids = anchor.get("point_paragraph_ids", [])
            accepted_anchor = anchor.get("accepted_text", "")
            original_anchor = anchor.get("original_text", "")
            if not accepted_anchor and not original_anchor and point_ids:
                first = self.para_by_id.get(point_ids[0])
                if first:
                    accepted_anchor = first["accepted_text"]
                    original_anchor = first["original_text"]
            self.comment_meta[raw_id] = {
                "raw_id": raw_id,
                "author": definition.get("author"),
                "initials": definition.get("initials"),
                "date_raw": definition.get("date_raw"),
                "date_normalized": normalized,
                "date_utc": normalized if aware else None,
                "modified_date_raw": modified,
                "accepted_text": definition.get("accepted_text", ""),
                "original_text": definition.get("original_text", ""),
                "paragraph_ids": anchor.get("paragraph_ids", []),
                "point_paragraph_ids": point_ids,
                "anchor_accepted_text": excerpt(accepted_anchor, 1000),
                "anchor_original_text": excerpt(original_anchor, 1000),
                "anchor_deleted": bool(original_anchor) and not bool(accepted_anchor),
                "parent_raw": parent_raw,
                "resolved": extended.get("done") if extended else None,
                "orphan_anchor": False,
            }
            self.comment_meta[raw_id]["_parsed"] = parsed
            self.comment_meta[raw_id]["_sort_ts"] = sorting_ts(parsed)
        comments_sorted = sorted(
            self.comment_meta.values(),
            key=lambda item: (
                item.get("_sort_ts") is None,
                item.get("_sort_ts") if item.get("_sort_ts") is not None else 0,
                str(item.get("raw_id")),
            ),
        )
        code_by_raw: dict[str, str] = {}
        for index, comment in enumerate(comments_sorted, 1):
            comment["code"] = f"C{index:04d}"
            code_by_raw[comment["raw_id"]] = comment["code"]
        for comment in comments_sorted:
            parent_raw = comment.get("parent_raw")
            comment["parent_code"] = code_by_raw.get(parent_raw) if parent_raw else None
            if parent_raw and parent_raw not in code_by_raw:
                self.anomaly("medium", "comment_parent_missing",
                              f"批注 {comment['code']} 引用了不存在的父批注 {parent_raw}")
            paragraph_ids = comment.get("paragraph_ids") or []
            clauses: list[str] = []
            for para_id in paragraph_ids:
                paragraph = self.para_by_id.get(para_id)
                if paragraph and paragraph.get("clause_id"):
                    clause = paragraph["clause_id"]
                    if clause not in clauses:
                        clauses.append(clause)
            comment["clause_ids"] = clauses
            comment["clause_id"] = clauses[0] if clauses else None
            comment["clause_path"] = self.para_by_id[paragraph_ids[0]]["clause_path"] if paragraph_ids and paragraph_ids[0] in self.para_by_id else None
            comment["children"] = []
            if comment.get("orphan_anchor"):
                self.anomaly("medium", "orphan_comment_anchor",
                              f"批注 {comment.get('code')} 有批注锚点但没有批注正文",
                              {"raw_id": comment["raw_id"]})
            elif not paragraph_ids and not comment.get("point_paragraph_ids") and not comment.get("parent_raw"):
                comment["orphan"] = True
                self.anomaly("low", "orphan_comment",
                              f"批注 {comment['code']} 有正文但找不到正文锚点",
                              {"raw_id": comment["raw_id"]})
            if comment.get("anchor_deleted"):
                self.anomaly("high", "comment_anchor_deleted",
                              f"批注 {comment.get('code')} 的原文已被删除，接受视图无对应文字",
                              {"raw_id": comment.get("raw_id")})
        self.comments = {item["raw_id"]: item for item in comments_sorted}
        for comment in comments_sorted:
            if comment.get("parent_code"):
                parent = next((item for item in comments_sorted if item["code"] == comment["parent_code"]), None)
                if parent:
                    parent["children"].append(comment["code"])
                    if not comment.get("clause_id"):
                        comment["clause_id"] = parent.get("clause_id")
                        comment["clause_path"] = parent.get("clause_path")
        for comment in comments_sorted:
            thread = [comment]
            seen = {comment["code"]}
            current = comment
            while current.get("parent_code"):
                parent = next((item for item in comments_sorted if item["code"] == current["parent_code"]), None)
                if parent is None or parent["code"] in seen:
                    break
                thread.append(parent)
                seen.add(parent["code"])
                current = parent
            states = [item.get("resolved") for item in thread]
            known_states = [state for state in states if state is not None]
            comment["thread_status"] = (
                "open" if any(not state for state in known_states) else
                "resolved" if states and len(known_states) == len(states) and all(known_states) else "unknown"
            )
        code_by_raw = {item["raw_id"]: item["code"] for item in comments_sorted}
        for paragraph in self.paragraphs:
            paragraph["comment_raw_ids"] = paragraph.get("comment_ids", [])
            paragraph["comment_ids"] = [code_by_raw.get(str(raw), str(raw)) for raw in paragraph.get("comment_ids", [])]

    def _diff_signals(self, original: str, accepted: str) -> list[dict[str, Any]]:
        if original == accepted:
            return []
        snippets = []
        matcher = difflib.SequenceMatcher(None, original, accepted)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                continue
            old = original[i1:i2]
            new = accepted[j1:j2]
            snippets.append({"type": tag, "original": excerpt(old, 300), "accepted": excerpt(new, 300)})
        categories = {
            "amount": r"\d[\d,.]*%?|[一二三四五六七八九十百千万亿]+",
            "payment": r"付款|价款|价格|发票|税率|含税|不含税|结算|账户",
            "acceptance": r"验收|交付|交货|签收|检验|资料",
            "term": r"工作日|日历日|期限|到期|起算|日|月|年",
            "liability": r"违约|赔偿|责任|上限|免责|损失|处罚",
            "modality": r"应当|必须|不得|可以|有权|无需|有权要求",
            "termination": r"解除|终止|退出|期满|不续",
            "ip": r"知识产权|著作权|专利|商标|许可|成果",
            "confidentiality": r"保密|秘密|披露",
            "warranty": r"质保|保修|保证|瑕疵",
            "jurisdiction": r"管辖|仲裁|法院|适用法|争议",
        }
        changed = " ".join((item["original"] + " " + item["accepted"]) for item in snippets)
        result = []
        for category, pattern in categories.items():
            if re.search(pattern, changed):
                directions = []
                if any(re.search(pattern, item["original"]) for item in snippets):
                    directions.append("removed")
                if any(re.search(pattern, item["accepted"]) for item in snippets):
                    directions.append("added")
                result.append({
                    "category": category,
                    "direction": directions or ["touched"],
                    "evidence": snippets[:5],
                })
        return result

    def _build_views(self) -> dict[str, list[dict[str, Any]]]:
        views: dict[str, list[dict[str, Any]]] = {"accepted": [], "original": [], "markup": []}
        for mode in ("accepted", "original"):
            pending_join = False
            for paragraph in self.paragraphs:
                excluded_key = "accepted_excluded" if mode == "accepted" else "original_excluded"
                if paragraph[excluded_key]:
                    pending_join = False
                    continue
                text = paragraph[f"{mode}_text"]
                mark = paragraph["paragraph_mark"]
                if not text and not mark["inserted"] and not mark["deleted"]:
                    pending_join = False
                    continue
                join = pending_join and views[mode] and views[mode][-1]["flow"] == paragraph["flow"]
                if join:
                    block = views[mode][-1]
                    block["text"] = normalize_text(block["text"] + text)
                    block["paragraph_ids"].append(paragraph["id"])
                    block["locations"].append(paragraph["location"])
                else:
                    views[mode].append({
                        "mode": mode,
                        "paragraph_ids": [paragraph["id"]],
                        "locations": [paragraph["location"]],
                        "flow": paragraph["flow"],
                        "story": paragraph["story"],
                        "text": text,
                        "clause_id": paragraph["clause_id"],
                        "clause_path": paragraph["clause_path"],
                        "heading": paragraph["heading"],
                    })
                pending_join = mark["deleted"] if mode == "accepted" else mark["inserted"]
        for paragraph in self.paragraphs:
            if paragraph["markup_text"] or paragraph["hidden_text"] or paragraph["row_states"] or paragraph["cell_states"] or paragraph["comment_ids"]:
                views["markup"].append({
                    "paragraph_id": paragraph["id"],
                    "location": paragraph["location"],
                    "flow": paragraph["flow"],
                    "clause_id": paragraph["clause_id"],
                    "clause_path": paragraph["clause_path"],
                    "heading": paragraph["heading"],
                    "text": paragraph["markup_text"],
                    "comments": paragraph["comment_ids"],
                    "hidden_text": paragraph["hidden_text"],
                    "row_states": paragraph["row_states"],
                    "cell_states": paragraph["cell_states"],
                })
        return views

    def _build_clauses(self) -> None:
        clause_ids: list[str] = []
        for paragraph in self.paragraphs:
            if paragraph.get("clause_id") and paragraph["clause_id"] not in clause_ids:
                clause_ids.append(paragraph["clause_id"])
        revisions_by_clause: dict[str, list[dict[str, Any]]] = {}
        for revision in self.revisions:
            if revision.get("clause_id"):
                revisions_by_clause.setdefault(revision["clause_id"], []).append(revision)
        comments_by_clause: dict[str, list[dict[str, Any]]] = {}
        for comment in self.comments.values():
            clause = comment.get("clause_id")
            if clause:
                comments_by_clause.setdefault(clause, []).append(comment)
        for index, clause_id in enumerate(clause_ids):
            blocks = [p for p in self.paragraphs if p.get("clause_id") == clause_id]
            if not blocks:
                continue
            accepted = normalize_text("\n".join(
                block["text"] for block in self.views["accepted"]
                if block.get("clause_id") == clause_id
            ))
            original = normalize_text("\n".join(
                block["text"] for block in self.views["original"]
                if block.get("clause_id") == clause_id
            ))
            path = blocks[0].get("clause_path") or clause_id
            heading = next((block.get("heading") for block in blocks if block.get("heading")), None)
            revisions = revisions_by_clause.get(clause_id, [])
            comments = comments_by_clause.get(clause_id, [])
            formatting_types = {"runFormatting", "paragraphFormatting", "tableFormatting",
                                "rowFormatting", "cellFormatting", "tableGridFormatting",
                                "sectionFormatting", "numberingFormatting"}
            text_revisions = [item for item in revisions if item["type"] not in formatting_types | {
                "tableRowInserted", "tableRowDeleted", "tableCellInserted", "tableCellDeleted",
                "tableCellMerged",
            }]
            formatting = [item for item in revisions if item["type"] in formatting_types]
            changed = original != accepted
            signals = self._diff_signals(original, accepted)
            snippets = []
            if changed:
                matcher = difflib.SequenceMatcher(None, original, accepted)
                for opcode, i1, i2, j1, j2 in matcher.get_opcodes():
                    if opcode == "equal":
                        continue
                    snippets.append({
                        "type": opcode,
                        "original": excerpt(original[i1:i2], 300),
                        "accepted": excerpt(accepted[j1:j2], 300),
                    })
                    if len(snippets) >= 20:
                        break
            comments_ids = [item["code"] for item in comments]
            comment_brief = [
                {
                    "code": item["code"],
                    "author": item.get("author"),
                    "date": item.get("date_raw"),
                    "text": excerpt(item.get("accepted_text") or "", 300),
                    "thread_status": item.get("thread_status"),
                }
                for item in comments
            ]
            self.clauses.append({
                "clause_id": clause_id,
                "order": index + 1,
                "clause_path": path,
                "heading": heading,
                "paragraph_ids": [item["id"] for item in blocks],
                "original_text": original,
                "accepted_text": accepted,
                "changed": changed,
                "net_zero": bool(revisions) and not changed,
                "formatting_only": bool(formatting) and not changed and not text_revisions,
                "signals": signals,
                "changes": snippets,
                "revision_ids": [item["timeline_id"] for item in revisions],
                "revision_types": [item["type"] for item in revisions],
                "comment_ids": comments_ids,
                "comments": comment_brief,
                "unresolved_comment_ids": [
                    item["code"] for item in comments if item.get("thread_status") != "resolved"
                ],
                "authors": sorted({item["author"] for item in revisions if item.get("author")} |
                                   {item["author"] for item in comments if item.get("author")}),
                "date_raws": [item["date_raw"] for item in revisions if item.get("date_raw")],
                "comment_threads": len({item["code"] for item in comments if not item.get("parent_code")}),
            })

    def _build_events(self) -> None:
        events = []
        for revision in self.revisions:
            events.append({
                "kind": "revision",
                "code": revision["timeline_id"],
                "date_raw": revision["date_raw"],
                "date_normalized": revision["date_normalized"],
                "date_utc": revision["date_utc"],
                "author": revision["author"],
                "title": revision["type"],
                "clause_id": revision.get("clause_id"),
                "clause_path": revision.get("clause_path"),
                "location": revision.get("location"),
                "text": excerpt(revision.get("text") or "", 1000),
                "source_order": revision["source_order"],
                "move_group": revision.get("move_group"),
                "move_paired": revision.get("move_paired"),
                "parent_code": None,
                "_sort_ts": revision.get("_sort_ts"),
            })
        for comment in self.comments.values():
            events.append({
                "kind": "comment",
                "code": comment["code"],
                "date_raw": comment.get("date_raw"),
                "date_normalized": comment.get("date_normalized"),
                "date_utc": comment.get("date_utc"),
                "author": comment.get("author"),
                "title": "comment",
                "clause_id": comment.get("clause_id"),
                "clause_path": comment.get("clause_path"),
                "location": None,
                "text": excerpt(comment.get("accepted_text") or "", 1000),
                "source_order": 0,
                "move_group": None,
                "move_paired": None,
                "parent_code": comment.get("parent_code"),
                "_sort_ts": comment.get("_sort_ts"),
            })
        events.sort(key=lambda item: (
            item["_sort_ts"] is None,
            item["_sort_ts"] if item["_sort_ts"] is not None else 0,
            item["kind"] != "revision",
            item["source_order"],
            item["code"],
        ))
        for index, event in enumerate(events, 1):
            event["event_id"] = f"EV{index:04d}"
            event.pop("_sort_ts", None)
        self.events = events
        # 按作者和相邻间隔推断活动时段；明确标为推断，不称为谈判轮次。
        by_author: dict[str, list[dict[str, Any]]] = {}
        for event in events:
            if event.get("author") and event.get("date_normalized"):
                by_author.setdefault(event["author"], []).append(event)
        session_index = 0
        for author, author_events in sorted(by_author.items()):
            author_events.sort(key=lambda item: (item.get("date_normalized") or "", item["event_id"]))
            current = None
            previous_dt = None
            for event in author_events:
                parsed, _, _ = parse_datetime(event.get("date_normalized"))
                if parsed is None:
                    continue
                if previous_dt is None or (parsed - previous_dt).total_seconds() > self.session_gap_minutes * 60:
                    session_index += 1
                    current = {
                        "session_id": f"S{session_index:04d}",
                        "author": author,
                        "start": event["date_normalized"],
                        "end": event["date_normalized"],
                        "event_ids": [],
                        "revision_count": 0,
                        "comment_count": 0,
                        "inferred": True,
                    }
                    self.sessions.append(current)
                if current is None:
                    continue
                current["end"] = event["date_normalized"]
                current["event_ids"].append(event["event_id"])
                if event["kind"] == "revision":
                    current["revision_count"] += 1
                else:
                    current["comment_count"] += 1
                previous_dt = parsed

    def _build_anomalies(self) -> None:
        revisions = self.revisions
        comments = list(self.comments.values())
        unknown_dates = [item for item in revisions if not item.get("date_raw")]
        if unknown_dates:
            self.anomaly("medium", "revision_date_missing",
                          f"{len(unknown_dates)} 条修订没有日期，无法放入可靠时间顺序",
                          {"count": len(unknown_dates)})
        invalid_revision_dates = [
            item for item in revisions
            if item.get("date_raw") and item.get("_parsed") is None
        ]
        if invalid_revision_dates:
            self.anomaly("medium", "revision_date_invalid",
                          f"{len(invalid_revision_dates)} 条修订日期无法解析，保留原文但不强行排序",
                          {"timeline_ids": [item["timeline_id"] for item in invalid_revision_dates[:20]]})
        tz_unknown = [item for item in revisions if item.get("date_raw") and not item.get("date_timezone_aware")]
        if tz_unknown:
            self.anomaly("low", "revision_timezone_unspecified",
                          f"{len(tz_unknown)} 条修订时间没有时区，按记录值展示，不换算时区",
                          {"count": len(tz_unknown)})
        now = dt.datetime.now(dt.timezone.utc)
        future = [
            item for item in revisions
            if item.get("date_utc") and dt.datetime.fromisoformat(item["date_utc"]) > now + dt.timedelta(days=1)
        ]
        if future:
            self.anomaly("medium", "revision_future",
                          f"{len(future)} 条修订日期晚于当前时间，可能来自错误系统时钟",
                          {"timeline_ids": [item["timeline_id"] for item in future[:20]]})
        created, _, _ = parse_datetime(self.core.get("created"))
        if created:
            created_timestamp = sorting_ts(created)
            before = [
                item for item in revisions
                if created_timestamp is not None and item.get("_sort_ts") is not None
                and item["_sort_ts"] < created_timestamp - 86400
            ]
            if before:
                self.anomaly("medium", "revision_before_created",
                              f"{len(before)} 条修订早于文档创建时间，时间线不可直接信",
                              {"timeline_ids": [item["timeline_id"] for item in before[:20]]})
        if revisions and self.settings.get("track_changes") is not None and not self.settings.get("track_changes"):
            self.anomaly("low", "track_changes_off_but_history",
                          "当前跟踪修改开关为关闭，但文件仍保留历史修订")
        comment_dates_missing = [item for item in comments if not item.get("date_raw")]
        if comment_dates_missing:
            self.anomaly("medium", "comment_date_missing",
                          f"{len(comment_dates_missing)} 条批注没有日期，无法放入可靠时间线",
                          {"comments": [item.get("raw_id") for item in comment_dates_missing[:50]]})
        invalid_comment_dates = [
            item for item in comments
            if item.get("date_raw") and item.get("_parsed") is None
        ]
        if invalid_comment_dates:
            self.anomaly("medium", "comment_date_invalid",
                          f"{len(invalid_comment_dates)} 条批注日期无法解析，保留原文但不强行排序",
                          {"comments": [item.get("code") for item in invalid_comment_dates[:50]]})
        future_comments = [
            item for item in comments
            if item.get("date_utc") and dt.datetime.fromisoformat(item["date_utc"]) > now + dt.timedelta(days=1)
        ]
        if future_comments:
            self.anomaly("medium", "comment_future",
                          f"{len(future_comments)} 条批注日期晚于当前时间，可能来自错误系统时钟",
                          {"comments": [item.get("code") for item in future_comments[:20]]})
        if self.settings.get("track_changes") is None:
            self.anomaly("low", "track_changes_setting_missing",
                          "未能读取 Word 的跟踪修改开关；不据此判断修订是否正在记录")
        if not revisions and not comments:
            self.anomaly("info", "no_revisions_or_comments",
                          "文件中没有可识别的修订或批注；这不等同于合同没有风险")
        table_structural = [
            item for item in revisions
            if item["type"].startswith("table") or item["type"] in {"tableCellMerged", "tableCellInserted", "tableCellDeleted"}
        ]
        if table_structural:
            self.anomaly("low", "table_structural_views",
                          "存在表格结构修订；阅读视图保留单元格文本，但行/列布局以 Word 接受/拒绝结果为准",
                          {"count": len(table_structural)})
        if self.objects.get("embedded_objects") or self.objects.get("alt_chunks"):
            self.anomaly("medium", "unreadable_objects",
                          "检测到嵌入对象或 altChunk，文本阅读器未解析其内部内容",
                          self.objects)
        if self.objects.get("drawings") or self.objects.get("pictures"):
            self.anomaly("low", "visual_objects",
                          "检测到图片/绘图对象；本报告不包含图片中的文字或视觉内容",
                          self.objects)
        if self.objects.get("fields") or self.objects.get("field_instructions"):
            self.anomaly("low", "fields_present",
                          "检测到字段/域；视图读取显示结果，不保证域代码计算结果与 Word 当前一致",
                          self.objects)
        if self.objects.get("content_controls"):
            self.anomaly("low", "content_controls_present",
                          "检测到内容控件；控件占位符或绑定值可能影响 Word 中的最终显示",
                          self.objects)
        hidden_texts = [item for item in self.paragraphs if item.get("hidden_text")]
        if hidden_texts:
            self.anomaly("medium", "hidden_text",
                          f"{len(hidden_texts)} 个段落含直接或样式判定的隐藏文字；默认不计入阅读视图",
                          {"paragraph_ids": [item["id"] for item in hidden_texts[:30]]})
        if self.security.get("blockers"):
            relevant = [
                item for item in self.security["blockers"]
                if "现代批注扩展部件" not in item
            ]
            if relevant:
                self.anomaly("high", "package_protection",
                              "文件包含宏、签名、编辑保护或其他需注意的部件；仅做只读解析，不移除",
                              relevant)
        if self.security.get("namespace") == redline.STRICT_W_NS:
            self.anomaly("medium", "strict_ooxml",
                          "文件使用 Strict OOXML 命名空间；阅读器按本地元素名解析，Word 客户端兼容性需复核")
        unresolved_or_unknown = [
            item for item in comments
            if item.get("thread_status") in {"open", "unknown"} and not item.get("parent_code")
        ]
        if unresolved_or_unknown:
            self.anomaly("medium", "comment_status_not_resolved",
                          f"{len(unresolved_or_unknown)} 个批注线程未明确解决或没有现代解决状态",
                          {"comments": [item["code"] for item in unresolved_or_unknown[:50]]})
        net_zero = [item for item in self.clauses if item.get("net_zero")]
        if net_zero:
            self.anomaly("medium", "net_zero_changes",
                          f"{len(net_zero)} 个条款虽有修订记录但接受后文字与原文相同，可能发生撤销或仅格式变化",
                          {"clauses": [item["clause_id"] for item in net_zero]})
        if not self.comments and self.anchors:
            self.anomaly("medium", "comments_missing_definitions",
                          "检测到批注锚点但缺少 comments.xml 批注正文", {"count": len(self.anchors)})
        if self.modern_comment_parts and any(item.get("resolved") is None for item in comments):
            self.anomaly("low", "legacy_resolution_unknown",
                          "部分批注没有现代解决状态；旧式批注无法可靠判断是否已解决")

    def summary(self) -> dict[str, Any]:
        authors = sorted({
            author for item in self.revisions if (author := item.get("author"))
        } | {
            author for item in self.comments.values() if (author := item.get("author"))
        })
        dates = [
            value for item in self.revisions if (value := item.get("date_normalized"))
        ] + [
            value for item in self.comments.values() if (value := item.get("date_normalized"))
        ]
        dated_values = []
        for value in dates:
            parsed_value, _, _ = parse_datetime(value)
            timestamp = sorting_ts(parsed_value)
            if timestamp is not None:
                dated_values.append((timestamp, value))
        changed_clauses = [item for item in self.clauses if item.get("changed")]
        unresolved = [
            item for item in self.comments.values()
            if item.get("thread_status") in {"open", "unknown"} and not item.get("parent_code")
        ]
        return {
            "schema": SCHEMA,
            "security": self.security,
            "file": {
                "path": str(self.path),
                "sha256": self.original_sha256,
                "size": self.file_size,
                "mtime": self.file_mtime,
            },
            "core": self.core,
            "app": self.app,
            "settings": self.settings,
            "coverage": {
                "parts": self.story_parts,
                "paragraphs": len(self.paragraphs),
                "tables": sum(1 for item in self.paragraphs if item.get("table_id")),
                "objects": self.objects,
                "modern_comment_parts": getattr(self, "modern_comment_parts", []),
                "hidden_paragraphs": sum(1 for item in self.paragraphs if item.get("hidden_text")),
                "comment_resolution": "modern" if getattr(self, "modern_comment_parts", []) else "unknown_or_legacy",
            },
            "counts": {
                "revisions": len(self.revisions),
                "revision_types": self._count_types(item["type"] for item in self.revisions),
                "comments": len(self.comments),
                "comment_threads": len({item["code"] for item in self.comments.values() if not item.get("parent_code")}),
                "unresolved_or_unknown_comments": len(unresolved),
                "clauses": len(self.clauses),
                "changed_clauses": len(changed_clauses),
                "net_zero_clauses": len([item for item in self.clauses if item.get("net_zero")]),
                "anomalies": len(self.anomalies),
            },
            "authors": authors,
            "date_range": {
                "first": min(dated_values)[1] if dated_values else None,
                "last": max(dated_values)[1] if dated_values else None,
            },
            "sessions": self.sessions,
            "anomalies": self.anomalies,
            "clauses": self.clauses,
            "paragraphs": self.paragraphs,
            "moves": self.moves,
            "comments": [
                {key: value for key, value in item.items() if not key.startswith("_")}
                for item in self.comments.values()
            ],
            "revisions": [
                {key: value for key, value in item.items() if not key.startswith("_")}
                for item in self.revisions
            ],
            "events": self.events,
            "views": {
                "accepted": self.views["accepted"],
                "original": self.views["original"],
                "markup": self.views["markup"],
            },
        }

    @staticmethod
    def _count_types(values: Any) -> dict[str, int]:
        result: dict[str, int] = {}
        for value in values:
            result[value] = result.get(value, 0) + 1
        return result


def md_cell(value: Any) -> str:
    if value is None:
        return "—"
    return str(value).replace("|", "\\|").replace("\n", " ").strip()


def render_report(data: dict[str, Any]) -> str:
    summary = data
    counts = summary["counts"]
    file_info = summary["file"]
    core = summary["core"]
    lines = [
        "# Word 合同深读报告（事实层）",
        "",
        "> 本报告由只读 OOXML 解析生成；接受/拒绝视图是阅读分析，不写回 Word。"
        "修订和批注的“意图”不能仅凭记录推定，业务结论须结合交易背景和用户确认。",
        "",
        "## 一、文件和阅读口径",
        "",
        f"- 文件：`{Path(file_info['path']).name}`",
        f"- SHA-256：`{file_info['sha256']}`",
        f"- 大小：{file_info['size']} 字节；文件修改时间：{file_info['mtime']}",
        f"- 作者/最后保存者：{core.get('creator') or '未记录'} / {core.get('last_modified_by') or '未记录'}",
        f"- 创建/最后修改（文件核心属性）：{core.get('created') or '未记录'} / {core.get('modified') or '未记录'}",
        f"- 当前跟踪修改开关：{summary['settings'].get('track_changes') if summary['settings'].get('track_changes') is not None else '未知'}",
        "",
        "阅读视图：`accepted.md` 为接受全部修订、`original.md` 为拒绝全部修订、`markup.md` 为带修订/批注内联标记的审阅视图。"
        "它们都是分析产物；Word/WPS 当前显示模式、自动编号和复杂表格布局仍须在客户端确认。",
        "",
        "## 二、事实摘要",
        "",
        f"- 修订：**{counts['revisions']}** 条；批注：**{counts['comments']}** 条；条款组：**{counts['clauses']}** 个。",
        f"- 有文字变化的条款组：**{counts['changed_clauses']}**；无净文字变化的修订条款组：**{counts['net_zero_clauses']}**。",
        f"- 未解决或状态未知的批注线程：**{counts['unresolved_or_unknown_comments']}**。",
        f"- 修订类型：{', '.join(f'{key}={value}' for key, value in counts['revision_types'].items()) or '无'}。",
        f"- 作者/参与者：{', '.join(summary['authors']) or '未记录'}。",
        f"- 修订/批注日期范围：{summary['date_range']['first'] or '未知'} 至 {summary['date_range']['last'] or '未知'}。",
        "",
        "## 三、修改时间线",
        "",
        "### 作者活动时段（按时间间隔推断，不是谈判轮次）",
        "",
        "| 时段 | 作者 | 开始 | 结束 | 修订 | 批注 |",
        "|---|---|---|---|---:|---:|",
    ]
    for session in summary["sessions"]:
        lines.append(
            f"| {md_cell(session['session_id'])} | {md_cell(session['author'])} | "
            f"{md_cell(session['start'])} | {md_cell(session['end'])} | "
            f"{session['revision_count']} | {session['comment_count']} |"
        )
    if not summary["sessions"]:
        lines.append("| — | — | — | — | 0 | 0 |")
    lines.extend([
        "",
        "### 事件顺序（精确记录值）",
        "",
        "| 事件 | 时间（原文） | 作者 | 类型 | 条款 | 内容摘要 |",
        "|---|---|---|---|---|---|",
    ])
    for event in summary["events"][:500]:
        lines.append(
            f"| {md_cell(event['event_id'])} | {md_cell(event['date_raw'])} | "
            f"{md_cell(event['author'])} | {md_cell(event['title'])} | "
            f"{md_cell(event.get('clause_path') or event.get('clause_id'))} | "
            f"{md_cell(excerpt(event.get('text') or '', 180))} |"
        )
    if len(summary["events"]) > 500:
        lines.append("")
        lines.append(f"> 报告只展示前 500 个事件；完整事件见 `deep-read.json` / `timeline.csv`（共 {len(summary['events'])} 个）。")
    lines.extend([
        "",
        "## 四、条款变化链",
        "",
        "> “业务信号”按关键词和文本差异提取，是阅读提示，不是法律定性。条款原文、作者、日期和批注必须逐项回看。",
        "",
    ])
    changed = [item for item in summary["clauses"] if item.get("changed")]
    if not changed:
        lines.append("没有检测到接受视图与原文视图之间的文字差异。")
    for clause in changed:
        lines.append(f"### {clause['clause_path']}")
        lines.append("")
        lines.append(f"- 原文：{excerpt(clause['original_text'], 700) or '（空）'}")
        lines.append(f"- 接受后：{excerpt(clause['accepted_text'], 700) or '（空）'}")
        lines.append(f"- 修订：{', '.join(clause['revision_ids']) or '无'}；批注：{', '.join(clause['comment_ids']) or '无'}")
        for change in clause.get("changes", [])[:10]:
            if change.get("original"):
                lines.append(f"- 变更前片段：{excerpt(change['original'], 300)}")
            if change.get("accepted"):
                lines.append(f"- 变更后片段：{excerpt(change['accepted'], 300)}")
        if clause.get("signals"):
            lines.append("- 业务信号：" + "；".join(
                f"{item['category']}={','.join(item['direction'])}" for item in clause["signals"]
            ))
        if clause.get("comments"):
            lines.append("- 批注涉及：" + "；".join(
                excerpt(item.get("text", ""), 160) for item in clause["comments"][:4]
            ))
        lines.append("")
    lines.extend([
        "## 五、批注与未决事项",
        "",
        "| 批注 | 父批注 | 作者 | 日期 | 状态 | 锚定文字（接受视图） | 正文 |",
        "|---|---|---|---|---|---|---|",
    ])
    for comment in summary["comments"]:
        state = {"resolved": "已解决", "open": "未解决", "unknown": "状态未知"}.get(
            comment.get("thread_status"), "状态未知"
        )
        lines.append(
            f"| {md_cell(comment.get('code'))} | {md_cell(comment.get('parent_code') or '—')} | "
            f"{md_cell(comment.get('author'))} | "
            f"{md_cell(comment.get('date_raw'))} | {state} | "
            f"{md_cell(excerpt(comment.get('anchor_accepted_text', ''), 100))} | "
            f"{md_cell(excerpt(comment.get('accepted_text', ''), 180))} |"
        )
    if not summary["comments"]:
        lines.append("| — | — | — | — | — | — | 无批注正文 |")
    lines.extend([
        "",
        "## 六、异常与证据缺口",
        "",
        "| 严重度 | 类型 | 说明 |",
        "|---|---|---|",
    ])
    for item in summary["anomalies"]:
        lines.append(f"| {md_cell(item['severity'])} | {md_cell(item['kind'])} | {md_cell(item['message'])} |")
    if not summary["anomalies"]:
        lines.append("| — | — | 未检测到结构异常 |")
    lines.extend([
        "",
        "## 七、业务逻辑阅读框架",
        "",
        "先读 `original.md` 和 `accepted.md`，再用 `markup.md` 和本时间线回看每条修订，最后才形成业务判断。"
        "建议按以下顺序重建：",
        "",
        "1. **交易骨架**：当事人、标的/服务、价款与税、交付/验收、付款、质保、责任、解除、争议解决是否被改动？",
        "2. **风险迁移**：改动把哪一类风险从我方转给对方，或反向转移？是否有责任上限、免责、期限或证据条件？",
        "3. **对价与条件**：金额/比例/税的变化是否与履行范围、验收、付款节点和发票条件同步？",
        "4. **时间线**：谁在何时改动、是否先评论后修改、是否存在撤销/回改、是否有未解决批注？",
        "5. **文件层级**：若还有主协议、订单、附件或其他版本，必须另行勾稽，不能只看这一份。",
        "",
        "业务结论必须区分：**文件事实**（OOXML 实际记录）、**阅读推断**（作者意图/谈判目的）和**法律评价**；"
        "仅有批注或修订不证明双方已同意该业务安排。",
    ])
    return "\n".join(lines) + "\n"


def render_accepted_or_original(data: dict[str, Any], mode: str) -> str:
    title = "接受全部修订" if mode == "accepted" else "拒绝全部修订（原文）"
    lines = [
        f"# {title}阅读视图",
        "",
        "> 由 OOXML 重建的分析视图，不写回 Word。复杂编号、域、表格布局须在 Word/WPS 中确认。",
        "",
    ]
    current_clause = None
    for block in data["views"][mode]:
        if not block.get("text") and not block.get("heading"):
            continue
        heading = block.get("heading")
        if heading and block.get("clause_path") != current_clause:
            lines.append("")
            lines.append(f"## {heading.get('title')}")
            current_clause = block.get("clause_path")
        elif block.get("clause_path") and block["clause_path"] != current_clause and not heading:
            lines.append("")
            lines.append(f"## {block['clause_path']}")
            current_clause = block["clause_path"]
        if block.get("text"):
            location = block["locations"][0] if block.get("locations") else ""
            text = block["text"]
            if heading and text == heading.get("title"):
                continue
            if heading and text.startswith(heading.get("title") or ""):
                text = text[len(heading.get("title") or ""):].strip()
            if text:
                lines.append(f"[{block['paragraph_ids'][0]} · {location}] {text}")
    return "\n".join(lines) + "\n"


def render_markup(data: dict[str, Any]) -> str:
    lines = [
        "# 修订与批注审阅视图",
        "",
        "> 修订标记按 OOXML 记录重建；`⟦插入/删除/移入/移出…⟧` 是阅读标记，不是 Word 的接受状态。",
        "",
    ]
    current_clause = None
    for item in data["views"]["markup"]:
        if item.get("clause_path") and item["clause_path"] != current_clause:
            lines.append("")
            lines.append(f"## {item['clause_path']}")
            current_clause = item["clause_path"]
        notes = []
        if item.get("row_states") or item.get("cell_states"):
            notes.append("表格行/单元格有结构修订")
        if item.get("hidden_text"):
            notes.append("含隐藏文字，见下")
        suffix = f" 〔批注：{', '.join(item['comments'])}〕" if item.get("comments") else ""
        if notes:
            suffix += f" 〔{'；'.join(notes)}〕"
        lines.append(f"[{item['paragraph_id']} · {item['location']}] {item['text']}{suffix}")
        if item.get("hidden_text"):
            lines.append(f"  - 隐藏文字：{item['hidden_text']}")
    return "\n".join(lines) + "\n"


def write_outputs(data: dict[str, Any], output_dir: Path, *, force: bool) -> dict[str, Any]:
    targets = {
        "deep-read.json": json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        "deep-read.md": render_report(data),
        "accepted.md": render_accepted_or_original(data, "accepted"),
        "original.md": render_accepted_or_original(data, "original"),
        "markup.md": render_markup(data),
    }
    output_dir.mkdir(parents=True, exist_ok=True)
    existing = [name for name in targets if (output_dir / name).exists()]
    if existing and not force:
        raise ReaderError(f"输出文件已存在：{', '.join(existing)}；确认后用 --force 覆盖")
    temporary = Path(tempfile.mkdtemp(prefix=".word-read-", dir=output_dir))
    try:
        for name, content in targets.items():
            (temporary / name).write_text(content)
        # CSV 单独写，保留 Excel 对中文的兼容 BOM。
        with (temporary / "timeline.csv").open("w", newline="", encoding="utf-8-sig") as stream:
            writer = csv.writer(stream)
            writer.writerow(["event_id", "kind", "code", "date_raw", "date_utc", "author",
                             "type", "clause_id", "clause_path", "location", "text", "parent_code", "move_group", "move_paired"])
            for event in data["events"]:
                writer.writerow([
                    event.get("event_id"), event.get("kind"), event.get("code"),
                    event.get("date_raw"), event.get("date_utc"), event.get("author"),
                    event.get("title"), event.get("clause_id"), event.get("clause_path"),
                    event.get("location"), event.get("text"), event.get("parent_code"),
                    event.get("move_group"), event.get("move_paired"),
                ])
        with (temporary / "comments.csv").open("w", newline="", encoding="utf-8-sig") as stream:
            writer = csv.writer(stream)
            writer.writerow(["code", "raw_id", "parent_code", "author", "date_raw", "thread_status",
                             "clause_id", "clause_path", "anchor_accepted", "accepted_text"])
            for comment in data["comments"]:
                writer.writerow([
                    comment.get("code"), comment.get("raw_id"), comment.get("parent_code"),
                    comment.get("author"), comment.get("date_raw"), comment.get("thread_status"),
                    comment.get("clause_id"), comment.get("clause_path"),
                    comment.get("anchor_accepted_text"), comment.get("accepted_text"),
                ])
        for name in targets:
            os.replace(temporary / name, output_dir / name)
        os.replace(temporary / "timeline.csv", output_dir / "timeline.csv")
        os.replace(temporary / "comments.csv", output_dir / "comments.csv")
    finally:
        import shutil
        shutil.rmtree(temporary, ignore_errors=True)
    return {
        "output": str(output_dir),
        "files": list(targets) + ["timeline.csv", "comments.csv"],
        "input_sha256": data["file"]["sha256"],
        "counts": data["counts"],
        "date_range": data["date_range"],
    }


def analyze(path: Path, output_dir: Path | None = None, *, include_hidden: bool = False,
            session_gap_minutes: int = 30, force: bool = False,
            max_report_events: int = 500) -> dict[str, Any]:
    reader = WordReader(
        path, include_hidden=include_hidden,
        session_gap_minutes=session_gap_minutes,
        max_report_events=max_report_events,
    )
    reader.load()
    data = reader.summary()
    if sha256_file(path) != data["file"]["sha256"]:
        raise ReaderError("分析期间输入文件发生变化；请重新开始，不把混合版本结果写入产出")
    if output_dir is not None:
        data["outputs"] = write_outputs(data, output_dir, force=force)
    return data


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Word 合同精细阅读器（只读）")
    sub = parser.add_subparsers(dest="command", required=True)
    command = sub.add_parser("analyze", help="分析 DOCX 并生成阅读视图、时间线和报告")
    command.add_argument("file")
    command.add_argument("--out", required=True, help="任务产出目录")
    command.add_argument("--include-hidden", action="store_true", help="将直接/样式判定的隐藏文字纳入视图")
    command.add_argument("--session-gap-minutes", type=int, default=30)
    command.add_argument("--max-report-events", type=int, default=500)
    command.add_argument("--force", action="store_true", help="允许覆盖同名产出文件")
    command.set_defaults(handler=lambda args: analyze(Path(args.file).expanduser(), Path(args.out).expanduser(),
                                                      include_hidden=args.include_hidden,
                                                      session_gap_minutes=args.session_gap_minutes,
                                                      max_report_events=args.max_report_events,
                                                      force=args.force))
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        result = args.handler(args)
        summary = {
            "schema": SCHEMA,
            "ok": True,
            "counts": result.get("counts"),
            "date_range": result.get("date_range"),
            "outputs": result.get("outputs"),
        }
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 0
    except (ReaderError, redline.PackageError) as exc:
        print(json.dumps({"schema": SCHEMA, "ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
        return 2
    except Exception as exc:
        print(json.dumps({
            "schema": SCHEMA, "ok": False,
            "error": f"内部错误（{type(exc).__name__}）：{exc}",
        }, ensure_ascii=False, indent=2))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
