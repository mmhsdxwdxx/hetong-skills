#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hetong_redline.py — 受限、可验证的 Word 红线（修订 + 批注）引擎。

只依赖 Python 3.9+ 标准库（zipfile + xml.dom.minidom），不安装任何第三方包。
设计边界（首版刻意保守）：

- 只处理 word/document.xml 中「正文段落」与「普通表格单元格段落」的单段文字；
  段落必须足够简单：只含 pPr 和普通 w:r，且每个 run 只含 rPr 与 w:t。
- 真正写入 w:ins / w:del / w:delText 与传统 word/comments.xml 批注；
  不生成伪红色字体，不伪造 Word/WPS 实测结论。
- 未修改的 ZIP 成员按原内容字节写回；命名空间前缀与根声明原样保留，
  不使用 pretty-print。
- 原件只读：apply 永不 in-place，输出文件必须不存在，原子独占发布，
  全量内存预检 + 自动验证，任何一步失败都不产生文件。
- 不处理：表格行增删、段落增删、格式修订、图片/域/超链接/书签/sdt/文本框等
  复杂结构；run 的 rPr 直接声明隐藏文本（vanish/specVanish/webHidden）视为
  隐藏条款并 fail-closed；遇到即拒绝。通过样式继承的隐藏文本未解析，另以
  warning 披露该边界。

子命令
------
  inspect input.docx
      输出 JSON：sha256、全部段落（RL0001 独立编号）与当前可见文本、
      可编辑性与阻断原因、既有修订/批注概览、包级安全警告。

  plan input.docx edits.json -o plan.json
      校验编辑意图并生成计划（不写 docx）。edits.json 格式::

          [
            {"id": "R1", "para": "RL0003",
             "old": "原句（段内唯一精确子串）", "new": "目标句",
             "reason": "改为……的理由（会成为范围批注）",
             "broad_reason": "仅在大幅改写被警告时必填"},
            {"id": "C1", "para": "RL0005",
             "old": "批注锚文本（段内唯一精确子串）",
             "comment": "批注内容"}
          ]

      规则：old 非空、段内唯一；同一段落只允许一个 op；new 不得等于 old
      （new 为空表示删除 old 片段）；替换必须给 reason，仅批注用 comment
      （comment-only 不产生修订）。一个 op 只生成一条批注，不逐个小 diff 重复。

  apply input.docx plan.json -o result.docx --author '合同审阅'
      重新校验源文件 sha256、编辑内容与粒度（不信任 plan 里的定位或 XML），
      重新规划并与 plan 比对；通过后一次性生成并原子发布 result.docx，
      发布前自动执行 verify。

  verify input.docx result.docx plan.json
      校验 ZIP/XML/关系/内容类型、新增修订与批注闭环、接受/拒绝视图等价、
      未允许变更的 ZIP 成员字节不变；并以结果中唯一的新增 author/date 重建
      完整预期输出包，逐成员比对解压字节（重建一致性判据，非语义证明）。
      输出 JSON，失败退出码 3。

退出码：0 成功；1 参数/编辑/计划错误；2 文件/包结构错误；3 验证失败。
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import posixpath
import re
import sys
import tempfile
import zipfile
from collections import Counter
from datetime import datetime, timezone
from difflib import SequenceMatcher
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from typing import TypeGuard
else:  # Python 3.9 运行时不使用类型守卫；算法不依赖该名称
    TypeGuard = Any
from xml.dom import Node
from xml.dom import minidom
from xml.parsers.expat import ExpatError
from xml.sax.saxutils import escape as xml_escape
from xml.sax.saxutils import quoteattr

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STRICT_W_NS = "http://purl.oclc.org/ooxml/wordprocessingml/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COMMENTS_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
COMMENTS_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)

TOOL = "hetong_redline"
PLAN_VERSION = 1

REV_TAGS = {"ins", "del", "moveFrom", "moveTo"}
COMMENT_ANCHOR_TAGS = {"commentRangeStart", "commentRangeEnd", "commentReference"}
# 直接写在 w:rPr 里的隐藏文本开关；缺 w:val 即为生效
HIDDEN_TOGGLES = ("vanish", "specVanish", "webHidden")
OFF_VALUES = {"0", "false", "off"}
MODERN_COMMENT_PARTS = (
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
    "word/people.xml",
)

# 复杂段落 / 祖先阻断清单
BLOCKED_ANCESTORS = {
    "sdt",
    "sdtContent",
    "txbxContent",
    "customXml",
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "smartTag",
    "fldSimple",
    "hyperlink",
    "drawing",
    "pict",
    "object",
    "AlternateContent",
    "choice",
    "fallback",
}

ALLOWED_CHANGED_MEMBERS = {
    "word/document.xml",
    "word/comments.xml",
    "word/_rels/document.xml.rels",
    "[Content_Types].xml",
}
ALLOWED_ADDED_MEMBERS = {"word/comments.xml"}

STORY_PART_RE = re.compile(
    r"^word/(document|footnotes|endnotes|comments|header\d*|footer\d*)\.xml$"
)

MAX_ZIP_BYTES = 200 * 1024 * 1024
MAX_TOTAL_UNCOMPRESSED = 512 * 1024 * 1024
MAX_MEMBERS = 5000
MAX_PART_BYTES = 64 * 1024 * 1024

NUMBER_RE = re.compile(r"\d[\d,]*(?:\.\d+)?%?")
PARA_ID_RE = re.compile(r"^RL\d{4,}$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z$")


# --------------------------------------------------------------------------
# 错误
# --------------------------------------------------------------------------


class RedlineError(Exception):
    """带 CLI 退出码的受控错误。"""

    code = 1

    def __init__(self, message, **detail):
        super().__init__(message)
        self.message = message
        self.detail = detail

    def as_dict(self):
        out = {"ok": False, "error": self.message, "code": self.code}
        if self.detail:
            out["detail"] = self.detail
        return out


class InputError(RedlineError):
    code = 1


class PackageError(RedlineError):
    code = 2


class VerifyError(RedlineError):
    code = 3


# --------------------------------------------------------------------------
# 通用小工具
# --------------------------------------------------------------------------


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def utc_now_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def canonical_json(obj):
    return json.dumps(
        obj, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )


def json_dumps_pretty(obj):
    return json.dumps(obj, ensure_ascii=False, indent=2) + "\n"


def is_element(node) -> TypeGuard[minidom.Element]:
    return node.nodeType == Node.ELEMENT_NODE


def local_name(node):
    tag = getattr(node, "tagName", None)
    if not tag:
        return None
    return tag.rsplit(":", 1)[-1]


def w_prefix_ok(el, w_prefix):
    """元素是否属于 WordprocessingML（按前缀判断；无前缀按默认命名空间处理）。"""
    tag = el.tagName
    if ":" in tag:
        return tag.split(":", 1)[0] == w_prefix
    return True


def element_children(node):
    return [ch for ch in node.childNodes if is_element(ch)]


def iter_elements_in_order(root):
    """先序遍历元素（不包含 root 自身）。"""
    out = []
    stack = [root]
    while stack:
        cur = stack.pop()
        if cur is not root:
            out.append(cur)
        children = element_children(cur)
        for ch in reversed(children):
            stack.append(ch)
    return out


def text_of(node):
    parts = []
    for ch in node.childNodes:
        if ch.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE):
            parts.append(ch.data)
    return "".join(parts)


def attr_local(el, local):
    """按本地名取属性，兼容不同前缀。"""
    attrs = el.attributes
    for i in range(attrs.length):
        a = attrs.item(i)
        name = a.name.rsplit(":", 1)[-1]
        if name == local:
            return a.value
    return None


def _on_off_enabled(el):
    """ST_OnOff 语义：缺 w:val 为开，0/false/off 为关。"""
    val = attr_local(el, "val")
    if val is None:
        return True
    return val.strip().lower() not in OFF_VALUES


def parse_xml(data, label):
    if not isinstance(data, (bytes, bytearray)):
        raise PackageError("%s：内部错误，XML 不是字节" % label)
    lowered = bytes(data).lower()
    # 解析前拒绝 DTD/实体，避免实体扩展与外部实体风险
    if b"<!doctype" in lowered:
        raise PackageError("%s：不支持带 DOCTYPE 的 XML" % label)
    if b"<!entity" in lowered:
        raise PackageError("%s：不支持带 ENTITY 声明的 XML" % label)
    if lowered.startswith((b"\xff\xfe", b"\xfe\xff")):
        try:
            text_lower = bytes(data).decode("utf-16").lower()
        except UnicodeDecodeError:
            text_lower = ""
        if "<!doctype" in text_lower:
            raise PackageError("%s：不支持带 DOCTYPE 的 XML" % label)
        if "<!entity" in text_lower:
            raise PackageError("%s：不支持带 ENTITY 声明的 XML" % label)
    try:
        return minidom.parseString(bytes(data))
    except ExpatError as exc:
        raise PackageError("%s：XML 解析失败：%s" % (label, exc)) from exc
    except ValueError as exc:
        raise PackageError("%s：XML 解析失败：%s" % (label, exc)) from exc


def serialize_like(original_bytes, doc):
    """序列化 XML 树，保留原 XML 声明（编码统一改写为 UTF-8）。

    不使用 pretty-print；元素前缀、命名空间声明均由 minidom 原样保留。
    """
    bom = b"\xef\xbb\xbf" if original_bytes.startswith(b"\xef\xbb\xbf") else b""
    m = re.match(rb"\A(?:\xef\xbb\xbf)?(\s*<\?xml[^>]*\?>)", original_bytes)
    if m:
        decl = m.group(1)
        # 统一编码声明，避免原声明与 UTF-8 输出不一致
        if re.search(rb"encoding\s*=", decl, re.I):
            decl = re.sub(
                rb"encoding\s*=\s*[\"'][^\"']*[\"']",
                b'encoding="UTF-8"',
                decl,
                count=1,
                flags=re.I,
            )
        else:
            decl = decl.replace(
                b"<?xml", b'<?xml version="1.0" encoding="UTF-8"', 1
            ) if b"version" not in decl.lower() else decl
            if b"encoding" not in decl.lower():
                decl = re.sub(
                    rb"(\?>)", b' encoding="UTF-8"\\1', decl, count=1
                )
    else:
        decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    body = doc.documentElement.toxml().encode("utf-8")
    return bom + decl + body


def resolve_w_prefix(root):
    """返回 (prefix, namespace_uri)；仅接受 transitional / strict W 命名空间。

    以 documentElement.namespaceURI（解析器判定的命名空间）为准。若只看根元素
    上的 xmlns:w 声明，本函数为默认命名空间 transitional 文档补写的声明会在
    下一次调用时把真实的 Strict 命名空间遮蔽成 transitional，导致 Strict 阻断
    失效。Strict 一律返回真实前缀（默认命名空间时为 ""）且不写回 DOM，使
    analyze_package / DocContext 能稳定判出 Strict 并 fail-closed；默认命名空间
    transitional 仍补 xmlns:w，保持既有生成语义。
    """
    tag = root.tagName or ""
    declared_prefix = tag.split(":", 1)[0] if ":" in tag else ""
    uri = getattr(root, "namespaceURI", None)
    if uri in (W_NS, STRICT_W_NS):
        if declared_prefix:
            return declared_prefix, uri
        if uri == STRICT_W_NS:
            # Strict：保持默认命名空间原样，不加声明，交给上层阻断
            return "", STRICT_W_NS
        # 默认命名空间 transitional：为新建 w: 前缀元素补声明（既有语义）
        root.setAttribute("xmlns:w", W_NS)
        return "w", W_NS

    # 回退：解析器未提供 namespaceURI 时按根元素声明查找
    if declared_prefix:
        uri = root.getAttribute("xmlns:" + declared_prefix)
        if uri in (W_NS, STRICT_W_NS):
            return declared_prefix, uri
    attrs = root.attributes
    for i in range(attrs.length):
        a = attrs.item(i)
        if a.name.startswith("xmlns:") and a.value in (W_NS, STRICT_W_NS):
            return a.name.split(":", 1)[1], a.value
    default_uri = root.getAttribute("xmlns")
    if default_uri in (W_NS, STRICT_W_NS):
        if default_uri == STRICT_W_NS:
            return "", STRICT_W_NS
        root.setAttribute("xmlns:w", W_NS)
        return "w", default_uri
    raise PackageError("无法识别 word/document.xml 的 WordprocessingML 命名空间")


def wq(prefix, local):
    return (prefix + ":" + local) if prefix else local


def prefix_for_ns(root, uri):
    """返回根元素上绑定到 uri 的前缀；默认命名空间返回 ""。"""
    tag = root.tagName
    if ":" in tag:
        p = tag.split(":", 1)[0]
        if root.getAttribute("xmlns:" + p) == uri:
            return p
    attrs = root.attributes
    for i in range(attrs.length):
        a = attrs.item(i)
        if a.name.startswith("xmlns:") and a.value == uri:
            return a.name.split(":", 1)[1]
    if root.getAttribute("xmlns") == uri:
        return ""
    return None


def set_wattr(el, prefix, local, value):
    el.setAttribute(wq(prefix, local), value)


# --------------------------------------------------------------------------
# ZIP 包层
# --------------------------------------------------------------------------


class Package:
    def __init__(self, label, raw, members, infos, comment, issues):
        self.label = label
        self.raw = raw
        self.members = members  # name -> bytes（未压缩内容）
        self.infos = infos
        self.comment = comment
        self.issues = issues  # 包级问题（重复成员、异常路径等）
        self.sha256 = sha256_bytes(raw)

    @property
    def names(self):
        return [i.filename for i in self.infos]

    def has(self, name):
        return name in self.members

    def get(self, name):
        try:
            return self.members[name]
        except KeyError:
            raise PackageError("%s：缺少必需部件 %s" % (self.label, name)) from None


def _zip_issues(pkg, infos):
    issues = []
    names = [i.filename for i in infos]
    counts = Counter(names)
    if len(counts) != len(names):
        dupes = sorted(n for n, c in counts.items() if c > 1)
        issues.append("重复 ZIP 成员：%s" % "、".join(dupes[:5]))
    for name in names:
        if name.endswith("/"):
            continue
        norm = posixpath.normpath(name)
        if (
            name.startswith("/")
            or "\\" in name
            or norm != name
            or ".." in name.split("/")
            or re.match(r"^[A-Za-z]:", name)
        ):
            issues.append("异常 ZIP 路径：%r" % name)
    return issues


def package_from_bytes(raw, label):
    if len(raw) > MAX_ZIP_BYTES:
        raise PackageError("%s：ZIP 过大（%d 字节）" % (label, len(raw)))
    if raw.startswith(b"\xd0\xcf\x11\xe0"):
        raise PackageError("%s：加密文档或旧版二进制 .doc，不支持" % label)
    if not zipfile.is_zipfile(io.BytesIO(raw)):
        raise PackageError("%s：不是有效的 ZIP/OOXML 文件" % label)
    try:
        zf = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile as exc:
        raise PackageError("%s：ZIP 损坏：%s" % (label, exc)) from exc
    with zf:
        infos = zf.infolist()
        if len(infos) > MAX_MEMBERS:
            raise PackageError(
                "%s：ZIP 成员过多（%d > %d）" % (label, len(infos), MAX_MEMBERS)
            )
        total = 0
        for info in infos:
            if info.flag_bits & 0x1:
                raise PackageError(
                    "%s：ZIP 成员已加密，无法读取：%s" % (label, info.filename)
                )
            total += info.file_size
            if info.file_size > MAX_PART_BYTES:
                raise PackageError(
                    "%s：成员过大：%s（%d 字节）"
                    % (label, info.filename, info.file_size)
                )
        if total > MAX_TOTAL_UNCOMPRESSED:
            raise PackageError(
                "%s：解压总大小过大（%d 字节）" % (label, total)
            )
        members = {}
        actual_total = 0
        for info in infos:
            if info.filename.endswith("/"):
                members[info.filename] = b""
                continue
            chunks = []
            got = 0
            with zf.open(info) as src:
                while True:
                    chunk = src.read(1024 * 1024)
                    if not chunk:
                        break
                    got += len(chunk)
                    actual_total += len(chunk)
                    if got > MAX_PART_BYTES:
                        raise PackageError(
                            "%s：成员解压后超过限额：%s"
                            % (label, info.filename)
                        )
                    if actual_total > MAX_TOTAL_UNCOMPRESSED:
                        raise PackageError(
                            "%s：解压总大小超过限额" % label
                        )
                    chunks.append(chunk)
            members[info.filename] = b"".join(chunks)
        comment = zf.comment
    issues = _zip_issues(None, infos)
    return Package(label, raw, members, infos, comment, issues)


def load_package(path):
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise PackageError("文件不存在或不是普通文件：%s" % path)
    size = os.path.getsize(path)
    if size > MAX_ZIP_BYTES:
        raise PackageError("ZIP 过大（%d 字节）：%s" % (size, path))
    with open(path, "rb") as fh:
        raw = fh.read()
    return package_from_bytes(raw, os.path.basename(path))


def zip_bytes_from_package(pkg, replacements, additions):
    """按原顺序重写 ZIP；未替换成员使用原内容字节，元数据尽量保留。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zout:
        zout.comment = pkg.comment
        for info in pkg.infos:
            name = info.filename
            data = replacements.get(name)
            if data is None:
                data = pkg.members[name]
            zout.writestr(info, data)
        now = datetime.now().timetuple()[:6]
        for name, data in additions.items():
            info = zipfile.ZipInfo(name, date_time=now)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (0o644 & 0xFFFF) << 16
            zout.writestr(info, data)
    return buf.getvalue()


# --------------------------------------------------------------------------
# 包级安全分析
# --------------------------------------------------------------------------


def analyze_package(pkg, doc):
    """返回 {"blockers": [...], "warnings": [...], ...}。doc 可为你 None。

    一切保护结构（宏、数字签名、编辑/写保护、现代批注扩展、异常包结构）
    均按 fail-closed 处理：只报告，不修改、不移除、不绕过。
    """
    blockers = list(pkg.issues)
    warnings = [
        "隐藏文本检测只覆盖 run 的 w:rPr 直接属性（vanish / specVanish / "
        "webHidden）；通过 styles.xml、docDefaults 或段落/表格样式继承的隐藏"
        "文本未解析，属已知边界"
    ]

    for name in pkg.names:
        low = name.lower()
        if low == "word/vbaproject.bin" or low.startswith("word/vbadata"):
            blockers.append("包含 VBA 宏：%s" % name)
        if "vbaproject" in low:
            blockers.append("包含 VBA 工程部件：%s" % name)
        if low.startswith("_xmlsignatures/") or low.endswith(".sigs"):
            blockers.append("包含数字签名部件：%s" % name)
        if name == "encryptedPackage":
            blockers.append("包含加密包部件")

    # 内容类型层面的宏 / 签名 / 保护声明
    if pkg.has("[Content_Types].xml"):
        try:
            ct_doc = parse_xml(pkg.get("[Content_Types].xml"), "[Content_Types].xml")
            for el in element_children(ct_doc.documentElement):
                ln = local_name(el)
                if ln not in ("Override", "Default"):
                    continue
                ctype = (el.getAttribute("ContentType") or "").lower()
                part = el.getAttribute("PartName") or el.getAttribute("Extension") or ""
                if "macroenabled" in ctype or "vbaproject" in ctype:
                    blockers.append(
                        "内容类型声明了宏（%s）：%s" % (part, ctype)
                    )
                if "signature" in ctype or "digital-signature" in ctype:
                    blockers.append(
                        "内容类型声明了数字签名（%s）：%s" % (part, ctype)
                    )
        except PackageError as exc:
            blockers.append("无法读取 [Content_Types].xml：%s" % exc)

    present_modern = [n for n in MODERN_COMMENT_PARTS if n in pkg.members]
    if present_modern:
        blockers.append(
            "存在现代批注扩展部件（%s），拒绝追加传统批注"
            % "、".join(present_modern)
        )

    root_ns = None
    if doc is not None:
        root = doc.documentElement
        try:
            _, root_ns = resolve_w_prefix(root)
        except PackageError as exc:
            blockers.append(str(exc))
        if root_ns == STRICT_W_NS:
            blockers.append("不支持 Strict OOXML 命名空间")

    document_protection = None
    write_protection = None
    if pkg.has("word/settings.xml"):
        try:
            settings = parse_xml(pkg.get("word/settings.xml"), "word/settings.xml")
            for el in iter_elements_in_order(settings.documentElement):
                ln = local_name(el)
                if ln == "documentProtection":
                    document_protection = {
                        "enforcement": attr_local(el, "enforcement"),
                        "edit": attr_local(el, "edit"),
                    }
                elif ln == "writeProtection":
                    write_protection = {
                        "recommended": attr_local(el, "recommended"),
                    }
        except PackageError as exc:
            blockers.append("无法读取 word/settings.xml：%s" % exc)
    if document_protection is not None:
        blockers.append(
            "文档存在 documentProtection 保护结构（enforcement=%s, edit=%s），"
            "不自动修改或移除保护"
            % (
                document_protection.get("enforcement") or "未声明",
                document_protection.get("edit") or "未声明",
            )
        )
    if write_protection is not None:
        blockers.append(
            "文档存在 writeProtection 保护结构（recommended=%s），"
            "不自动修改或移除保护"
            % (write_protection.get("recommended") or "未声明")
        )

    return {
        "blockers": blockers,
        "warnings": warnings,
        "namespace": root_ns,
        "modern_comment_parts": present_modern,
        "document_protection": document_protection,
        "write_protection": write_protection,
    }


# --------------------------------------------------------------------------
# 段落模型
# --------------------------------------------------------------------------


class ParagraphSpec:
    __slots__ = ("node", "ancestors", "pid", "text", "blockers", "location")

    def __init__(self, node, ancestors, pid, text, blockers, location):
        self.node = node
        self.ancestors = ancestors
        self.pid = pid
        self.text = text
        self.blockers = blockers
        self.location = location

    @property
    def editable(self):
        return not self.blockers


def _simple_run_children_ok(run):
    """run 只允许 rPr（至多一个，且在最前）与 w:t。"""
    reasons = []
    seen_non_rpr = False
    rpr_count = 0
    for ch in element_children(run):
        ln = local_name(ch)
        if ln == "rPr":
            rpr_count += 1
            if rpr_count > 1:
                reasons.append("run_multiple_rPr")
            if seen_non_rpr:
                reasons.append("run_rPr_not_first")
            for desc in iter_elements_in_order(ch):
                if local_name(desc) in (
                    "rPrChange", "ins", "del", "moveFrom", "moveTo",
                ):
                    reasons.append("existing_revision:%s" % local_name(desc))
            # 直接写在 rPr 里的隐藏文本开关：隐藏条款不能作为普通正文改写
            for prop in element_children(ch):
                pln = local_name(prop)
                if pln in HIDDEN_TOGGLES and _on_off_enabled(prop):
                    reasons.append("hidden_text:%s" % pln)
        elif ln == "t":
            seen_non_rpr = True
        else:
            reasons.append("run_child:%s" % ln)
    return reasons


def paragraph_blockers(node, ancestors):
    reasons = []
    tbl_depth = 0
    for ln, anc_node in ancestors:
        if ln in BLOCKED_ANCESTORS:
            reasons.append("ancestor:%s" % ln)
        if ln == "tbl":
            tbl_depth += 1
        # 表格行/单元格/表格属性级修订或合并，会与单元格文字修订相交
        if ln in ("tr", "tc", "tbl"):
            for prop in element_children(anc_node):
                pl = local_name(prop)
                if pl not in ("trPr", "tcPr", "tblPr", "tblPrEx"):
                    continue
                for desc in iter_elements_in_order(prop):
                    dl = local_name(desc)
                    if dl in REV_TAGS or dl in (
                        "trPrChange", "tcPrChange", "tblPrChange",
                        "tblPrExChange", "cellIns", "cellDel", "cellMerge",
                    ):
                        reasons.append("ancestor_revision:%s" % dl)
    if tbl_depth > 1:
        reasons.append("nested_table")

    ppr_count = 0
    for ch in element_children(node):
        ln = local_name(ch)
        if ln == "pPr":
            ppr_count += 1
            if ppr_count > 1:
                reasons.append("paragraph_multiple_pPr")
            for desc in iter_elements_in_order(ch):
                dln = local_name(desc)
                if dln in ("pPrChange", "ins", "del", "moveFrom", "moveTo"):
                    reasons.append("existing_revision:%s" % dln)
                if dln in COMMENT_ANCHOR_TAGS:
                    reasons.append("existing_comment_anchor:%s" % dln)
        elif ln == "r":
            reasons.extend(_simple_run_children_ok(ch))
        else:
            reasons.append("paragraph_child:%s" % ln)
    return sorted(set(reasons))


def simple_paragraph_text(node):
    parts = []
    for ch in element_children(node):
        if local_name(ch) != "r":
            continue
        for rc in element_children(ch):
            if local_name(rc) == "t":
                parts.append(text_of(rc))
    return "".join(parts)


def render_visible_text(node):
    """近似当前可见文本（accept 视图：包含 w:ins，排除 w:del 内容）。

    仅用于 inspect 展示；可编辑段落上等价于 run 文本拼接。
    """
    out = []

    def walk(cur):
        for ch in element_children(cur):
            ln = local_name(ch)
            if ln in ("del", "moveFrom"):
                continue
            if ln in ("drawing", "pict", "object", "txbxContent"):
                # 文本框等内部段落会被单独编号，避免重复计入外层
                continue
            if ln == "t":
                out.append(text_of(ch))
            elif ln == "tab":
                out.append("\t")
            elif ln in ("br", "cr"):
                out.append("\n")
            elif ln == "noBreakHyphen":
                out.append("-")
            elif ln == "softHyphen":
                continue
            elif ln in ("instrText", "delText"):
                continue
            else:
                walk(ch)

    walk(node)
    return "".join(out)


def find_body(doc):
    for ch in element_children(doc.documentElement):
        if local_name(ch) == "body":
            return ch
    raise PackageError("word/document.xml 缺少 w:body")


def collect_paragraphs(doc, w_prefix):
    body = find_body(doc)
    paragraphs = []
    body_ancestors = [("body", body)]
    stack = []
    for ch in reversed(element_children(body)):
        stack.append((ch, body_ancestors))
    while stack:
        node, ancestors = stack.pop()
        ln = local_name(node)
        if ln == "p" and w_prefix_ok(node, w_prefix):
            text = simple_paragraph_text(node)
            blockers = paragraph_blockers(node, ancestors)
            loc = "table" if any(a[0] == "tbl" for a in ancestors) else "body"
            paragraphs.append(
                ParagraphSpec(
                    node, ancestors, "RL%04d" % (len(paragraphs) + 1), text,
                    blockers, loc,
                )
            )
        child_ancestors = ancestors + [(ln, node)]
        for ch in reversed(element_children(node)):
            stack.append((ch, child_ancestors))
    return paragraphs


# --------------------------------------------------------------------------
# 既有修订 / 批注概览与 ID 空间
# --------------------------------------------------------------------------


def story_parts(pkg):
    return [n for n in pkg.names if STORY_PART_RE.match(n)]


def collect_existing(pkg, strict=True):
    revisions = []
    anchors = []
    comments = []
    errors = []
    for name in story_parts(pkg):
        try:
            doc = parse_xml(pkg.get(name), name)
        except PackageError as exc:
            if strict:
                raise
            errors.append(str(exc))
            continue
        try:
            part_prefix, _ = resolve_w_prefix(doc.documentElement)
        except PackageError:
            part_prefix = None
        for el in iter_elements_in_order(doc.documentElement):
            if part_prefix is not None and not w_prefix_ok(el, part_prefix):
                continue
            ln = local_name(el)
            if ln in REV_TAGS:
                revisions.append(
                    {
                        "part": name,
                        "tag": ln,
                        "id": attr_local(el, "id"),
                        "author": attr_local(el, "author"),
                        "date": attr_local(el, "date"),
                    }
                )
            if ln in COMMENT_ANCHOR_TAGS:
                anchors.append(
                    {"part": name, "tag": ln, "id": attr_local(el, "id")}
                )
            if ln == "comment" and name == "word/comments.xml":
                comments.append(
                    {
                        "part": name,
                        "id": attr_local(el, "id"),
                        "author": attr_local(el, "author"),
                        "date": attr_local(el, "date"),
                        "text": "".join(
                            text_of(t)
                            for t in iter_elements_in_order(el)
                            if local_name(t) == "t"
                        ),
                    }
                )
    return {
        "revisions": revisions,
        "anchors": anchors,
        "comments": comments,
        "errors": errors,
    }


def _max_int(values):
    best = 0
    for v in values:
        try:
            n = int(str(v))
        except (TypeError, ValueError):
            continue
        if n > best:
            best = n
    return best


def scan_existing_ids(pkg, strict=True):
    existing = collect_existing(pkg, strict=strict)
    max_rev = _max_int(r.get("id") for r in existing["revisions"])
    comment_ids = [c.get("id") for c in existing["comments"]]
    comment_ids += [a.get("id") for a in existing["anchors"]]
    max_comment = _max_int(comment_ids)
    return {
        "max_revision_id": max_rev,
        "max_comment_id": max_comment,
        "existing": existing,
    }


# --------------------------------------------------------------------------
# 文档上下文
# --------------------------------------------------------------------------


class DocContext:
    def __init__(self, pkg):
        self.pkg = pkg
        self.sha256 = pkg.sha256
        if not pkg.has("word/document.xml"):
            raise PackageError("缺少 word/document.xml")
        self.doc_xml = pkg.get("word/document.xml")
        self.doc = parse_xml(self.doc_xml, "word/document.xml")
        root = self.doc.documentElement
        if local_name(root) != "document":
            raise PackageError("word/document.xml 根元素不是 w:document")
        self.w_prefix, self.namespace = resolve_w_prefix(root)
        self.safety = analyze_package(pkg, self.doc)
        self.paragraphs = collect_paragraphs(self.doc, self.w_prefix)
        self.by_id = {p.pid: p for p in self.paragraphs}


def load_context(path, for_write=False):
    pkg = load_package(path)
    if for_write and pkg.issues:
        raise PackageError(
            "包结构异常，拒绝写入：%s" % "；".join(pkg.issues)
        )
    ctx = DocContext(pkg)
    if for_write and ctx.safety["blockers"]:
        raise PackageError(
            "该文档不允许写入：%s" % "；".join(ctx.safety["blockers"])
        )
    return ctx


# --------------------------------------------------------------------------
# 编辑校验与粒度
# --------------------------------------------------------------------------


EDIT_KEYS = {"id", "para", "old", "new", "reason", "comment", "broad_reason"}


def normalize_edits(edits):
    if not isinstance(edits, list) or not edits:
        raise InputError("edits 必须是非空数组")
    normalized = []
    seen_ids = set()
    for idx, raw in enumerate(edits):
        if not isinstance(raw, dict):
            raise InputError("edits[%d] 必须是对象" % idx)
        unknown = set(raw) - EDIT_KEYS
        if unknown:
            raise InputError(
                "edits[%d] 含未知字段：%s" % (idx, "、".join(sorted(unknown)))
            )
        eid = raw.get("id")
        if not isinstance(eid, str) or not eid.strip():
            raise InputError("edits[%d].id 必须是非空字符串" % idx)
        eid = eid.strip()
        if eid in seen_ids:
            raise InputError("编辑 id 重复：%s" % eid)
        seen_ids.add(eid)
        para = raw.get("para")
        if not isinstance(para, str) or not PARA_ID_RE.match(para):
            raise InputError(
                "edits[%d].para 必须形如 RL0001（本引擎独立编号）" % idx
            )
        old = raw.get("old")
        if not isinstance(old, str) or old == "":
            raise InputError("edits[%d].old 必须是非空精确子串（空锚拒绝）" % idx)
        has_new = "new" in raw
        has_comment = "comment" in raw
        if has_new and has_comment:
            raise InputError(
                "edits[%d] 不能同时给 new 与 comment（替换用 reason，仅批注用 comment）"
                % idx
            )
        if not has_new and not has_comment:
            raise InputError("edits[%d] 必须提供 new 或 comment" % idx)
        item = {"id": eid, "para": para, "old": old}
        if has_new:
            new = raw.get("new")
            if not isinstance(new, str):
                raise InputError("edits[%d].new 必须是字符串" % idx)
            if new == old:
                raise InputError("edits[%d] old 与 new 相同，属于空修改" % idx)
            reason = raw.get("reason")
            if not isinstance(reason, str) or not reason.strip():
                raise InputError(
                    "edits[%d].reason 必须是非空批注理由（替换会产生范围批注）" % idx
                )
            item["new"] = new
            item["reason"] = reason.strip()
            if "broad_reason" in raw:
                br = raw.get("broad_reason")
                if not isinstance(br, str) or not br.strip():
                    raise InputError(
                        "edits[%d].broad_reason 必须是非空字符串" % idx
                    )
                item["broad_reason"] = br.strip()
        else:
            comment = raw.get("comment")
            if not isinstance(comment, str) or not comment.strip():
                raise InputError("edits[%d].comment 必须是非空字符串" % idx)
            item["comment"] = comment.strip()
            if "reason" in raw:
                raise InputError(
                    "edits[%d] 仅批注请用 comment，不要用 reason" % idx
                )
            if "broad_reason" in raw:
                raise InputError(
                    "edits[%d] 仅批注不需要 broad_reason" % idx
                )
        normalized.append(item)
    return normalized


def find_occurrences(text, needle, limit=2):
    """返回 needle 在 text 中的（允许重叠的）起始位置，最多 limit 个。"""
    positions = []
    start = 0
    while len(positions) < limit:
        pos = text.find(needle, start)
        if pos < 0:
            break
        positions.append(pos)
        start = pos + 1
    return positions


def tokenize(text):
    """数字（含千分位/小数/百分比）作为不可分 token，其余按字符切分。"""
    tokens = []
    i = 0
    n = len(text)
    while i < n:
        m = NUMBER_RE.match(text, i)
        if m and m.end() > i:
            tokens.append((m.group(0), i, m.end()))
            i = m.end()
        else:
            tokens.append((text[i], i, i + 1))
            i += 1
    return tokens


def compute_diff_blocks(old, new):
    """old/new 的 token 级 diff；相邻变化间隔 <=2 字符时合并。"""
    a = tokenize(old)
    b = tokenize(new)
    sm = SequenceMatcher(
        None, [t[0] for t in a], [t[0] for t in b], autojunk=False
    )
    raw = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        if i1 < len(a):
            s = a[i1][1]
        else:
            s = len(old)
        e = a[i2 - 1][2] if i2 > i1 else s
        target = "".join(t[0] for t in b[j1:j2])
        raw.append([s, e, target])
    merged = []
    for s, e, target in raw:
        if merged and s - merged[-1][1] <= 2:
            prev = merged[-1]
            # 把两块之间的未变文字补回目标串，避免合并丢字
            prev[2] = prev[2] + old[prev[1]:s] + target
            prev[1] = max(prev[1], e)
        else:
            merged.append([s, e, target])
    # 校验目标串可重建
    rebuilt = []
    cursor = 0
    for s, e, target in merged:
        if s < cursor or e < s:
            raise InputError("内部 diff 区间异常")
        rebuilt.append(old[cursor:s])
        rebuilt.append(target)
        cursor = e
    rebuilt.append(old[cursor:])
    if "".join(rebuilt) != new:
        # 保守回退：整段作为一个替换块
        merged = [[0, len(old), new]]
    return merged


def granularity_report(old, new, blocks):
    changed_chars = sum(max(e - s, len(target)) for s, e, target in blocks)
    span_chars = sum(e - s for s, e, target in blocks)
    para_chars = len(old)
    ratio = changed_chars / max(1, para_chars)
    fragment_count = len(blocks)
    warnings = []
    if changed_chars > 60:
        warnings.append("单处改动超过 60 字（%d 字）" % changed_chars)
    if para_chars > 60 and ratio > 0.6:
        warnings.append(
            "长段落改动比例超过 60%%（%.0f%%）" % (ratio * 100)
        )
    if fragment_count > 6:
        warnings.append("改动分散为 %d 块（>6）" % fragment_count)
    return {
        "changed_chars": changed_chars,
        "revised_span_chars": span_chars,
        "paragraph_chars": para_chars,
        "change_ratio": round(ratio, 4),
        "fragment_count": fragment_count,
        "warnings": warnings,
        "requires_broad_reason": bool(warnings),
    }


def build_plan(ctx, edits, created_at=None):
    normalized = normalize_edits(edits)
    ops = []
    warnings = []
    seen_para = set()
    for edit in normalized:
        para = ctx.by_id.get(edit["para"])
        if para is None:
            raise InputError(
                "段落 %s 不存在（请先运行 inspect 获取 RL 编号）" % edit["para"]
            )
        if not para.editable:
            raise InputError(
                "段落 %s 不可编辑：%s"
                % (edit["para"], "、".join(para.blockers))
            )
        if para.pid in seen_para:
            raise InputError(
                "同一段落只允许一个 op，%s 出现多次" % para.pid
            )
        seen_para.add(para.pid)
        text = para.text
        old = edit["old"]
        positions = find_occurrences(text, old)
        if not positions:
            raise InputError(
                "段落 %s 中找不到 old 精确子串（跨段引用或文本不符）：%r"
                % (edit["para"], old[:80])
            )
        if len(positions) > 1:
            raise InputError(
                "段落 %s 中 old 命中多次，必须唯一：%r"
                % (edit["para"], old[:80])
            )
        start = positions[0]
        end = start + len(old)

        if "comment" in edit:
            ops.append(
                {
                    "id": edit["id"],
                    "para": edit["para"],
                    "kind": "comment",
                    "old": old,
                    "comment": edit["comment"],
                    "paragraph_text": text,
                    "target_text": text,
                    "diff": [],
                    "granularity": {
                        "changed_chars": 0,
                        "revised_span_chars": 0,
                        "paragraph_chars": len(text),
                        "change_ratio": 0.0,
                        "fragment_count": 0,
                        "warnings": [],
                        "requires_broad_reason": False,
                    },
                }
            )
            continue

        new = edit["new"]
        blocks = compute_diff_blocks(old, new)
        for block in blocks:
            block[0] += start
            block[1] += start
        target_text = text[:start] + new + text[end:]
        diff = [
            {
                "src_start": s,
                "src_end": e,
                "old": text[s:e],
                "new": target,
            }
            for s, e, target in blocks
        ]
        gran = granularity_report(text, target_text, blocks)
        if gran["requires_broad_reason"] and not edit.get("broad_reason"):
            raise InputError(
                "段落 %s 属大幅改写（%s），op %s 必须提供非空 broad_reason 以记录用户确认"
                % (
                    edit["para"],
                    "；".join(gran["warnings"]),
                    edit["id"],
                )
            )
        op = {
            "id": edit["id"],
            "para": edit["para"],
            "kind": "replace",
            "old": old,
            "new": new,
            "reason": edit["reason"],
            "paragraph_text": text,
            "target_text": target_text,
            "diff": diff,
            "granularity": gran,
        }
        if edit.get("broad_reason"):
            op["broad_reason"] = edit["broad_reason"]
        ops.append(op)

    if len(ops) > 15:
        warnings.append(
            "本计划将新增 %d 条批注（>15），建议复核批注数量" % len(ops)
        )

    payload = plan_payload(ctx.sha256, normalized, ops, warnings)
    return {
        "tool": TOOL,
        "version": PLAN_VERSION,
        "created_at": created_at or utc_now_iso(),
        "source": {
            "file": os.path.basename(ctx.pkg.label),
            "sha256": ctx.sha256,
            "size": len(ctx.pkg.raw),
        },
        "source_sha256": ctx.sha256,
        "edits": normalized,
        "ops": ops,
        "warnings": warnings,
        "comment_count": len(ops),
        "fingerprint": {
            "algorithm": "sha256",
            "canonical": "json-sort-keys-utf8",
            "value": fingerprint_value(payload),
        },
    }


def plan_payload(source_sha256, edits, ops, warnings):
    return {
        "version": PLAN_VERSION,
        "source_sha256": source_sha256,
        "edits": edits,
        "ops": ops,
        "warnings": warnings,
    }


def fingerprint_value(payload):
    return sha256_bytes(canonical_json(payload).encode("utf-8"))


def validate_plan(ctx, plan):
    """重新规划并与计划比对；返回重建后的计划。不信任 plan 中的定位/XML。"""
    if not isinstance(plan, dict):
        raise InputError("plan 必须是 JSON 对象")
    if plan.get("tool") != TOOL:
        raise InputError("plan.tool 不是 %s" % TOOL)
    if plan.get("version") != PLAN_VERSION:
        raise InputError("不支持的 plan.version：%r" % plan.get("version"))
    if plan.get("source_sha256") != ctx.sha256:
        raise InputError(
            "源文件已变化或计划已过期：plan=%s 实际=%s"
            % (plan.get("source_sha256"), ctx.sha256)
        )
    fp = plan.get("fingerprint") or {}
    if fp.get("algorithm") != "sha256":
        raise InputError("plan.fingerprint 缺失或算法不符")
    edits = plan.get("edits")
    if not isinstance(edits, list):
        raise InputError("plan.edits 缺失")
    rebuilt = build_plan(ctx, edits, created_at=plan.get("created_at"))
    if (
        rebuilt["edits"] != plan.get("edits")
        or rebuilt["ops"] != plan.get("ops")
        or rebuilt["warnings"] != plan.get("warnings")
    ):
        raise InputError("计划内容与源文件重新规划结果不一致（可能被篡改）")
    payload = plan_payload(
        rebuilt["source_sha256"],
        rebuilt["edits"],
        rebuilt["ops"],
        rebuilt["warnings"],
    )
    expected = fingerprint_value(payload)
    if fp.get("value") != expected or rebuilt["fingerprint"]["value"] != expected:
        raise InputError("计划指纹不匹配（可能被篡改或过期）")
    return rebuilt


# --------------------------------------------------------------------------
# 段落重写
# --------------------------------------------------------------------------


class RunSpec:
    __slots__ = ("node", "rpr", "text", "start", "end")

    def __init__(self, node, rpr, text, start, end):
        self.node = node
        self.rpr = rpr
        self.text = text
        self.start = start
        self.end = end


def _rpr_of(run):
    for ch in element_children(run):
        if local_name(ch) == "rPr":
            return ch
    return None


def simple_runs(paragraph):
    runs = []
    cursor = 0
    for ch in element_children(paragraph):
        if local_name(ch) != "r":
            continue
        parts = []
        for rc in element_children(ch):
            if local_name(rc) == "t":
                parts.append(text_of(rc))
        text = "".join(parts)
        runs.append(RunSpec(ch, _rpr_of(ch), text, cursor, cursor + len(text)))
        cursor += len(text)
    return runs


class Emitted:
    __slots__ = ("node", "kind", "src")

    def __init__(self, node, kind, src):
        self.node = node
        self.kind = kind  # equal | changed
        self.src = src


def make_run(doc, prefix, rpr, text, text_tag="t"):
    run = doc.createElement(wq(prefix, "r"))
    if rpr is not None:
        run.appendChild(rpr.cloneNode(True))
    t = doc.createElement(wq(prefix, text_tag))
    t.setAttribute("xml:space", "preserve")
    t.appendChild(doc.createTextNode(text))
    run.appendChild(t)
    return run


def make_revision(doc, prefix, tag, rid, author, date, rpr, text):
    el = doc.createElement(wq(prefix, tag))
    set_wattr(el, prefix, "id", str(rid))
    set_wattr(el, prefix, "author", author)
    set_wattr(el, prefix, "date", date)
    el.appendChild(
        make_run(
            doc,
            prefix,
            rpr,
            text,
            "delText" if tag == "del" else "t",
        )
    )
    return el


def make_comment_reference(doc, prefix, cid):
    run = doc.createElement(wq(prefix, "r"))
    rpr = doc.createElement(wq(prefix, "rPr"))
    style = doc.createElement(wq(prefix, "rStyle"))
    set_wattr(style, prefix, "val", "CommentReference")
    rpr.appendChild(style)
    run.appendChild(rpr)
    ref = doc.createElement(wq(prefix, "commentReference"))
    set_wattr(ref, prefix, "id", str(cid))
    run.appendChild(ref)
    return run


def make_comment_range(doc, prefix, tag, cid):
    el = doc.createElement(wq(prefix, tag))
    set_wattr(el, prefix, "id", str(cid))
    return el


def _style_at(runs, pos):
    for run in runs:
        if run.start <= pos < run.end:
            return run.rpr
    prev = None
    for run in runs:
        if run.end <= pos and run.start < run.end:
            prev = run
    if prev is not None:
        return prev.rpr
    for run in runs:
        if run.start < run.end:
            return run.rpr
    return None


def rewrite_paragraph(
    doc, prefix, paragraph, op, author, date, rev_id, comment_id
):
    """就地重写一个可编辑段落；返回 (next_rev_id, next_comment_id, info)。"""
    runs = simple_runs(paragraph)
    text = "".join(r.text for r in runs)
    if text != op["paragraph_text"]:
        raise InputError(
            "段落 %s 文本与计划不一致（源文件可能已变化）" % op["para"]
        )
    old = op["old"]
    positions = find_occurrences(text, old)
    if len(positions) != 1:
        raise InputError("段落 %s 的 old 不再唯一" % op["para"])
    start = positions[0]
    end = start + len(old)

    if op["kind"] == "comment":
        blocks = []
        splits = {start, end}
    else:
        blocks = [
            (d["src_start"], d["src_end"], d["new"]) for d in op["diff"]
        ]
        splits = set()

    empty_by_offset = {}
    for idx, run in enumerate(runs):
        if run.start == run.end:
            empty_by_offset.setdefault(run.start, []).append(idx)
    emitted_empty = set()

    emitted = []

    def emit_empty_index(idx):
        if idx in emitted_empty:
            return
        emitted_empty.add(idx)
        run = runs[idx]
        emitted.append(
            Emitted(
                run.node.cloneNode(True), "equal", (run.start, run.end)
            )
        )

    def flush_empty(pos):
        for idx in empty_by_offset.get(pos, []):
            emit_empty_index(idx)

    def emit_equal(a, b):
        for run_idx, run in enumerate(runs):
            if run.start == run.end:
                if a < run.start < b:
                    emit_empty_index(run_idx)
                continue
            lo = max(a, run.start)
            hi = min(b, run.end)
            if lo >= hi:
                continue
            cuts = [lo]
            cuts.extend(sorted(o for o in splits if lo < o < hi))
            cuts.append(hi)
            for x, y in zip(cuts, cuts[1:]):
                if x >= y:
                    continue
                if x == run.start and y == run.end:
                    node = run.node.cloneNode(True)
                else:
                    node = make_run(doc, prefix, run.rpr, text[x:y])
                emitted.append(Emitted(node, "equal", (x, y)))

    def emit_deletion(s, e):
        nonlocal rev_id
        for run_idx, run in enumerate(runs):
            if run.start == run.end:
                if s < run.start < e:
                    emit_empty_index(run_idx)
                continue
            lo = max(s, run.start)
            hi = min(e, run.end)
            if lo < hi:
                emitted.append(
                    Emitted(
                        make_revision(
                            doc, prefix, "del", rev_id, author, date,
                            run.rpr, text[lo:hi],
                        ),
                        "changed",
                        (lo, hi),
                    )
                )
                rev_id += 1

    def emit_insertion(at_pos, target):
        nonlocal rev_id
        if not target:
            return
        rpr = _style_at(runs, at_pos)
        emitted.append(
            Emitted(
                make_revision(
                    doc, prefix, "ins", rev_id, author, date, rpr, target
                ),
                "changed",
                (at_pos, at_pos),
            )
        )
        rev_id += 1

    cursor = 0
    flush_empty(0)
    for s, e, target in blocks:
        if s > cursor:
            emit_equal(cursor, s)
        flush_empty(s)
        if e > s:
            emit_deletion(s, e)
        emit_insertion(s, target)
        cursor = max(cursor, e)
        flush_empty(cursor)
    if cursor < len(text):
        emit_equal(cursor, len(text))
    flush_empty(len(text))

    if not emitted:
        raise InputError("段落 %s 未产生任何输出" % op["para"])

    if op["kind"] == "comment":
        indexes = [
            i
            for i, em in enumerate(emitted)
            if em.kind == "equal" and em.src[1] > start and em.src[0] < end
        ]
    else:
        indexes = [i for i, em in enumerate(emitted) if em.kind == "changed"]
    if not indexes:
        raise InputError("段落 %s 未定位到批注锚区域" % op["para"])
    first, last = min(indexes), max(indexes)

    cstart = make_comment_range(doc, prefix, "commentRangeStart", comment_id)
    cend = make_comment_range(doc, prefix, "commentRangeEnd", comment_id)
    cref = make_comment_reference(doc, prefix, comment_id)
    comment_id += 1

    final = []
    for i, em in enumerate(emitted):
        if i == first:
            final.append(cstart)
        final.append(em.node)
        if i == last:
            final.append(cend)
            final.append(cref)

    ppr = None
    for ch in element_children(paragraph):
        if local_name(ch) == "pPr":
            ppr = ch
            break
    for ch in list(paragraph.childNodes):
        if ch is ppr:
            continue
        paragraph.removeChild(ch)
    for node in final:
        paragraph.appendChild(node)

    return rev_id, comment_id, {
        "para": op["para"],
        "kind": op["kind"],
    }


# --------------------------------------------------------------------------
# comments.xml / rels / content types
# --------------------------------------------------------------------------


def comment_xml_text(prefix, cid, author, date, body):
    q = lambda local: wq(prefix, local)  # noqa: E731
    lines = body.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines:
        lines = [""]
    run_parts = []
    for idx, line in enumerate(lines):
        if idx:
            run_parts.append("<%s/>" % q("br"))
        run_parts.append(
            '<%s xml:space="preserve">%s</%s>'
            % (q("t"), xml_escape(line), q("t"))
        )
    val = wq(prefix, "val")
    body_xml = (
        '<%s><%s><%s %s="CommentText"/></%s>'
        '<%s><%s><%s %s="CommentReference"/></%s><%s/></%s>'
        '<%s>%s</%s></%s>'
        % (
            q("p"), q("pPr"), q("pStyle"), val, q("pPr"),
            q("r"), q("rPr"), q("rStyle"), val, q("rPr"),
            q("annotationRef"), q("r"),
            q("r"), "".join(run_parts), q("r"), q("p"),
        )
    )
    return (
        '<%s %s=%s %s=%s %s=%s>%s</%s>'
        % (
            q("comment"),
            wq(prefix, "id"), quoteattr(str(cid)),
            wq(prefix, "author"), quoteattr(author),
            wq(prefix, "date"), quoteattr(date),
            body_xml,
            q("comment"),
        )
    )


def build_comments_part(pkg, ctx, comments_data, author, date):
    """返回 (new_comments_bytes, is_new)。existing 内容按原始字节保留。

    comments_data 为 [(comment_id, body_text), ...]；新批注 XML 使用目标部件
    自身的命名空间前缀，避免 document.xml 与 comments.xml 前缀不同导致非法 XML。
    """
    if pkg.has("word/comments.xml"):
        original = pkg.get("word/comments.xml")
        doc = parse_xml(original, "word/comments.xml")
        root = doc.documentElement
        if local_name(root) != "comments":
            raise PackageError("word/comments.xml 根元素不是 w:comments")
        prefix, _ = resolve_w_prefix(root)
        if not prefix:
            raise PackageError("word/comments.xml 使用默认命名空间，拒绝追加")
        new_comments = [
            comment_xml_text(prefix, cid, author, date, body)
            for cid, body in comments_data
        ]
        enc_match = re.search(
            rb"encoding\s*=\s*[\"']([^\"']+)", bytes(original[:300]), re.I
        )
        if enc_match and enc_match.group(1).lower() not in (b"utf-8", b"utf8"):
            raise PackageError(
                "word/comments.xml 声明编码为 %s，拒绝在原地追加"
                % enc_match.group(1).decode("ascii", "ignore")
            )
        try:
            text = original.decode("utf-8")
        except UnicodeDecodeError as exc:
            raise PackageError(
                "word/comments.xml 不是 UTF-8，拒绝原地追加：%s" % exc
            ) from exc
        close = "</%s:comments>" % prefix
        idx = text.rfind(close)
        if idx >= 0:
            new_text = text[:idx] + "".join(new_comments) + text[idx:]
        else:
            m = re.search(
                r"<%s:comments\b[^>]*/>" % re.escape(prefix), text
            )
            if not m:
                raise PackageError("无法定位 word/comments.xml 根元素闭合位置")
            tag = m.group(0)
            open_tag = tag[:-2].rstrip() + ">"
            new_text = (
                text[: m.start()]
                + open_tag
                + "".join(new_comments)
                + close
                + text[m.end():]
            )
        new_bytes = new_text.encode("utf-8")
        # 必须仍是合法 XML
        parse_xml(new_bytes, "word/comments.xml")
        return new_bytes, False

    prefix = ctx.w_prefix or "w"
    if not prefix:
        prefix = "w"
    body = "".join(
        comment_xml_text(prefix, cid, author, date, text)
        for cid, text in comments_data
    )
    text = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<%s xmlns:%s="%s">%s</%s>'
        % (wq(prefix, "comments"), prefix, W_NS, body, wq(prefix, "comments"))
    )
    data = text.encode("utf-8")
    parse_xml(data, "word/comments.xml")
    return data, True


def ensure_content_type(pkg):
    name = "[Content_Types].xml"
    if not pkg.has(name):
        raise PackageError("缺少 [Content_Types].xml")
    original = pkg.get(name)
    doc = parse_xml(original, name)
    root = doc.documentElement
    if root is None or local_name(root) != "Types":
        raise PackageError("[Content_Types].xml 根元素不是 Types")
    prefix = prefix_for_ns(root, CT_NS)
    if prefix is None:
        raise PackageError("[Content_Types].xml 缺少内容类型命名空间声明")
    found = False
    for el in element_children(root):
        if local_name(el) == "Override":
            part = el.getAttribute("PartName")
            if part.lower() == "/word/comments.xml":
                ctype = el.getAttribute("ContentType")
                if ctype != COMMENTS_CT:
                    raise PackageError(
                        "word/comments.xml 的 ContentType 异常：%s" % ctype
                    )
                found = True
    if found:
        return original
    ov = doc.createElement(wq(prefix, "Override"))
    ov.setAttribute("PartName", "/word/comments.xml")
    ov.setAttribute("ContentType", COMMENTS_CT)
    root.appendChild(ov)
    return serialize_like(original, doc)


def _next_rid(rels_doc):
    used = set()
    for el in element_children(rels_doc.documentElement):
        rid = el.getAttribute("Id")
        if rid:
            used.add(rid)
    n = 1
    while "rId%d" % n in used:
        n += 1
    return "rId%d" % n, used


def ensure_comments_relationship(pkg):
    name = "word/_rels/document.xml.rels"
    if not pkg.has(name):
        raise PackageError("缺少 %s" % name)
    original = pkg.get(name)
    doc = parse_xml(original, name)
    root = doc.documentElement
    if root is None or local_name(root) != "Relationships":
        raise PackageError("%s 根元素不是 Relationships" % name)
    prefix = prefix_for_ns(root, REL_NS)
    if prefix is None:
        raise PackageError("%s 缺少关系命名空间声明" % name)
    existing = None
    for el in element_children(root):
        if local_name(el) != "Relationship":
            continue
        if el.getAttribute("Type") == COMMENTS_REL:
            target = el.getAttribute("Target")
            part = _resolve_rel_target(target)
            if part != "word/comments.xml":
                raise PackageError(
                    "批注关系指向意外的目标：%s" % target
                )
            existing = el
    if existing is not None:
        return original
    rid, _ = _next_rid(doc)
    rel = doc.createElement(wq(prefix, "Relationship"))
    rel.setAttribute("Id", rid)
    rel.setAttribute("Type", COMMENTS_REL)
    rel.setAttribute("Target", "comments.xml")
    root.appendChild(rel)
    return serialize_like(original, doc)


# --------------------------------------------------------------------------
# apply
# --------------------------------------------------------------------------


def validate_author(author):
    if not isinstance(author, str) or not author.strip():
        raise InputError("--author 不能为空")
    value = author.strip()
    if len(value) > 100:
        raise InputError("--author 过长")
    return value


def build_result(ctx, plan, author, timestamp=None):
    """全量内存构建结果；返回 (result_bytes, stats)。不写盘，不调用 verify_core。

    timestamp 仅供 verify 以结果中提取的唯一 author/date 重建预期包；默认取
    当前 UTC 时间。date 是本函数唯一的不确定输入。
    """
    ids = scan_existing_ids(ctx.pkg, strict=True)
    rev_id = ids["max_revision_id"] + 1
    comment_id = ids["max_comment_id"] + 1
    comment_id_start = comment_id
    date = utc_now_iso() if timestamp is None else str(timestamp)
    if not DATE_RE.match(date):
        raise InputError(
            "timestamp 必须形如 2024-01-01T00:00:00Z：%r" % date
        )

    # 先按计划顺序逐段重写（都在内存树上）
    comments_data = []
    stats: dict[str, Any] = {"revisions": 0, "comments": 0}
    for op in plan["ops"]:
        para = ctx.by_id.get(op["para"])
        if para is None:
            raise InputError("段落 %s 不存在" % op["para"])
        if not para.editable:
            raise InputError(
                "段落 %s 不可编辑：%s"
                % (op["para"], "、".join(para.blockers))
            )
        before_rev = rev_id
        this_comment_id = comment_id
        rev_id, comment_id, _ = rewrite_paragraph(
            ctx.doc, ctx.w_prefix, para.node, op, author, date,
            rev_id, comment_id,
        )
        stats["revisions"] += rev_id - before_rev
        stats["comments"] += 1
        body = op.get("reason") if op["kind"] == "replace" else op.get("comment")
        comments_data.append((this_comment_id, body or ""))

    doc_bytes = serialize_like(ctx.doc_xml, ctx.doc)
    comments_bytes, is_new = build_comments_part(
        ctx.pkg, ctx, comments_data, author, date
    )
    ct_bytes = ensure_content_type(ctx.pkg)
    rels_bytes = ensure_comments_relationship(ctx.pkg)

    replacements = {
        "word/document.xml": doc_bytes,
        "word/comments.xml": comments_bytes,
        "[Content_Types].xml": ct_bytes,
        "word/_rels/document.xml.rels": rels_bytes,
    }
    additions = {}
    if is_new:
        additions["word/comments.xml"] = comments_bytes
        replacements.pop("word/comments.xml")

    result_bytes = zip_bytes_from_package(ctx.pkg, replacements, additions)
    stats["new_revision_ids"] = list(
        range(ids["max_revision_id"] + 1, rev_id)
    )
    stats["new_comment_ids"] = list(
        range(comment_id_start, comment_id)
    )
    stats["date"] = date
    stats["author"] = author
    return result_bytes, stats


def publish_new_file(path, data):
    """独占发布：目标必须不存在；同目录临时文件 + os.link 原子创建。"""
    path = os.path.abspath(path)
    parent = os.path.dirname(path) or "."
    if not os.path.isdir(parent):
        raise InputError("输出目录不存在：%s" % parent)
    if os.path.exists(path):
        raise InputError("输出文件已存在，拒绝覆盖：%s" % path)
    fd, tmp = tempfile.mkstemp(prefix=".hetong_redline.", dir=parent)
    try:
        with os.fdopen(fd, "wb") as fh:
            fh.write(data)
            fh.flush()
            os.fsync(fh.fileno())
        try:
            os.link(tmp, path)
        except FileExistsError:
            raise InputError("输出文件已存在，拒绝覆盖：%s" % path) from None
        except OSError as exc:
            raise PackageError(
                "无法原子独占发布输出文件（%s）：%s" % (path, exc)
            ) from exc
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass
    return path


# --------------------------------------------------------------------------
# 视图模拟与验证
# --------------------------------------------------------------------------


def _remove_node(node):
    parent = node.parentNode
    if parent is not None:
        parent.removeChild(node)


def _unwrap(node):
    parent = node.parentNode
    if parent is None:
        return
    for ch in list(node.childNodes):
        parent.insertBefore(ch, node)
    parent.removeChild(node)


def _unwrap_deleted(doc, prefix, node):
    """拒绝删除：把 delText 还原为 t，并把 del 的子 run 提升到父级。"""
    parent = node.parentNode
    if parent is None:
        return
    for ch in list(node.childNodes):
        if not is_element(ch):
            continue
        if local_name(ch) == "r":
            for rc in element_children(ch):
                if local_name(rc) == "delText":
                    new_t = doc.createElement(wq(prefix, "t"))
                    for i in range(rc.attributes.length):
                        a = rc.attributes.item(i)
                        if not isinstance(a, minidom.Attr):
                            raise VerifyError("删除文字属性结构异常")
                        new_t.setAttribute(a.name, a.value)
                    for txt in list(rc.childNodes):
                        new_t.appendChild(txt)
                    ch.replaceChild(new_t, rc)
            parent.insertBefore(ch, node)
    parent.removeChild(node)


def _remove_comment_reference(node):
    run = node.parentNode
    if run is not None and local_name(run) == "r":
        left = [
            ch
            for ch in element_children(run)
            if local_name(ch) not in ("rPr", "commentReference")
        ]
        if not left:
            _remove_node(run)
            return
    _remove_node(node)


def simulate_new_changes(doc, prefix, new_rev_ids, new_comment_ids, mode):
    """仅模拟接受/拒绝本轮新增的修订与批注 ID，原修订保持不动。"""
    new_rev_ids = {str(x) for x in new_rev_ids}
    new_comment_ids = {str(x) for x in new_comment_ids}
    for el in list(iter_elements_in_order(doc.documentElement)):
        if not w_prefix_ok(el, prefix):
            continue
        ln = local_name(el)
        rid = attr_local(el, "id")
        if ln in REV_TAGS and rid in new_rev_ids:
            if mode == "reject":
                if ln == "ins":
                    _remove_node(el)
                elif ln == "del":
                    _unwrap_deleted(doc, prefix, el)
                else:
                    _remove_node(el)
            else:  # accept
                if ln == "ins":
                    _unwrap(el)
                elif ln == "del":
                    _remove_node(el)
                else:
                    _remove_node(el)
        elif ln in COMMENT_ANCHOR_TAGS and rid in new_comment_ids:
            if ln == "commentReference":
                _remove_comment_reference(el)
            else:
                _remove_node(el)


def paragraph_option_signature(node):
    for ch in element_children(node):
        if local_name(ch) == "pPr":
            return ch.toxml()
    return ""


def paragraph_runs_signature(node):
    """(pPr, 合并相邻同样式 run 后的 [(rPr xml, text)])。"""
    runs = []

    def walk(cur):
        for ch in element_children(cur):
            ln = local_name(ch)
            if ln in ("del", "moveFrom"):
                continue
            if ln == "r":
                rpr = ""
                parts = []
                for rc in element_children(ch):
                    rln = local_name(rc)
                    if rln == "rPr":
                        rpr = rc.toxml()
                    elif rln == "t":
                        parts.append(text_of(rc))
                    elif rln == "tab":
                        parts.append("\t")
                    elif rln in ("br", "cr"):
                        parts.append("\n")
                runs.append((rpr, "".join(parts)))
            else:
                walk(ch)

    walk(node)
    merged = []
    for rpr, txt in runs:
        if merged and merged[-1][0] == rpr:
            merged[-1][1] += txt
        else:
            merged.append([rpr, txt])
    return paragraph_option_signature(node), tuple(
        (rpr, txt) for rpr, txt in merged
    )


def document_signature(doc, w_prefix):
    paragraphs = collect_paragraphs(doc, w_prefix)
    return [paragraph_runs_signature(p.node) for p in paragraphs]


def document_view_texts(doc, w_prefix):
    paragraphs = collect_paragraphs(doc, w_prefix)
    return [render_visible_text(p.node) for p in paragraphs]


def _parse_comment_texts(comments_bytes):
    result = {}
    if comments_bytes is None:
        return result
    doc = parse_xml(comments_bytes, "word/comments.xml")
    for el in iter_elements_in_order(doc.documentElement):
        if local_name(el) == "comment":
            cid = attr_local(el, "id")
            text = "".join(
                text_of(t)
                for t in iter_elements_in_order(el)
                if local_name(t) == "t"
            )
            result[str(cid)] = {
                "author": attr_local(el, "author"),
                "date": attr_local(el, "date"),
                "text": text,
            }
    return result


def _resolve_rel_target(target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("word", target))


def _relations_comments_target(rels_bytes):
    doc = parse_xml(rels_bytes, "word/_rels/document.xml.rels")
    targets = []
    for el in element_children(doc.documentElement):
        if (
            local_name(el) == "Relationship"
            and el.getAttribute("Type") == COMMENTS_REL
        ):
            targets.append(_resolve_rel_target(el.getAttribute("Target")))
    return targets


def _has_comments_content_type(ct_bytes):
    doc = parse_xml(ct_bytes, "[Content_Types].xml")
    for el in element_children(doc.documentElement):
        if local_name(el) == "Override":
            if el.getAttribute("PartName").lower() == "/word/comments.xml":
                return el.getAttribute("ContentType") == COMMENTS_CT
    return False


def _check_new_revision_structure(doc, prefix, new_rev_ids, result):
    errors = []
    seen = set()
    for el in iter_elements_in_order(doc.documentElement):
        if not w_prefix_ok(el, prefix):
            continue
        ln = local_name(el)
        if ln not in ("ins", "del"):
            continue
        rid = attr_local(el, "id")
        if str(rid) not in new_rev_ids:
            continue
        seen.add(str(rid))
        author = attr_local(el, "author")
        date = attr_local(el, "date")
        if not author:
            errors.append("新修订 %s 缺少 author" % rid)
        if not date or not DATE_RE.match(date or ""):
            errors.append("新修订 %s 的 date 非法：%r" % (rid, date))
        result["authors"].add(author or "")
        if ln == "ins":
            for ch in element_children(el):
                if local_name(ch) != "r":
                    errors.append("w:ins %s 含非 run 子元素 %s" % (rid, local_name(ch)))
                    continue
                for rc in element_children(ch):
                    rln = local_name(rc)
                    if rln not in ("rPr", "t"):
                        errors.append("w:ins %s 的 run 含 %s" % (rid, rln))
                    if rln == "t" and not text_of(rc):
                        errors.append("w:ins %s 存在空文本 run" % rid)
        else:
            for ch in element_children(el):
                if local_name(ch) != "r":
                    errors.append("w:del %s 含非 run 子元素 %s" % (rid, local_name(ch)))
                    continue
                for rc in element_children(ch):
                    rln = local_name(rc)
                    if rln == "t":
                        errors.append("w:del %s 含 w:t（应为 delText）" % rid)
                    elif rln == "delText":
                        if not text_of(rc):
                            errors.append("w:del %s 存在空 delText" % rid)
                    elif rln not in ("rPr", "delText"):
                        errors.append("w:del %s 的 run 含 %s" % (rid, rln))
    missing = sorted(new_rev_ids - seen, key=lambda x: int(x))
    if missing:
        errors.append("计划中的新修订 ID 缺失：%s" % "、".join(missing))
    return errors


def _check_comments(document_doc, prefix, comments_map, new_comment_ids, ops, result):
    errors = []
    order = {}
    anchors = {"commentRangeStart": [], "commentRangeEnd": [], "commentReference": []}
    for idx, el in enumerate(iter_elements_in_order(document_doc.documentElement)):
        ln = local_name(el)
        if ln in anchors:
            anchors[ln].append((idx, attr_local(el, "id"), el))
    for kind, entries in anchors.items():
        for idx, cid, el in entries:
            order[(kind, str(cid))] = idx
    for cid in new_comment_ids:
        cid = str(cid)
        starts = [e for e in anchors["commentRangeStart"] if str(e[1]) == cid]
        ends = [e for e in anchors["commentRangeEnd"] if str(e[1]) == cid]
        refs = [e for e in anchors["commentReference"] if str(e[1]) == cid]
        if len(starts) != 1 or len(ends) != 1 or len(refs) != 1:
            errors.append(
                "批注 %s 锚点不闭环：start=%d end=%d ref=%d"
                % (cid, len(starts), len(ends), len(refs))
            )
            continue
        if not (order[("commentRangeStart", cid)] < order[("commentRangeEnd", cid)]):
            errors.append("批注 %s 的 start 不在 end 之前" % cid)
        if not (
            order[("commentRangeEnd", cid)] < order[("commentReference", cid)]
        ):
            errors.append("批注 %s 的 reference 不在 end 之后" % cid)
        # reference run 必须带 CommentReference 样式
        ref = refs[0][2]
        run = ref.parentNode
        styled = False
        if run is not None and local_name(run) == "r":
            for rc in element_children(run):
                if local_name(rc) == "rPr":
                    for st in element_children(rc):
                        if (
                            local_name(st) == "rStyle"
                            and attr_local(st, "val") == "CommentReference"
                        ):
                            styled = True
        if not styled:
            errors.append("批注 %s 的 reference run 缺少 CommentReference 样式" % cid)
        comment = comments_map.get(cid)
        if comment is None:
            errors.append("word/comments.xml 缺少批注 %s" % cid)
        else:
            if not comment["author"]:
                errors.append("批注 %s 缺少 author" % cid)
            if not comment["date"] or not DATE_RE.match(comment["date"] or ""):
                errors.append("批注 %s 的 date 非法" % cid)
            result["authors"].add(comment["author"] or "")
    return errors


def _comment_anchor_paragraph_ids(document_doc, w_prefix):
    """批注 id -> 出现其锚点的段落 RL 集合（用于计划追溯与跨段检测）。"""
    paragraphs = collect_paragraphs(document_doc, w_prefix)
    mapping = {}
    for p in paragraphs:
        for el in iter_elements_in_order(p.node):
            ln = local_name(el)
            if ln in COMMENT_ANCHOR_TAGS:
                mapping.setdefault(str(attr_local(el, "id")), set()).add(p.pid)
    return mapping


def _comment_signatures(pkg):
    if not pkg.has("word/comments.xml"):
        return {}
    doc = parse_xml(pkg.get("word/comments.xml"), "word/comments.xml")
    out = {}
    for el in iter_elements_in_order(doc.documentElement):
        if local_name(el) == "comment":
            out[str(attr_local(el, "id"))] = el.toxml()
    return out


def _revision_signatures(pkg, part="word/document.xml"):
    doc = parse_xml(pkg.get(part), part)
    out = {}
    for el in iter_elements_in_order(doc.documentElement):
        ln = local_name(el)
        if ln in REV_TAGS:
            key = (ln, str(attr_local(el, "id")))
            out[key] = el.toxml()
    return out


def _anchor_counts(pkg, part="word/document.xml"):
    """按 (段落 RL, 锚类型, id) 统计，确保既有批注锚未被移动或删除。"""
    doc = parse_xml(pkg.get(part), part)
    w_prefix, _ = resolve_w_prefix(doc.documentElement)
    counts = Counter()
    for paragraph in collect_paragraphs(doc, w_prefix):
        for el in iter_elements_in_order(paragraph.node):
            ln = local_name(el)
            if ln in COMMENT_ANCHOR_TAGS:
                counts[(paragraph.pid, ln, str(attr_local(el, "id")))] += 1
    return counts


def _rels_entries(pkg):
    doc = parse_xml(pkg.get("word/_rels/document.xml.rels"), "document.xml.rels")
    entries = set()
    for el in element_children(doc.documentElement):
        if local_name(el) != "Relationship":
            continue
        entries.add(
            (
                el.getAttribute("Id"),
                el.getAttribute("Type"),
                el.getAttribute("Target"),
                el.getAttribute("TargetMode"),
            )
        )
    return entries


def _content_type_entries(pkg):
    doc = parse_xml(pkg.get("[Content_Types].xml"), "[Content_Types].xml")
    entries = set()
    for el in element_children(doc.documentElement):
        ln = local_name(el)
        if ln == "Override":
            entries.add((ln, el.getAttribute("PartName"), el.getAttribute("ContentType")))
        elif ln == "Default":
            entries.add((ln, el.getAttribute("Extension"), el.getAttribute("ContentType")))
    return entries


def _revision_id_counts(pkg, part="word/document.xml"):
    """修订 id -> 出现次数（跨 ins/del/moveFrom/moveTo 共用一个 ID 空间）。"""
    doc = parse_xml(pkg.get(part), part)
    counts = Counter()
    for el in iter_elements_in_order(doc.documentElement):
        if local_name(el) in REV_TAGS:
            counts[str(attr_local(el, "id"))] += 1
    return counts


def _comment_id_counts(pkg):
    """批注 id -> 出现次数（comments.xml 内）。"""
    counts = Counter()
    if not pkg.has("word/comments.xml"):
        return counts
    doc = parse_xml(pkg.get("word/comments.xml"), "word/comments.xml")
    for el in iter_elements_in_order(doc.documentElement):
        if local_name(el) == "comment":
            counts[str(attr_local(el, "id"))] += 1
    return counts


def _new_change_metadata(result_pkg, min_rev_id, min_comment_id):
    """提取结果中本轮新增修订/批注的 (author, date, label)。

    仅用于重建预期包前确认 author/date 唯一，不从结果复制任何 XML 作为预期。
    """
    meta = []
    doc = parse_xml(result_pkg.get("word/document.xml"), "word/document.xml")
    for el in iter_elements_in_order(doc.documentElement):
        ln = local_name(el)
        if ln not in REV_TAGS:
            continue
        rid = _max_int([attr_local(el, "id")])
        if rid > min_rev_id:
            meta.append(
                (
                    attr_local(el, "author"),
                    attr_local(el, "date"),
                    "%s %s" % (ln, attr_local(el, "id")),
                )
            )
    if result_pkg.has("word/comments.xml"):
        cdoc = parse_xml(
            result_pkg.get("word/comments.xml"), "word/comments.xml"
        )
        for el in iter_elements_in_order(cdoc.documentElement):
            if local_name(el) != "comment":
                continue
            cid = _max_int([attr_local(el, "id")])
            if cid > min_comment_id:
                meta.append(
                    (
                        attr_local(el, "author"),
                        attr_local(el, "date"),
                        "comment %s" % attr_local(el, "id"),
                    )
                )
    return meta


def verify_core(input_pkg, result_pkg, plan):
    """核心验证；返回报告 dict（ok / checks / errors）。"""
    errors = []
    checks = {}

    ctx = DocContext(input_pkg)
    if input_pkg.issues:
        errors.append("源文件包结构异常：%s" % "；".join(input_pkg.issues))
    if ctx.safety["blockers"]:
        errors.append(
            "源文件包含不允许写入的保护结构：%s"
            % "；".join(ctx.safety["blockers"])
        )
    validate_plan(ctx, plan)
    rebuilt = build_plan(ctx, plan["edits"], created_at=plan.get("created_at"))
    ops = rebuilt["ops"]

    # 1. 包结构与成员字节
    if result_pkg.issues:
        errors.append("结果文件包结构异常：%s" % "；".join(result_pkg.issues))
    if not result_pkg.has("word/document.xml"):
        raise PackageError("结果文件缺少 word/document.xml")
    input_names = set(input_pkg.members)
    result_names = set(result_pkg.members)
    extra = result_names - input_names
    missing = input_names - result_names
    if extra - ALLOWED_ADDED_MEMBERS:
        errors.append("结果文件新增了未允许的成员：%s" % "、".join(sorted(extra)))
    if missing:
        errors.append("结果文件缺少原成员：%s" % "、".join(sorted(missing)))
    changed = []
    for name in sorted(input_names & result_names):
        if name in ALLOWED_CHANGED_MEMBERS:
            continue
        if input_pkg.members[name] != result_pkg.members[name]:
            changed.append(name)
    if changed:
        errors.append("未允许变更的 ZIP 成员字节被修改：%s" % "、".join(changed))
    checks["zip_members_preserved"] = not changed and not (extra - ALLOWED_ADDED_MEMBERS) and not missing

    # 2. XML 可解析
    result_doc = parse_xml(result_pkg.get("word/document.xml"), "word/document.xml")
    if local_name(result_doc.documentElement) != "document":
        errors.append("结果 word/document.xml 根元素不是 w:document")
    try:
        result_prefix, result_ns = resolve_w_prefix(result_doc.documentElement)
    except PackageError as exc:
        result_prefix, result_ns = None, None
        errors.append(str(exc))
    if result_ns != ctx.namespace:
        errors.append(
            "结果文档命名空间与源文件不一致：%r != %r"
            % (result_ns, ctx.namespace)
        )
    if result_prefix != ctx.w_prefix:
        errors.append(
            "结果文档命名空间前缀与源文件不一致：%r != %r"
            % (result_prefix, ctx.w_prefix)
        )
    if not result_pkg.has("word/comments.xml"):
        raise PackageError("结果文件缺少 word/comments.xml")
    comments_bytes = result_pkg.get("word/comments.xml")
    comments_doc = parse_xml(comments_bytes, "word/comments.xml")
    if local_name(comments_doc.documentElement) != "comments":
        errors.append("word/comments.xml 根元素不是 w:comments")
    rels_doc = parse_xml(
        result_pkg.get("word/_rels/document.xml.rels"), "document.xml.rels"
    )
    if local_name(rels_doc.documentElement) != "Relationships":
        errors.append("document.xml.rels 根元素不是 Relationships")
    ct_doc = parse_xml(result_pkg.get("[Content_Types].xml"), "[Content_Types].xml")
    if local_name(ct_doc.documentElement) != "Types":
        errors.append("[Content_Types].xml 根元素不是 Types")
    if prefix_for_ns(ct_doc.documentElement, CT_NS) is None:
        errors.append("[Content_Types].xml 缺少内容类型命名空间声明")
    if prefix_for_ns(rels_doc.documentElement, REL_NS) is None:
        errors.append("document.xml.rels 缺少关系命名空间声明")
    checks["xml_well_formed"] = True

    # 3. ID 空间
    input_ids = scan_existing_ids(input_pkg)
    max_rev = input_ids["max_revision_id"]
    max_comment = input_ids["max_comment_id"]
    input_comments = _parse_comment_texts(
        input_pkg.get("word/comments.xml") if input_pkg.has("word/comments.xml") else None
    )

    result_existing = collect_existing(result_pkg)
    new_revs = [
        r for r in result_existing["revisions"]
        if r["part"] == "word/document.xml"
        and _max_int([r.get("id")]) > max_rev
    ]
    new_rev_ids = {str(r["id"]) for r in new_revs}
    result_comments = _parse_comment_texts(comments_bytes)
    new_comment_ids = sorted(
        (cid for cid in result_comments if cid not in input_comments),
        key=lambda x: int(x),
    )
    if any(int(cid) <= max_comment for cid in new_comment_ids):
        errors.append("新批注 ID 未在既有最大 ID 之后分配")
    if any(int(rid) <= max_rev for rid in new_rev_ids):
        errors.append("新修订 ID 未在既有最大 ID 之后分配")
    checks["id_namespaces_advanced"] = not any(
        int(cid) <= max_comment for cid in new_comment_ids
    ) and not any(int(rid) <= max_rev for rid in new_rev_ids)

    ids_unique = len(new_rev_ids) == len(new_revs) and len(
        set(new_comment_ids)
    ) == len(new_comment_ids)
    if not ids_unique:
        errors.append("新增修订/批注 ID 存在重复")
    checks["new_ids_unique"] = ids_unique

    result = {"authors": set()}
    rev_errors = _check_new_revision_structure(
        result_doc, ctx.w_prefix, new_rev_ids, result
    )
    errors.extend(rev_errors)
    checks["new_revisions_structure"] = not rev_errors

    # 新修订数量必须与计划中的 diff 粒度一致（防止空修订或计划外增删）
    expected_ins = 0
    expected_del = 0
    for op in ops:
        if op["kind"] != "replace":
            continue
        para = ctx.by_id[op["para"]]
        runs = simple_runs(para.node)
        for d in op["diff"]:
            if d["new"]:
                expected_ins += 1
            if d["src_end"] > d["src_start"]:
                for run in runs:
                    if run.start == run.end:
                        continue
                    if run.start < d["src_end"] and run.end > d["src_start"]:
                        expected_del += 1
    actual_ins = sum(1 for r in new_revs if r["tag"] == "ins")
    actual_del = sum(1 for r in new_revs if r["tag"] == "del")
    count_ok = actual_ins == expected_ins and actual_del == expected_del
    if not count_ok:
        errors.append(
            "新增修订数量与计划粒度不一致：ins %d/%d，del %d/%d"
            % (actual_ins, expected_ins, actual_del, expected_del)
        )
    checks["new_revision_count_matches_plan"] = count_ok

    # 4. 批注闭环与关系/内容类型
    comment_errors = _check_comments(
        result_doc, ctx.w_prefix, result_comments, new_comment_ids, ops, result
    )
    errors.extend(comment_errors)
    checks["comments_closure"] = not comment_errors

    rel_targets = _relations_comments_target(
        result_pkg.get("word/_rels/document.xml.rels")
    )
    rel_ok = "word/comments.xml" in rel_targets
    if not rel_ok:
        errors.append("结果文件缺少 comments 关系（Target=comments.xml）")
    ct_ok = _has_comments_content_type(result_pkg.get("[Content_Types].xml"))
    if not ct_ok:
        errors.append("结果文件缺少 comments 内容类型注册")
    checks["comments_registration"] = rel_ok and ct_ok

    # 5. 既有批注保留
    preserved = True
    for cid, info in input_comments.items():
        got = result_comments.get(cid)
        if (
            got is None
            or got["author"] != info["author"]
            or got["date"] != info["date"]
            or got["text"] != info["text"]
        ):
            preserved = False
            errors.append("既有批注 %s 未被原样保留" % cid)
    checks["existing_comments_preserved"] = preserved

    input_comment_xml = _comment_signatures(input_pkg)
    result_comment_xml = _comment_signatures(result_pkg)
    comment_xml_ok = all(
        result_comment_xml.get(cid) == value
        for cid, value in input_comment_xml.items()
    )
    if not comment_xml_ok:
        errors.append("既有批注的 XML 结构/格式未被原样保留")
    checks["existing_comment_elements_preserved"] = comment_xml_ok

    # 6. 作者一致性
    authors = {a for a in result["authors"] if a}
    if len(authors) != 1:
        errors.append("本轮新增修订/批注的 author 不一致或为空：%r" % sorted(authors))
    checks["author_consistent"] = len(authors) == 1

    # 6b. 既有修订结构、批注锚、关系与内容类型原样保留
    input_revs = _revision_signatures(input_pkg)
    result_revs = _revision_signatures(result_pkg)
    revs_preserved = all(
        result_revs.get(key) == value for key, value in input_revs.items()
    )
    if not revs_preserved:
        errors.append("既有修订元素结构未被原样保留")
    checks["existing_revisions_preserved"] = revs_preserved

    # 含既有修订/批注锚的段落应完全未触碰；按序列化内容逐段比对
    result_paragraphs = collect_paragraphs(result_doc, ctx.w_prefix)
    op_pids = {op["para"] for op in ops}
    untouched_ok = True
    for idx, paragraph in enumerate(ctx.paragraphs):
        if paragraph.pid in op_pids:
            continue
        if idx >= len(result_paragraphs) or (
            result_paragraphs[idx].node.toxml() != paragraph.node.toxml()
        ):
            untouched_ok = False
            errors.append("未列入计划的段落 %s 被改动" % paragraph.pid)
    checks["untouched_paragraphs_unchanged"] = untouched_ok

    locked_ok = True
    for idx, paragraph in enumerate(ctx.paragraphs):
        locked = any(
            local_name(el) in REV_TAGS or local_name(el) in COMMENT_ANCHOR_TAGS
            for el in iter_elements_in_order(paragraph.node)
        )
        if not locked:
            continue
        if idx >= len(result_paragraphs) or (
            result_paragraphs[idx].node.toxml() != paragraph.node.toxml()
        ):
            locked_ok = False
            errors.append("含既有修订/批注锚的段落 %s 被改动" % paragraph.pid)
    checks["locked_paragraphs_unchanged"] = locked_ok

    input_anchors = _anchor_counts(input_pkg)
    result_anchors = _anchor_counts(result_pkg)
    anchors_preserved = all(
        result_anchors.get(key, 0) >= count for key, count in input_anchors.items()
    )
    if not anchors_preserved:
        errors.append("既有批注范围锚（其它段落）未被原样保留")
    checks["existing_comment_anchors_preserved"] = anchors_preserved

    input_rels = _rels_entries(input_pkg)
    result_rels = _rels_entries(result_pkg)
    rels_preserved = input_rels <= result_rels
    if not rels_preserved:
        errors.append("既有 document.xml 关系条目未被原样保留")
    checks["relationship_entries_preserved"] = rels_preserved

    input_ct = _content_type_entries(input_pkg)
    result_ct = _content_type_entries(result_pkg)
    ct_preserved = input_ct <= result_ct
    if not ct_preserved:
        errors.append("既有 Content_Types 条目未被原样保留")
    checks["content_type_entries_preserved"] = ct_preserved

    # 7. 视图等价
    baseline_doc = ctx.doc
    baseline_sig = document_signature(baseline_doc, ctx.w_prefix)
    baseline_texts = document_view_texts(baseline_doc, ctx.w_prefix)

    reject_doc = parse_xml(result_pkg.get("word/document.xml"), "word/document.xml")
    simulate_new_changes(
        reject_doc, ctx.w_prefix, new_rev_ids, new_comment_ids, "reject"
    )
    reject_sig = document_signature(reject_doc, ctx.w_prefix)
    reject_ok = reject_sig == baseline_sig
    if not reject_ok:
        errors.append("拒绝本轮修订后的视图与基线（段落序列/run 样式）不一致")
    checks["reject_equals_baseline"] = reject_ok

    accept_doc = parse_xml(result_pkg.get("word/document.xml"), "word/document.xml")
    simulate_new_changes(
        accept_doc, ctx.w_prefix, new_rev_ids, new_comment_ids, "accept"
    )
    accept_texts = document_view_texts(accept_doc, ctx.w_prefix)
    expected = list(baseline_texts)
    target_ok = True
    for op in ops:
        para = ctx.by_id.get(op["para"])
        if para is None:
            target_ok = False
            errors.append("计划段落 %s 不存在" % op["para"])
            continue
        idx = ctx.paragraphs.index(para)
        if op["kind"] == "replace":
            text = baseline_texts[idx]
            positions = find_occurrences(text, op["old"])
            if len(positions) != 1:
                target_ok = False
                errors.append("段落 %s 的 old 在基线中不再唯一" % op["para"])
                continue
            pos = positions[0]
            expected[idx] = text[:pos] + op["new"] + text[pos + len(op["old"]):]
            if expected[idx] != op["target_text"]:
                target_ok = False
                errors.append("段落 %s 的目标文本与计划不一致" % op["para"])
        else:
            expected[idx] = baseline_texts[idx]
    if accept_texts != expected:
        target_ok = False
        diff_pids = [
            ctx.paragraphs[i].pid
            for i in range(min(len(accept_texts), len(expected)))
            if accept_texts[i] != expected[i]
        ]
        errors.append(
            "接受本轮修订后的文本与 old/new 独立计算的目标不一致%s"
            % (("（段落：%s）" % "、".join(diff_pids)) if diff_pids else "")
        )
    checks["accept_matches_target"] = target_ok

    # 8. 原修订未被模拟接受：接受/拒绝视图中原修订仍在
    baseline_rev_ids = {
        str(r["id"]) for r in input_ids["existing"]["revisions"]
        if r["part"] == "word/document.xml" and r["id"] is not None
    }
    accept_rev_ids = {
        str(attr_local(el, "id"))
        for el in iter_elements_in_order(accept_doc.documentElement)
        if local_name(el) in REV_TAGS
    }
    reject_rev_ids = {
        str(attr_local(el, "id"))
        for el in iter_elements_in_order(reject_doc.documentElement)
        if local_name(el) in REV_TAGS
    }
    originals_kept = baseline_rev_ids <= accept_rev_ids and baseline_rev_ids <= reject_rev_ids
    if not originals_kept:
        errors.append("模拟接受/拒绝本轮修订时，原有修订被错误处理")
    checks["original_revisions_untouched"] = originals_kept

    # 9. 批注内容与计划一致
    content_ok = True
    for op, cid in zip(ops, new_comment_ids):
        got = result_comments.get(str(cid))
        want = op.get("reason") if op["kind"] == "replace" else op.get("comment")
        if got is None or "".join(got["text"].split()) != "".join((want or "").split()):
            content_ok = False
            errors.append("批注 %s 内容与计划不一致" % cid)
    if len(new_comment_ids) != len(ops):
        content_ok = False
        errors.append(
            "新增批注数量 %d 与计划 op 数 %d 不一致"
            % (len(new_comment_ids), len(ops))
        )
    checks["comments_match_plan"] = content_ok

    # 10. 批注锚点在计划段落内
    anchor_para = _comment_anchor_paragraph_ids(result_doc, ctx.w_prefix)
    anchor_ok = True
    for op, cid in zip(ops, new_comment_ids):
        pids = anchor_para.get(str(cid))
        if pids != {op["para"]}:
            anchor_ok = False
            errors.append(
                "批注 %s 锚定在 %s，计划要求 %s"
                % (cid, sorted(pids) if pids else None, op["para"])
            )
    checks["comment_anchor_in_planned_paragraph"] = anchor_ok

    # 11. 结果中不得新增重复 ID（本生成器从不为一个 ID 重复输出）
    input_rev_counts = _revision_id_counts(input_pkg)
    result_rev_counts = _revision_id_counts(result_pkg)
    dup_revs = sorted(
        rid
        for rid, count in result_rev_counts.items()
        if count > 1 and count > input_rev_counts.get(rid, 0)
    )
    input_comment_counts = _comment_id_counts(input_pkg)
    result_comment_counts = _comment_id_counts(result_pkg)
    dup_comments = sorted(
        cid
        for cid, count in result_comment_counts.items()
        if count > 1 and count > input_comment_counts.get(cid, 0)
    )
    dup_ok = not dup_revs and not dup_comments
    if dup_revs:
        errors.append("结果新增了重复的修订 ID：%s" % "、".join(dup_revs[:5]))
    if dup_comments:
        errors.append(
            "结果新增了重复的批注 ID：%s" % "、".join(dup_comments[:5])
        )
    checks["no_duplicate_change_ids"] = dup_ok

    # 12. 完整授权输出包：以结果中唯一 author/date + 全新 DocContext + 已校验的
    # plan 重建预期包，逐 ZIP 成员比对解压字节（不比 ZIP 时间等容器元数据）。
    # 这是重建一致性判据，不是独立语义证明；上面的独立文字签名检查仍保留。
    meta = _new_change_metadata(result_pkg, max_rev, max_comment)
    meta_problems = []
    if not meta:
        meta_problems.append("结果中没有可识别的本轮新增修订/批注元数据")
    else:
        authors_found = sorted({(item[0] or "") for item in meta})
        dates_found = sorted({(item[1] or "") for item in meta})
        if len(authors_found) != 1 or not authors_found[0]:
            meta_problems.append(
                "本轮新增修订/批注的 author 不唯一或为空：%r" % authors_found
            )
        if len(dates_found) != 1:
            meta_problems.append(
                "本轮新增修订/批注的 date 不唯一：%r" % dates_found
            )
        elif not DATE_RE.match(dates_found[0]):
            meta_problems.append(
                "本轮新增修订/批注的 date 非法：%r" % dates_found[0]
            )
    if meta_problems:
        errors.extend(meta_problems)
        checks["rebuilt_package_matches"] = False
    else:
        expected_ok = True
        expected_author = meta[0][0]
        expected_date = meta[0][1]
        try:
            expected_author = validate_author(expected_author)
            fresh_ctx = DocContext(input_pkg)
            expected_bytes, _ = build_result(
                fresh_ctx, rebuilt, expected_author, timestamp=expected_date
            )
            expected_pkg = package_from_bytes(expected_bytes, "预期重建产物")
        except RedlineError as exc:
            expected_ok = False
            errors.append("无法重建预期结果包：%s" % exc.message)
        except Exception as exc:  # 兜底：重建失败按验证失败处理
            expected_ok = False
            errors.append(
                "重建预期结果包时发生内部错误：%s：%s"
                % (type(exc).__name__, exc)
            )
        else:
            expected_names = sorted(expected_pkg.members)
            actual_names = sorted(result_pkg.members)
            if expected_names != actual_names:
                expected_ok = False
                extra_names = sorted(set(actual_names) - set(expected_names))
                missing_names = sorted(set(expected_names) - set(actual_names))
                errors.append(
                    "结果 ZIP 成员集合与预期重建不一致（多出 %s；缺少 %s）"
                    % (extra_names[:5] or "无", missing_names[:5] or "无")
                )
            mismatched = [
                name
                for name in expected_names
                if name in result_pkg.members
                and result_pkg.members[name] != expected_pkg.members[name]
            ]
            if mismatched:
                expected_ok = False
                errors.append(
                    "结果成员字节与按 plan 重建的预期不一致：%s"
                    % "、".join(mismatched[:5])
                )
        checks["rebuilt_package_matches"] = expected_ok

    return {
        "ok": not errors,
        "input_sha256": input_pkg.sha256,
        "result_sha256": result_pkg.sha256,
        "checks": checks,
        "errors": errors,
        "new_revisions": {
            "insertions": sum(1 for r in new_revs if r["tag"] == "ins"),
            "deletions": sum(1 for r in new_revs if r["tag"] == "del"),
            "ids": sorted(new_rev_ids, key=lambda x: int(x)),
        },
        "new_comments": {"count": len(new_comment_ids), "ids": new_comment_ids},
        "author": next(iter(authors)) if authors else None,
        "notes": [
            "本报告为脚本级结构验证，不等于 Microsoft Word / WPS 的实测渲染结果。",
            "「完整授权输出包」判据：用结果中唯一的新增 author/date 与全新 "
            "DocContext、已校验的 plan 重建预期包，再逐 ZIP 成员比对解压字节"
            "（不比较 ZIP 时间等容器元数据）。它是重建一致性比对，不构成独立"
            "语义证明；原有的独立文字签名与视图等价检查仍保留。",
        ],
    }


# --------------------------------------------------------------------------
# CLI 实现
# --------------------------------------------------------------------------


def cmd_inspect(args):
    pkg = load_package(args.file)
    ctx = DocContext(pkg)
    existing = collect_existing(pkg, strict=False)
    paragraphs = []
    for p in ctx.paragraphs:
        paragraphs.append(
            {
                "id": p.pid,
                "location": p.location,
                "text": render_visible_text(p.node),
                "editable": p.editable,
                "blockers": p.blockers,
            }
        )
    revisions = existing["revisions"]
    comments = existing["comments"]
    modern = [
        n for n in MODERN_COMMENT_PARTS if pkg.has(n)
    ]
    unsupported = {}
    for p in ctx.paragraphs:
        for b in p.blockers:
            unsupported[b] = unsupported.get(b, 0) + 1
    warnings = list(ctx.safety["warnings"]) + list(existing["errors"])
    if unsupported:
        warnings.append(
            "存在复杂/未覆盖结构，相关段落不可编辑（见 unsupported_structure_counts）"
        )
    hidden_pids = [
        p.pid
        for p in ctx.paragraphs
        if any(b.startswith("hidden_text") for b in p.blockers)
    ]
    if hidden_pids:
        warnings.append(
            "检测到隐藏文本属性（hidden_text：vanish/specVanish/webHidden），"
            "相关段落不可编辑：%s" % "、".join(hidden_pids[:10])
        )
    return {
        "ok": True,
        "tool": TOOL,
        "command": "inspect",
        "file": os.path.basename(os.path.abspath(args.file)),
        "sha256": pkg.sha256,
        "size": len(pkg.raw),
        "package": {
            "safe_for_apply": not ctx.safety["blockers"],
            "blockers": ctx.safety["blockers"],
            "warnings": warnings,
            "namespace": ctx.safety["namespace"],
            "member_count": len(pkg.infos),
            "document_protection": ctx.safety["document_protection"],
            "write_protection": ctx.safety["write_protection"],
        },
        "warnings": warnings,
        "paragraphs": paragraphs,
        "editable_paragraph_count": sum(1 for p in ctx.paragraphs if p.editable),
        "unsupported_structure_counts": unsupported,
        "existing_revisions": {
            "count": len(revisions),
            "ids": [r["id"] for r in revisions],
            "parts": sorted({r["part"] for r in revisions}),
            "items": revisions,
        },
        "existing_comments": {
            "has_part": pkg.has("word/comments.xml"),
            "count": len(comments),
            "ids": [c["id"] for c in comments],
            "modern_extensions": modern,
            "document_anchors": existing["anchors"],
        },
    }


def _read_json(path, label):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except FileNotFoundError:
        raise InputError("%s 不存在：%s" % (label, path)) from None
    except json.JSONDecodeError as exc:
        raise InputError("%s 不是合法 JSON：%s" % (label, exc)) from exc


def cmd_plan(args):
    if os.path.abspath(args.out) == os.path.abspath(args.file):
        raise InputError("plan 输出不能覆盖输入 docx")
    edits = _read_json(args.edits, "edits.json")
    ctx = load_context(args.file, for_write=True)
    plan = build_plan(ctx, edits)
    data = json_dumps_pretty(plan).encode("utf-8")
    if os.path.exists(os.path.abspath(args.out)):
        raise InputError("plan 输出已存在，拒绝覆盖：%s" % args.out)
    publish_new_file(args.out, data)
    return {
        "ok": True,
        "command": "plan",
        "plan": os.path.abspath(args.out),
        "source_sha256": ctx.sha256,
        "ops": len(plan["ops"]),
        "comment_count": plan["comment_count"],
        "warnings": plan["warnings"],
        "fingerprint": plan["fingerprint"]["value"],
    }


def _load_plan_file(path):
    return _read_json(path, "plan.json")


def cmd_apply(args):
    in_abs = os.path.abspath(args.file)
    out_abs = os.path.abspath(args.out)
    plan_abs = os.path.abspath(args.plan)
    if out_abs == in_abs:
        raise InputError("拒绝 in-place 改写：输出不能等于输入")
    if out_abs == plan_abs:
        raise InputError("输出不能覆盖 plan.json")
    if os.path.exists(out_abs):
        raise InputError("输出文件已存在，拒绝覆盖：%s" % args.out)
    author = validate_author(args.author)
    ctx = load_context(args.file, for_write=True)
    plan = _load_plan_file(args.plan)
    rebuilt = validate_plan(ctx, plan)

    result_bytes, stats = build_result(ctx, rebuilt, author)
    # 发布前先自动验证（纯内存）
    result_pkg = package_from_bytes(result_bytes, os.path.basename(out_abs))
    report = verify_core(ctx.pkg, result_pkg, rebuilt)
    if not report["ok"]:
        raise VerifyError(
            "自动验证失败，未发布结果文件：%s" % "；".join(report["errors"]),
            checks=report["checks"],
        )
    publish_new_file(out_abs, result_bytes)
    return {
        "ok": True,
        "command": "apply",
        "result": out_abs,
        "result_sha256": sha256_bytes(result_bytes),
        "source_sha256": ctx.sha256,
        "author": author,
        "new_revisions": stats["new_revision_ids"],
        "new_comments": stats["new_comment_ids"],
        "verification": {
            "ok": True,
            "checks": report["checks"],
            "notes": report["notes"],
        },
    }


def cmd_verify(args):
    input_pkg = load_package(args.input)
    result_pkg = load_package(args.result)
    plan = _load_plan_file(args.plan)
    report = verify_core(input_pkg, result_pkg, plan)
    if not report["ok"]:
        raise VerifyError("；".join(report["errors"]), report=report)
    report["tool"] = TOOL
    report["command"] = "verify"
    return report


DESCRIPTION = (
    "受限、可验证的 Word 红线引擎：inspect / plan / apply / verify。\n"
    "不动原件；apply 输出必须不存在；所有写入原子独占发布；发布前自动验证。\n"
    "只支持正文与普通表格单元格中的简单单段文字，隐藏文本（vanish/"
    "specVanish/webHidden）与复杂结构一律 fail-closed。"
)

EPILOG = """\
edits.json 格式（数组，每一项二选一）：
  替换/插入/删除：
    {"id":"R1","para":"RL0003","old":"段内唯一精确子串",
     "new":"目标文本（改写/包含 old；new 为空表示删除该片段；不得等于 old）",
     "reason":"替换理由（成为范围批注）",
     "broad_reason":"大幅改写理由（仅当 plan 给出粒度警告时必填）"}
  仅批注：
    {"id":"C1","para":"RL0005","old":"段内唯一精确锚文本",
     "comment":"批注内容"}

规则摘要：
  - para 使用 inspect 输出的 RL0001 编号；同一段落只允许一个 op。
  - old 为空、跨段、段内命中 0 次或多次都拒绝。
  - 一个 op 只产生一条批注；替换 op 必定给 reason，仅批注 op 不产生修订。
  - old==new 属于空修改；new 为空表示删除 old，属于实质删除而非空修改。
  - 大幅改写（>60 字，或长段 >60% 改动，或 >6 块变化）需 broad_reason。

退出码：0 成功；1 参数/编辑/计划错误；2 文件/包结构错误；3 验证失败。
"""


def build_parser():
    parser = argparse.ArgumentParser(
        prog="hetong_redline.py",
        description=DESCRIPTION,
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser(
        "inspect",
        help="只读检视：段落 RL 编号、可见文本、可编辑性、既有修订/批注",
    )
    p.add_argument("file", help="输入 docx（只读）")
    p.set_defaults(func=cmd_inspect)

    p = sub.add_parser(
        "plan", help="校验 edits.json 并生成可复核计划（不写 docx）"
    )
    p.add_argument("file", help="输入 docx（只读）")
    p.add_argument("edits", help="edits.json")
    p.add_argument("-o", "--out", required=True, help="输出 plan.json（必须不存在）")
    p.set_defaults(func=cmd_plan)

    p = sub.add_parser(
        "apply", help="按 plan 生成带修订/批注的新 docx（自动验证、独占发布）"
    )
    p.add_argument("file", help="输入 docx（只读，不得与输出相同）")
    p.add_argument("plan", help="plan.json")
    p.add_argument("-o", "--out", required=True, help="输出 docx（必须不存在）")
    p.add_argument(
        "--author", default="合同审阅", help="修订/批注作者（默认：合同审阅）"
    )
    p.set_defaults(func=cmd_apply)

    p = sub.add_parser(
        "verify", help="验证结果文件与 plan/基线的一致性"
    )
    p.add_argument("input", help="原始 docx")
    p.add_argument("result", help="apply 生成的 docx")
    p.add_argument("plan", help="plan.json")
    p.set_defaults(func=cmd_verify)
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        payload = args.func(args)
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    except RedlineError as exc:
        print(json.dumps(exc.as_dict(), ensure_ascii=False, indent=2))
        return exc.code
    except BrokenPipeError:
        return 0
    except OSError as exc:
        print(
            json.dumps(
                {"ok": False, "error": "文件系统错误：%s" % exc, "code": 2},
                ensure_ascii=False,
                indent=2,
            )
        )
        return 2
    except Exception as exc:  # 兜底：不输出 traceback，fail-closed
        print(
            json.dumps(
                {
                    "ok": False,
                    "error": "内部错误（%s）：%s" % (type(exc).__name__, exc),
                    "code": 2,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 2


if __name__ == "__main__":
    sys.exit(main())
