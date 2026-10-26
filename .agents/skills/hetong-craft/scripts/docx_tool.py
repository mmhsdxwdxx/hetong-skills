#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_tool.py — 纯标准库（Python >= 3.8）的 docx 段落级读写工具，供合同工作使用。

设计目标：
- 提取全文为带稳定段落编号（P0001）与表格坐标（T0001 R01 C01）的文本，便于审阅与定位；
- 在不破坏版式的前提下做定点修改：跨 run 文本替换、整段改写、克隆插入段落、
  删除段落、克隆新增表格行；命名空间前缀保持原样（含 mc:Ignorable 前缀），
  保证 Word 可正常打开；
- 一切修改默认先备份原件，支持 --dry-run 预演与批量 edits.json。

子命令：
  extract       提取全文（含表格、可选页眉页脚）
  search        按子串/正则定位段落；--numbers 预设金额税率期限类
  replace       跨 run 文本替换
  set           整段改写（保留段落格式与首个 run 的字符格式）
  insert-after  在指定段后克隆插入新段（继承该段格式）
  delete        删除段落
  add-row       在表格中克隆新增一行
  batch         按 edits.json 批量执行（所有编号以执行前文档状态为准）
  verify        校验 zip 与 XML 完整性

退出码：0 成功；1 参数/匹配错误；2 文件/结构错误。
"""

import argparse
import copy
import json
import os
import re
import shutil
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, NoReturn, cast

# Windows 控制台默认编码可能不是 UTF-8；命令行输出统一为 UTF-8。
for _stream in (sys.stdout, sys.stderr):
    try:
        cast(Any, _stream).reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def qn(local):
    return "{%s}%s" % (W, local)


# ---------------------------------------------------------------- 命名空间

def collect_namespaces(xml_text):
    """从原始 XML 文本的根元素上收集 prefix -> uri 映射。"""
    ns = {}
    m = re.search(r"<w:document\b[^>]*>", xml_text, re.S)
    head = m.group(0) if m else xml_text[:8000]
    for pm in re.finditer(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', head):
        ns[pm.group(1)] = pm.group(2)
    m2 = re.search(r'xmlns="([^"]+)"', head)
    if m2:
        ns[""] = m2.group(1)
    return ns


def register_namespaces(ns):
    for prefix, uri in ns.items():
        if prefix.lower() in ("xml", "xmlns"):
            continue
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass


def serialize_root(root, ns):
    """序列化根元素，并确保根标签声明了原始文档出现过的全部前缀。

    ElementTree 只输出实际用到的命名空间；mc:Ignorable 等属性可能按前缀名
    引用未使用的前缀，缺声明会导致 Word 校验失败，因此逐一补齐。
    """
    body = ET.tostring(root, encoding="unicode")
    m = re.match(r"<w:document\b[^>]*>", body, re.S)
    if not m:
        # 极端情况下根元素不是 w:document，直接返回
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body
    tag = m.group(0)
    add = []
    for prefix, uri in ns.items():
        if prefix == "":
            if 'xmlns="' not in tag:
                add.append(' xmlns="%s"' % uri)
        else:
            if ('xmlns:%s="' % prefix) not in tag:
                add.append(' xmlns:%s="%s"' % (prefix, uri))
    if add:
        new_tag = tag[:-1] + "".join(add) + ">"
        body = new_tag + body[len(tag):]
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body


# ---------------------------------------------------------------- 文档装载

class Docx(object):
    def __init__(self, path):
        self.path = path
        if not zipfile.is_zipfile(path):
            fail(2, "不是有效的 docx（zip）文件：%s" % path)
        with zipfile.ZipFile(path, "r") as z:
            self.entries = {}
            for info in z.infolist():
                self.entries[info.filename] = z.read(info.filename)
        if "word/document.xml" not in self.entries:
            fail(2, "docx 内缺少 word/document.xml：%s" % path)
        self.doc_xml_text = self.entries["word/document.xml"].decode("utf-8")
        self.ns = collect_namespaces(self.doc_xml_text)
        register_namespaces(self.ns)
        try:
            self.root = ET.fromstring(self.doc_xml_text)
        except ET.ParseError as e:
            fail(2, "word/document.xml 解析失败：%s" % e)
        self.body = self.root.find(qn("body"))
        if self.body is None:
            fail(2, "word/document.xml 缺少 w:body")
        self.styles = self._load_styles()
        self._reindex()

    def _load_styles(self):
        """styleId -> 可读名称（styles.xml 缺失时用 styleId）。"""
        styles = {}
        data = self.entries.get("word/styles.xml")
        if not data:
            return styles
        try:
            sroot = ET.fromstring(data.decode("utf-8"))
        except ET.ParseError:
            return styles
        for st in sroot.findall(qn("style")):
            sid = st.get(qn("styleId"))
            name_el = st.find(qn("name"))
            if sid:
                styles[sid] = name_el.get(qn("val")) if name_el is not None else sid
        return styles

    # ---------------- 索引：段落/表格按文档顺序编号 ----------------

    def _reindex(self):
        """重建 段落编号->元素、表格编号->元素、父元素映射。"""
        self.paras = {}      # "P0007" -> w:p 元素
        self.tables = {}     # "T0001" -> {"el": w:tbl, "rows": [[(cid, [p元素...])...]]}
        self.parent = {}     # 元素 -> 父元素
        self.para_loc = {}   # "P0007" -> 定位描述（body 或 表格坐标）
        p_count = 0
        t_count = 0
        assert self.body is not None  # __init__ rejects documents without a body
        for child in list(self.body):
            self.parent[child] = self.body
            if child.tag == qn("p"):
                p_count += 1
                pid = "P%04d" % p_count
                self.paras[pid] = child
                self.para_loc[pid] = "正文"
            elif child.tag == qn("tbl"):
                t_count += 1
                tid = "T%04d" % t_count
                tbl_info = {"el": child, "rows": []}
                self.tables[tid] = tbl_info
                for ri, tr in enumerate(child.findall(qn("tr")), start=1):
                    self.parent[tr] = child
                    row_cells = []
                    tbl_info["rows"].append(row_cells)
                    for ci, tc in enumerate(tr.findall(qn("tc")), start=1):
                        self.parent[tc] = tr
                        cell_paras = []
                        row_cells.append(cell_paras)
                        for p in tc.findall(qn("p")):
                            self.parent[p] = tc
                            p_count += 1
                            pid = "P%04d" % p_count
                            self.paras[pid] = p
                            cell_paras.append(pid)
                            self.para_loc[pid] = "%s R%02d C%02d" % (tid, ri, ci)
                        # 单元格内的嵌套表格暂不展开（合同模板未出现）

    def para_style(self, p):
        ppr = p.find(qn("pPr"))
        if ppr is not None:
            ps = ppr.find(qn("pStyle"))
            if ps is not None:
                sid = ps.get(qn("val"))
                return self.styles.get(sid, sid)
        return ""

    def save(self, out_path, backup=True):
        """把修改后的 document.xml 写回新 docx；out_path 与原文件相同时先备份。"""
        # Clean-text editing must never flatten an existing review baseline.
        # Inspect ORIGINAL parts, not the already-mutated in-memory tree.
        revision_tags = {qn(x) for x in (
            "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange",
            "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange",
            "tblGridChange", "numberingChange", "cellIns", "cellDel", "cellMerge",
            "commentRangeStart", "commentRangeEnd", "commentReference",
        )}
        for part, raw in self.entries.items():
            if not part.startswith("word/") or not part.endswith(".xml"):
                continue
            if "comment" in part.lower():
                fail(2, "基线含批注部件；普通编辑工具禁止写入，请使用 hetong_redline.py 或经用户确认另建清洁基线。")
            try:
                original_root = ET.fromstring(raw)
            except ET.ParseError:
                fail(2, "基线XML不可解析，停止写入：%s" % part)
            if any(
                el.tag in revision_tags
                or (el.tag == qn("trackRevisions") and el.get(qn("val"), "true").lower() not in {"0", "false", "off"})
                for el in original_root.iter()
            ):
                fail(2, "基线含修订/批注或开启修订模式；禁止静默清洁写入，请使用 hetong_redline.py，不得自动接受历史修订。")
        new_xml = serialize_root(self.root, self.ns).encode("utf-8")
        if os.path.abspath(out_path) == os.path.abspath(self.path):
            if backup:
                stamp = time.strftime("%Y%m%dT%H%M%S")
                bak = "%s.bak.%s" % (self.path, stamp)
                shutil.copy2(self.path, bak)
                note("已备份原件：%s" % bak)
        tmp = out_path + ".tmp.%d" % os.getpid()
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in self.entries.items():
                if name == "word/document.xml":
                    data = new_xml
                z.writestr(name, data)
        os.replace(tmp, out_path)
        note("已写出：%s" % out_path)


# ---------------------------------------------------------------- 文本工具

def para_text(p):
    """段落纯文本：拼接所有 w:t（含超链接内文本），跳过域代码 instrText。"""
    parts = []
    for el in p.iter():
        if el.tag == qn("t"):
            parts.append(el.text or "")
    return "".join(parts)


def text_nodes(p):
    """段落内全部 w:t 元素列表（文档顺序）。"""
    return [el for el in p.iter() if el.tag == qn("t")]


def set_para_text(p, text):
    """整段改写：保留 pPr 与首个 run 的 rPr，删除其余内容。"""
    first_rpr = None
    for r in p.findall(qn("r")):
        rpr = r.find(qn("rPr"))
        if rpr is not None:
            first_rpr = copy.deepcopy(rpr)
            break
    ppr = p.find(qn("pPr"))
    for child in list(p):
        if child is not ppr:
            p.remove(child)
    r = ET.SubElement(p, qn("r"))
    if first_rpr is not None:
        r.append(first_rpr)
    t = ET.SubElement(r, qn("t"))
    if text != text.strip():
        t.set("{%s}space" % XML_NS, "preserve")
    t.text = text


def replace_in_para(p, old, new, count=0):
    """在单个段落内跨 run 替换 old -> new，返回替换次数。

    策略：命中的首个 w:t 承接替换后文本（继承其字符格式），其余被命中的
    w:t 截去对应片段。合同段落内格式通常一致，该策略对定点修订足够安全。
    count=0 表示全部替换。
    """
    nodes = text_nodes(p)
    if not nodes:
        return 0
    texts = [n.text or "" for n in nodes]
    full = "".join(texts)
    if old not in full:
        return 0
    replaced = 0
    pos = 0
    while True:
        i = full.find(old, pos)
        if i < 0:
            break
        if count and replaced >= count:
            break
        j = i + len(old)
        # 把 [i, j) 区间映射到各 w:t
        offset = 0
        first_done = False
        for n, txt in zip(nodes, texts):
            if txt is None:
                txt = ""
            a, b = offset, offset + len(txt)
            offset = b
            if b <= i or a >= j:
                continue  # 本节点不在命中区间
            s = max(i, a) - a
            e = min(j, b) - a
            if not first_done:
                keep_tail = txt[e:] if b > j else ""
                merged = txt[:s] + new + keep_tail
                if b <= j and e == len(txt):
                    # 命中区间覆盖到本节点末尾，后续节点的尾部已在循环中处理
                    merged = txt[:s] + new
                n.text = merged
                if merged != merged.strip():
                    n.set("{%s}space" % XML_NS, "preserve")
                first_done = True
            else:
                n.text = txt[:s] + txt[e:]
                if n.text != n.text.strip():
                    n.set("{%s}space" % XML_NS, "preserve")
        replaced += 1
        # 重算
        texts = [n.text or "" for n in nodes]
        full = "".join(texts)
        pos = i + len(new)
    return replaced


def fail(code, msg) -> NoReturn:
    sys.stderr.write("错误：%s\n" % msg)
    sys.exit(code)


def note(msg):
    sys.stderr.write("%s\n" % msg)


def para_label(doc, pid):
    p = doc.paras[pid]
    style = doc.para_style(p)
    loc = doc.para_loc[pid]
    head = pid if loc == "正文" else "%s（%s）" % (pid, loc)
    return "%s%s" % (head, ("〔%s〕" % style) if style else "")


# ---------------------------------------------------------------- 子命令

def cmd_extract(args):
    doc = Docx(args.file)
    out = []
    t_done = set()
    assert doc.body is not None  # Docx.__init__ has validated this
    for child in list(doc.body):
        if child.tag == qn("p"):
            pid = _pid_of(doc, child)
            out.append("%s\t%s" % (para_label(doc, pid), para_text(child)))
        elif child.tag == qn("tbl"):
            tid = _tid_of(doc, child)
            info = doc.tables[tid]
            nrows = len(info["rows"])
            ncols = max((len(r) for r in info["rows"]), default=0)
            out.append("%s　表格开始（%d 行 × %d 列）" % (tid, nrows, ncols))
            for ri, row in enumerate(info["rows"], start=1):
                for ci, cell_pids in enumerate(row, start=1):
                    text = " ⏎ ".join(para_text(doc.paras[pid2]) for pid2 in cell_pids)
                    ids = "/".join(cell_pids)
                    out.append("%s R%02d C%02d %s\t%s" % (tid, ri, ci, ids, text))
            out.append("%s　表格结束" % tid)
            t_done.add(tid)
    if args.footers:
        for name in sorted(doc.entries):
            if re.match(r"word/(header|footer)\d*\.xml$", name):
                try:
                    rt = ET.fromstring(doc.entries[name].decode("utf-8"))
                except ET.ParseError:
                    continue
                txt = " | ".join(para_text(p) for p in rt.iter(qn("p")) if para_text(p).strip())
                if txt.strip():
                    out.append("【%s】 %s" % (name, txt))
    data = "\n".join(out) + "\n"
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(data)
        note("已写出提取文本：%s（%d 段，%d 表）" % (args.out, len(doc.paras), len(doc.tables)))
    else:
        sys.stdout.write(data)


def _pid_of(doc, el):
    for pid, cand in doc.paras.items():
        if cand is el:
            return pid
    return "P????"


def _tid_of(doc, el):
    for tid, info in doc.tables.items():
        if info["el"] is el:
            return tid
    return "T????"


NUMBERS_RE = re.compile(
    r"(¥|￥|人民币|万元|元整|％|%|‰|税率|征收率|工作日|日内|天內|天内|个月|"
    r"违约金|赔偿金|质保金|保证金|定金|订金|上限|限额|比例|万分之|千分之)"
)


def cmd_search(args):
    doc = Docx(args.file)
    if args.numbers:
        rx = NUMBERS_RE
    elif args.regex:
        rx = re.compile(args.pattern)
    else:
        rx = re.compile(re.escape(args.pattern))
    hits = 0
    for pid in sorted(doc.paras):
        text = para_text(doc.paras[pid])
        if rx.search(text):
            hits += 1
            print("%s\t%s" % (para_label(doc, pid), text))
    if hits == 0:
        note("无匹配。")
    else:
        note("共 %d 处匹配。" % hits)


def _require_in_place_or_out(args, doc):
    if getattr(args, "dry_run", False):
        return None
    if getattr(args, "out", None):
        return args.out
    if getattr(args, "in_place", False):
        return doc.path
    fail(1, "未指定输出：请加 --in-place（改写原件并自动备份）或 -o 新文件，或先用 --dry-run 预演。")


def _parse_pids(doc, specs):
    pids = []
    for s in specs:
        for part in s.split(","):
            part = part.strip()
            if not part:
                continue
            pid = part.upper()
            if not re.fullmatch(r"P\d{4}", pid) or pid not in doc.paras:
                fail(1, "无效段落编号：%s（请先 extract 获取当前编号）" % part)
            pids.append(pid)
    return pids


def cmd_replace(args):
    doc = Docx(args.file)
    if args.para:
        scope = _parse_pids(doc, [args.para])
    else:
        scope = sorted(doc.paras)
    plan = []
    for pid in scope:
        text = para_text(doc.paras[pid])
        n = text.count(args.old)
        if n:
            plan.append((pid, n))
    if not plan:
        fail(1, "未找到匹配文本：%r" % args.old)
    total = sum(n for _, n in plan)
    if len(plan) > 1 and not (args.all or args.para):
        sys.stderr.write("匹配分布于 %d 个段落（共 %d 处），请用 --para 指定或 --all 确认全部替换：\n" % (len(plan), total))
        for pid, n in plan:
            sys.stderr.write("  %s（%d 处）\t%s\n" % (para_label(doc, pid), n, para_text(doc.paras[pid])[:80]))
        sys.exit(1)
    for pid, n in plan:
        sys.stderr.write("%s\t%s\n" % (para_label(doc, pid), para_text(doc.paras[pid])[:120]))
    if args.dry_run:
        note("【预演】将替换 %d 个段落中的 %d 处；未写盘。" % (len(plan), total))
        return
    done = 0
    for pid, _ in plan:
        done += replace_in_para(doc.paras[pid], args.old, args.new)
    out = _require_in_place_or_out(args, doc)
    doc.save(out)
    note("完成替换 %d 处。" % done)


def cmd_set(args):
    doc = Docx(args.file)
    pids = _parse_pids(doc, [args.para])
    pid = pids[0]
    p = doc.paras[pid]
    old = para_text(p)
    sys.stderr.write("%s\n  原：%s\n  新：%s\n" % (para_label(doc, pid), old[:120], args.text[:120]))
    if args.dry_run:
        note("【预演】未写盘。")
        return
    set_para_text(p, args.text)
    out = _require_in_place_or_out(args, doc)
    doc.save(out)


def cmd_insert_after(args):
    doc = Docx(args.file)
    anchor_pid = _parse_pids(doc, [args.para])[0]
    clone_pid = _parse_pids(doc, [args.clone])[0] if args.clone else anchor_pid
    anchor = doc.paras[anchor_pid]
    ref = doc.paras[clone_pid]
    parent = doc.parent[anchor]
    siblings = list(parent)
    idx = siblings.index(anchor)
    sys.stderr.write("锚点 %s，格式克隆自 %s\n" % (para_label(doc, anchor_pid), para_label(doc, clone_pid)))
    if args.dry_run:
        for t in args.text:
            sys.stderr.write("  将插入：%s\n" % t[:120])
        note("【预演】未写盘。")
        return
    inserted = []
    for offset, text in enumerate(args.text, start=1):
        new_p = copy.deepcopy(ref)
        # 去掉段落唯一标识，避免重复
        for attr in list(new_p.attrib):
            if attr.endswith("}paraId") or attr.endswith("}textId"):
                del new_p.attrib[attr]
        set_para_text(new_p, text)
        parent.insert(idx + offset, new_p)
        inserted.append(new_p)
    out = _require_in_place_or_out(args, doc)
    doc.save(out)
    note("已插入 %d 段（新编号请重新 extract 获取）。" % len(inserted))


def cmd_delete(args):
    doc = Docx(args.file)
    pids = _parse_pids(doc, args.paras)
    for pid in pids:
        p = doc.paras[pid]
        parent = doc.parent[p]
        if parent.tag == qn("tc") and len(parent.findall(qn("p"))) <= 1:
            fail(1, "%s 是表格单元格内唯一段落，删除会使文档失效；请改用 set 清空文本。" % pid)
    for pid in pids:
        sys.stderr.write("将删除 %s\t%s\n" % (para_label(doc, pid), para_text(doc.paras[pid])[:120]))
    if args.dry_run:
        note("【预演】未写盘。")
        return
    for pid in pids:
        p = doc.paras[pid]
        doc.parent[p].remove(p)
    out = _require_in_place_or_out(args, doc)
    doc.save(out)
    note("已删除 %d 段。" % len(pids))


def cmd_add_row(args):
    doc = Docx(args.file)
    tid = args.table.upper()
    if tid not in doc.tables:
        fail(1, "无效表格编号：%s" % args.table)
    info = doc.tables[tid]
    rows = info["el"].findall(qn("tr"))
    if not rows:
        fail(1, "%s 没有可克隆的行。" % tid)
    if args.after:
        m = re.fullmatch(r"R?(\d+)", args.after.upper())
        if not m or not (1 <= int(m.group(1)) <= len(rows)):
            fail(1, "--after 行号无效：%s（共 %d 行）" % (args.after, len(rows)))
        ref_idx = int(m.group(1)) - 1
    else:
        ref_idx = len(rows) - 1
    ref_tr = rows[ref_idx]
    cells = args.cells.split("|") if args.cells else None
    ref_cells = ref_tr.findall(qn("tc"))
    if cells and len(cells) != len(ref_cells):
        fail(1, "--cells 提供了 %d 列，但被克隆行有 %d 列（用 | 分隔，空单元格留空）。"
               % (len(cells), len(ref_cells)))
    new_tr = copy.deepcopy(ref_tr)
    for ci, tc in enumerate(new_tr.findall(qn("tc"))):
        ps = tc.findall(qn("p"))
        if not ps:
            continue
        keep = ps[0]
        for extra in ps[1:]:
            tc.remove(extra)
        text = cells[ci] if cells else ""
        set_para_text(keep, text)
    if args.dry_run:
        sys.stderr.write("将在 %s 第 %d 行后插入新行：%s\n" % (tid, ref_idx + 1, args.cells or "（空白行）"))
        note("【预演】未写盘。")
        return
    # ElementTree 无 addnext，用父表 insert
    tbl = info["el"]
    tbl.insert(list(tbl).index(ref_tr) + 1, new_tr)
    out = _require_in_place_or_out(args, doc)
    doc.save(out)
    note("已在 %s 增加一行。" % tid)


def cmd_batch(args):
    with open(args.edits, "r", encoding="utf-8") as f:
        ops = json.load(f)
    if not isinstance(ops, list):
        fail(1, "edits 文件须为 JSON 数组。")
    doc = Docx(args.file)
    # 预解析全部引用：所有 P/T 编号均以批处理开始前的文档状态为准
    resolved = {}
    for k, op in enumerate(ops):
        kind = op.get("op")
        for key in ("para", "clone", "anchor"):
            if key in op:
                pid = str(op[key]).upper()
                if pid not in doc.paras:
                    fail(1, "第 %d 项操作引用无效段落：%s" % (k + 1, op[key]))
                resolved[(k, key)] = doc.paras[pid]
        if kind == "delete":
            plist = op.get("paras") or [op.get("para")]
            for p in plist:
                pid = str(p).upper()
                if pid not in doc.paras:
                    fail(1, "第 %d 项操作引用无效段落：%s" % (k + 1, p))
                resolved[(k, "del", pid)] = doc.paras[pid]
    deleted = set()
    applied = 0
    for k, op in enumerate(ops):
        kind = op.get("op")
        try:
            if kind == "replace":
                scope = []
                if (k, "para") in resolved:
                    scope = [resolved[(k, "para")]]
                else:
                    scope = [doc.paras[pid] for pid in sorted(doc.paras)]
                n = 0
                hits = 0
                for p in scope:
                    c = para_text(p).count(op["old"])
                    if c:
                        hits += 1
                        if not args.dry_run:
                            n += replace_in_para(p, op["old"], op["new"])
                        else:
                            n += c
                if hits == 0:
                    fail(1, "第 %d 项 replace 未找到：%r" % (k + 1, op["old"]))
                if hits > 1 and not op.get("all") and (k, "para") not in resolved:
                    fail(1, "第 %d 项 replace 匹配 %d 段，需 \"para\" 指定或 \"all\": true。" % (k + 1, hits))
                applied += 1
                sys.stderr.write("[%d] replace %d 处（%d 段）\n" % (k + 1, n, hits))
            elif kind == "set":
                p = resolved[(k, "para")]
                if id(p) in deleted:
                    fail(1, "第 %d 项 set 的目标已被前面操作删除。" % (k + 1))
                if not args.dry_run:
                    set_para_text(p, op["text"])
                applied += 1
                sys.stderr.write("[%d] set %s\n" % (k + 1, op.get("para")))
            elif kind == "insert-after":
                anchor = resolved[(k, "para")]
                if id(anchor) in deleted:
                    fail(1, "第 %d 项 insert-after 的锚点已被删除。" % (k + 1))
                ref = resolved.get((k, "clone"), anchor)
                parent = doc.parent[anchor]
                idx = list(parent).index(anchor)
                texts = op["text"] if isinstance(op["text"], list) else [op["text"]]
                if not args.dry_run:
                    for offset, text in enumerate(texts, start=1):
                        new_p = copy.deepcopy(ref)
                        for attr in list(new_p.attrib):
                            if attr.endswith("}paraId") or attr.endswith("}textId"):
                                del new_p.attrib[attr]
                        set_para_text(new_p, text)
                        parent.insert(idx + offset, new_p)
                applied += 1
                sys.stderr.write("[%d] insert-after %s × %d 段\n" % (k + 1, op.get("para"), len(texts)))
            elif kind == "delete":
                plist = op.get("paras") or [op.get("para")]
                for p in plist:
                    pid = str(p).upper()
                    el = resolved[(k, "del", pid)]
                    parent = doc.parent[el]
                    if parent.tag == qn("tc") and len(parent.findall(qn("p"))) <= 1:
                        fail(1, "第 %d 项 delete：%s 是单元格内唯一段落，不能删除。" % (k + 1, pid))
                    if not args.dry_run:
                        parent.remove(el)
                    deleted.add(id(el))
                applied += 1
                sys.stderr.write("[%d] delete × %d\n" % (k + 1, len(plist)))
            elif kind == "add-row":
                tid = str(op["table"]).upper()
                if tid not in doc.tables:
                    fail(1, "第 %d 项 add-row 引用无效表格：%s" % (k + 1, op["table"]))
                info = doc.tables[tid]
                rows = info["el"].findall(qn("tr"))
                ref_tr = rows[-1]
                if op.get("after"):
                    m = re.fullmatch(r"R?(\d+)", str(op["after"]).upper())
                    if not m or not (1 <= int(m.group(1)) <= len(rows)):
                        fail(1, "第 %d 项 add-row 的 after 行号无效：%s。" % (k + 1, op["after"]))
                    ref_tr = rows[int(m.group(1)) - 1]
                cells = op.get("cells", "")
                cell_list = cells.split("|") if cells else None
                if not args.dry_run:
                    new_tr = copy.deepcopy(ref_tr)
                    ref_cells = ref_tr.findall(qn("tc"))
                    if cell_list and len(cell_list) != len(ref_cells):
                        fail(1, "第 %d 项 add-row 列数不符：%d ≠ %d。" % (k + 1, len(cell_list), len(ref_cells)))
                    for ci, tc in enumerate(new_tr.findall(qn("tc"))):
                        ps = tc.findall(qn("p"))
                        if not ps:
                            continue
                        for extra in ps[1:]:
                            tc.remove(extra)
                        set_para_text(ps[0], cell_list[ci] if cell_list else "")
                    tbl = info["el"]
                    tbl.insert(list(tbl).index(ref_tr) + 1, new_tr)
                applied += 1
                sys.stderr.write("[%d] add-row %s\n" % (k + 1, tid))
            else:
                fail(1, "第 %d 项操作类型未知：%r" % (k + 1, kind))
        except KeyError as e:
            fail(1, "第 %d 项操作缺少字段：%s" % (k + 1, e))
    if args.dry_run:
        note("【预演】共 %d 项操作可执行；未写盘。" % applied)
        return
    out = _require_in_place_or_out(args, doc)
    doc.save(out)
    note("批量执行完成：%d 项。" % applied)


def cmd_verify(args):
    doc = Docx(args.file)  # 装载过程已做 zip 与 XML 校验
    with zipfile.ZipFile(args.file, "r") as z:
        bad = z.testzip()
        if bad:
            fail(2, "zip 条目损坏：%s" % bad)
    print("结构正常：%d 个段落，%d 张表格，%d 个 zip 条目。" % (len(doc.paras), len(doc.tables), len(doc.entries)))
    empties = [pid for pid in sorted(doc.paras) if not para_text(doc.paras[pid]).strip()]
    if empties:
        print("空段落 %d 个（多为版式空行）：%s%s" % (
            len(empties), " ".join(empties[:10]), " …" if len(empties) > 10 else ""))


# ---------------------------------------------------------------- 入口

def build_parser():
    ap = argparse.ArgumentParser(
        prog="docx_tool.py",
        description="纯标准库 docx 段落级读写工具（合同起草/修订/审核辅助）",
    )
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common_edit(p):
        p.add_argument("--in-place", action="store_true", help="改写原文件（自动备份为 .bak.时间戳）")
        p.add_argument("-o", "--out", help="输出到新文件")
        p.add_argument("--dry-run", action="store_true", help="只打印将发生的修改，不写盘")

    p = sub.add_parser("extract", help="提取全文（带段落/表格编号）")
    p.add_argument("file")
    p.add_argument("-o", "--out", help="写出到文本文件")
    p.add_argument("--footers", action="store_true", help="同时提取页眉页脚文本")
    p.set_defaults(fn=cmd_extract)

    p = sub.add_parser("search", help="定位包含指定文本的段落")
    p.add_argument("file")
    p.add_argument("pattern", nargs="?", default="")
    p.add_argument("-e", "--regex", action="store_true", help="pattern 按正则解释")
    p.add_argument("--numbers", action="store_true", help="预设：金额/税率/期限/违约金等数字敏感段落")
    p.set_defaults(fn=cmd_search)

    p = sub.add_parser("replace", help="跨 run 替换文本（保留格式）")
    p.add_argument("file")
    p.add_argument("old")
    p.add_argument("new")
    p.add_argument("--para", help="限定段落编号（如 P0012）")
    p.add_argument("--all", action="store_true", help="允许替换全部匹配段")
    common_edit(p)
    p.set_defaults(fn=cmd_replace)

    p = sub.add_parser("set", help="整段改写（保留段落与字符格式）")
    p.add_argument("file")
    p.add_argument("para", help="段落编号 Pxxxx")
    p.add_argument("--text", required=True)
    common_edit(p)
    p.set_defaults(fn=cmd_set)

    p = sub.add_parser("insert-after", help="在指定段后克隆插入新段（继承格式）")
    p.add_argument("file")
    p.add_argument("para", help="锚点段落编号 Pxxxx")
    p.add_argument("--text", action="append", required=True, help="新段文本；可重复多次")
    p.add_argument("--clone", help="格式来源段落编号（默认同锚点）")
    common_edit(p)
    p.set_defaults(fn=cmd_insert_after)

    p = sub.add_parser("delete", help="删除段落")
    p.add_argument("file")
    p.add_argument("paras", nargs="+", help="段落编号，可多个或逗号分隔")
    common_edit(p)
    p.set_defaults(fn=cmd_delete)

    p = sub.add_parser("add-row", help="在表格中克隆新增一行")
    p.add_argument("file")
    p.add_argument("table", help="表格编号 Txxxx")
    p.add_argument("--cells", help="各单元格文本，用 | 分隔；缺省则插入空白行")
    p.add_argument("--after", help="插到第 N 行后（如 R02 或 2），默认末行后")
    common_edit(p)
    p.set_defaults(fn=cmd_add_row)

    p = sub.add_parser("batch", help="按 edits.json 批量执行（编号以执行前状态为准）")
    p.add_argument("file")
    p.add_argument("edits", help="JSON 文件：replace/set/insert-after/delete/add-row 操作数组")
    common_edit(p)
    p.set_defaults(fn=cmd_batch)

    p = sub.add_parser("verify", help="校验 docx 结构完整性")
    p.add_argument("file")
    p.set_defaults(fn=cmd_verify)

    return ap


def main(argv=None):
    args = build_parser().parse_args(argv)
    if getattr(args, "cmd", None) == "search" and not args.numbers and not args.pattern:
        fail(1, "search 需要 pattern 或 --numbers。")
    args.fn(args)


if __name__ == "__main__":
    main()
