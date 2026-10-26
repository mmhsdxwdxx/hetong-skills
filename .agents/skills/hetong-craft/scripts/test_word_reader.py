#!/usr/bin/env python3
"""word_reader.py 的自动回归测试。

测试现场构造最小但真实的 OOXML 文件（含修订、传统/现代批注、移动、格式变化、
表格行变化和隐藏文字），不会读取或修改任何真实合同。
"""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SCRIPT = Path(__file__).resolve().parent / "word_reader.py"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>
<Override PartName="/word/commentsExtensible.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"/>
<Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="HiddenStyle"><w:name w:val="HiddenStyle"/><w:basedOn w:val="Normal"/><w:rPr><w:vanish/></w:rPr></w:style>
</w:styles>"""

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:trackChanges/>
</w:settings>"""

CORE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>虚构多轮审阅合同</dc:title><dc:creator>Alice</dc:creator>
<cp:lastModifiedBy>Dave</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T08:00:00Z</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-11T18:00:00Z</dcterms:modified>
</cp:coreProperties>"""


def run(text: str) -> str:
    return f"<w:r><w:t>{text}</w:t></w:r>"


def ins(rev_id: int, author: str, date: str | None, text: str) -> str:
    date_attr = f' w:date="{date}"' if date else ""
    return f'<w:ins w:id="{rev_id}" w:author="{author}"{date_attr}>{run(text)}</w:ins>'


def delete(rev_id: int, author: str, date: str, text: str) -> str:
    return f'<w:del w:id="{rev_id}" w:author="{author}" w:date="{date}"><w:r><w:delText>{text}</w:delText></w:r></w:del>'


def cstart(cid: int) -> str:
    return f'<w:commentRangeStart w:id="{cid}"/>'


def cend(cid: int) -> str:
    return f'<w:commentRangeEnd w:id="{cid}"/><w:r><w:commentReference w:id="{cid}"/></w:r>'


def cell(text: str, cid: int | None = None) -> str:
    inner = run(text) if cid is None else cstart(cid) + run(text) + cend(cid)
    return f'<w:tc><w:p>{inner}</w:p></w:tc>'


def document_xml(include_revisions: bool = True) -> str:
    if not include_revisions:
        body = (
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + run("第一条 价款") + "</w:p>"
            + "<w:p>" + run("合同价款为人民币100,000元。") + "</w:p>"
        )
    else:
        p1 = (
            "<w:p>" + cstart(0) + run("付款比例为")
            + ins(1, "Alice", "2026-01-02T09:00:00Z", "含税")
            + delete(2, "Bob", "2026-01-03T10:00:00Z", "30%")
            + ins(3, "Bob", "2026-01-03T10:01:00Z", "40%")
            + run("，验收期限为")
            + delete(4, "Alice", "2026-01-02T09:05:00Z", "5")
            + ins(5, "Alice", "2026-01-02T09:06:00Z", "7")
            + run("个工作日。") + cend(0) + "</w:p>"
        )
        p2 = (
            "<w:p>" + run("含税总价为") + cstart(1)
            + delete(6, "Bob", "2026-01-06T11:00:00Z", "123,600")
            + cend(1) + ins(7, "Bob", "2026-01-06T11:01:00Z", "130,000") + run("元。")
            + '<w:r><w:commentReference w:id="6"/></w:r>'
            + "</w:p>"
        )
        p3 = (
            '<w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/>'
            '<w:rPrChange w:id="8" w:author="Dave" w:date="2026-01-09T09:00:00Z">'
            '<w:rPr><w:sz w:val="24"/></w:rPr></w:rPrChange></w:rPr>'
            "<w:t>违约金为合同总价的10%。</w:t></w:r></w:p>"
        )
        hidden = (
            '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>内部说明：不得外传。</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:pStyle w:val="HiddenStyle"/></w:pPr><w:r><w:t>样式隐藏的底价信息。</w:t></w:r></w:p>'
        )
        moves = (
            "<w:p>" + run("争议解决：")
            + '<w:moveFrom w:id="30" w:name="M2" w:author="Carol" w:date="2026-01-08T09:00:00Z">'
            + run("提交甲方所在地法院。") + "</w:moveFrom></w:p>"
            "<w:p>" + run("修改后：")
            + '<w:moveTo w:id="31" w:name="M2" w:author="Carol" w:date="2026-01-08T09:01:00Z">'
            + run("提交甲方所在地法院。") + "</w:moveTo></w:p>"
        )
        marks = (
            '<w:p><w:pPr><w:rPr><w:ins w:id="9" w:author="Alice" w:date="2026-01-10T09:00:00Z"/></w:rPr></w:pPr>'
            + run("新增段落原稿无此分段。") + "</w:p>"
            '<w:p><w:pPr><w:rPr><w:del w:id="10" w:author="Alice" w:date="2026-01-10T09:05:00Z"/></w:rPr></w:pPr>'
            + run("本段接受后与下一段合并。") + "</w:p><w:p>" + run("下一段。") + "</w:p>"
        )
        extra = (
            "<w:p>" + ins(99, "FutureUser", "2030-01-01T00:00:00Z", "未来日期修订。") + "</w:p>"
            "<w:p>" + ins(98, "NoDate", None, "无日期修订。") + "</w:p>"
        )
        row1 = f"<w:tr>{cell('项目')}{cell('结果')}</w:tr>"
        row2 = (
            '<w:tr><w:trPr><w:ins w:id="20" w:author="Dave" w:date="2026-01-09T10:00:00Z"/></w:trPr>'
            + cell("新增行", 4) + cell("已插入") + "</w:tr>"
        )
        row3 = (
            '<w:tr><w:trPr><w:del w:id="21" w:author="Dave" w:date="2026-01-09T10:01:00Z"/></w:trPr>'
            + cell("删除行") + cell("已删除") + "</w:tr>"
        )
        table = (
            '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            + row1 + row2 + row3 + "</w:tbl>"
        )
        body = (
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + run("第一条 付款与验收") + "</w:p>"
            + p1 + p2 + p3 + hidden + moves
            + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + run("第二条 段落变更") + "</w:p>"
            + marks
            + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + run("第三条 表格变更") + "</w:p>"
            + table
            + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + run("第四条 异常时间线") + "</w:p>"
            + extra
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
        'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<w:body>{body}<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr></w:body></w:document>"
    )


def comments_xml() -> str:
    def item(cid: int, author: str, date: str, text: str, para_id: str) -> str:
        return (
            f'<w:comment w:id="{cid}" w:author="{author}" w:initials="{author[:1]}" w:date="{date}">'
            f'<w:p w14:paraId="{para_id}"><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        + item(0, "Alice", "2026-01-06T10:10:00Z", "请确认比例与验收期限。", "11111111")
        + item(1, "Bob", "2026-01-07T10:20:00Z", "旧金额为何删除？", "33333333")
        + item(2, "Bob", "2026-01-07T10:30:00Z", "已按40%和7个工作日调整。", "22222222")
        + item(4, "Dave", "2026-01-09T10:10:00Z", "新增行需确认保证金比例。", "44444444")
        + item(5, "Carol", "2026-01-09T10:20:00Z", "这条批注没有正文锚点。", "55555555")
        + "</w:comments>"
    )


def comments_extended() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
        '<w15:commentEx w15:paraId="11111111" w15:done="0"/>'
        '<w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="1"/>'
        '<w15:commentEx w15:paraId="33333333" w15:done="0"/>'
        '<w15:commentEx w15:paraId="44444444" w15:done="1"/>'
        '<w15:commentEx w15:paraId="55555555" w15:done="0"/>'
        "</w15:commentsEx>"
    )


def comments_ids() -> str:
    rows = "".join(
        f'<w16cid:commentId w16cid:paraId="{para}" w16cid:durableId="{durable}"/>'
        for para, durable in (("11111111", "9001"), ("22222222", "9002"), ("33333333", "9003"),
                              ("44444444", "9004"), ("55555555", "9005"))
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">'
        + rows + "</w16cid:commentsIds>"
    )


def comments_extensible() -> str:
    rows = "".join(
        f'<w16cex:commentExtensible w16cex:durableId="{durable}" w16cex:dateUtc="{date}"/>'
        for durable, date in (("9001", "2026-01-06T10:15:00Z"), ("9002", "2026-01-07T10:31:00Z"),
                              ("9003", "2026-01-07T10:21:00Z"), ("9004", "2026-01-09T10:11:00Z"),
                              ("9005", "2026-01-09T10:21:00Z"))
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">'
        + rows + "</w16cex:commentsExtensible>"
    )


def people_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
        '<w15:person w15:author="Alice"/><w15:person w15:author="Bob"/>'
        '<w15:person w15:author="Carol"/><w15:person w15:author="Dave"/>'
        "</w15:people>"
    )


def net_zero_xml() -> str:
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + run("第一条 回改") + "</w:p>"
        + "<w:p>" + run("保证金为") + delete(50, "Alice", "2026-01-12T09:00:00Z", "10%")
        + ins(51, "Bob", "2026-01-12T09:05:00Z", "10%") + run("。") + "</w:p>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + '<w:body>' + body + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>'
    )


def make_net_zero(path: Path, source: Path) -> None:
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                data = net_zero_xml().encode()
            dst.writestr(info, data)


def make_docx(path: Path, include_revisions: bool = True) -> None:
    files = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": RELS,
        "word/document.xml": document_xml(include_revisions),
        "word/styles.xml": STYLES,
        "word/settings.xml": SETTINGS,
        "word/_rels/document.xml.rels": DOC_RELS,
        "docProps/core.xml": CORE,
    }
    if include_revisions:
        files.update({
            "word/comments.xml": comments_xml(),
            "word/commentsExtended.xml": comments_extended(),
            "word/commentsIds.xml": comments_ids(),
            "word/commentsExtensible.xml": comments_extensible(),
            "word/people.xml": people_xml(),
        })
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_cli(*args: str, expect: int = 0) -> dict:
    completed = subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        capture_output=True, text=True, encoding="utf-8", check=False,
    )
    if completed.returncode != expect:
        raise AssertionError(
            f"exit={completed.returncode}, expected={expect}\n"
            f"stdout:\n{completed.stdout[-4000:]}\nstderr:\n{completed.stderr[-4000:]}"
        )
    try:
        return json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise AssertionError(f"输出不是 JSON：{completed.stdout[-2000:]}") from exc


class WordReaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temporary = tempfile.TemporaryDirectory(prefix="word-reader-test-")
        cls.root = Path(cls.temporary.name)
        cls.docx = cls.root / "虚构审阅合同.docx"
        cls.clean = cls.root / "虚构清洁合同.docx"
        make_docx(cls.docx)
        make_docx(cls.clean, include_revisions=False)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temporary.cleanup()

    def analyze(self, source: Path, name: str, *extra: str) -> tuple[dict, Path]:
        output = self.root / name
        summary = run_cli("analyze", str(source), "--out", str(output), *extra)
        self.assertTrue(summary["ok"])
        data = json.loads((output / "deep-read.json").read_text(encoding="utf-8"))
        return data, output

    def test_views_comments_timeline_and_clauses(self) -> None:
        before = sha256(self.docx)
        data, output = self.analyze(self.docx, "full")
        self.assertEqual(data["file"]["sha256"], before)
        self.assertEqual(before, sha256(self.docx))
        self.assertEqual(data["counts"]["revisions"], 16)
        self.assertEqual(data["counts"]["comments"], 6)
        self.assertEqual(data["counts"]["comment_threads"], 5)
        self.assertEqual(data["counts"]["unresolved_or_unknown_comments"], 4)

        accepted = (output / "accepted.md").read_text(encoding="utf-8")
        original = (output / "original.md").read_text(encoding="utf-8")
        markup = (output / "markup.md").read_text(encoding="utf-8")
        report = (output / "deep-read.md").read_text(encoding="utf-8")
        self.assertIn("[WR0002 · 正文] 付款比例为含税40%", accepted)
        self.assertNotIn("[WR0001 · 正文] 第一条 付款与验收", accepted)
        self.assertIn("批注涉及：请确认比例与验收期限", report)
        self.assertIn("付款比例为30%，验收期限为5个工作日。", original)
        self.assertNotIn("付款比例为30%", accepted)
        self.assertIn("123,600", original)
        self.assertNotIn("123,600", accepted)
        self.assertIn("130,000", accepted)
        self.assertIn("⟦插入", markup)
        self.assertIn("⟦删除", markup)
        self.assertIn("〔批注：C", markup)
        self.assertNotIn("内部说明", accepted)
        self.assertNotIn("样式隐藏的底价信息", original)
        self.assertIn("内部说明", markup)
        self.assertIn("本段接受后与下一段合并。下一段。", accepted)
        self.assertIn("新增段落原稿无此分段。", original)
        # 插入行只应出现在接受视图，删除行只应出现在原文视图。
        self.assertIn("新增行", accepted)
        self.assertNotIn("新增行", original)
        self.assertIn("删除行", original)
        self.assertNotIn("删除行", accepted)

        clauses = {item["clause_path"]: item for item in data["clauses"]}
        self.assertEqual(len(clauses), 4)
        self.assertTrue(all(item["changed"] for item in clauses.values()))
        payment = next(item for item in data["clauses"] if "第一条" in item["clause_path"])
        self.assertIn("C0001", payment["comment_ids"])
        self.assertTrue(payment["signals"])

        comments = {item["code"]: item for item in data["comments"]}
        self.assertEqual(comments["C0003"]["parent_code"], "C0001")
        self.assertEqual(comments["C0003"]["thread_status"], "open")
        self.assertTrue(comments["C0004"]["resolved"])
        self.assertTrue(comments["C0002"]["anchor_deleted"])
        self.assertTrue(comments["C0005"]["orphan"])
        self.assertEqual(comments["C0006"]["orphan_anchor"], True)

        timeline = (output / "timeline.csv").read_text(encoding="utf-8-sig")
        self.assertIn("EV0001", timeline)
        self.assertIn("C0001", timeline)
        self.assertIn("moveFrom", timeline)
        self.assertIn("moveTo", timeline)
        self.assertTrue(all((output / name).is_file() for name in (
            "deep-read.json", "deep-read.md", "accepted.md", "original.md",
            "markup.md", "timeline.csv", "comments.csv",
        )))

        types = data["counts"]["revision_types"]
        self.assertEqual(types["moveFrom"], 1)
        self.assertEqual(types["moveTo"], 1)
        self.assertEqual(types["runFormatting"], 1)
        self.assertEqual(types["tableRowInserted"], 1)
        self.assertEqual(types["tableRowDeleted"], 1)
        self.assertGreaterEqual(types["paragraphMarkInserted"], 1)
        self.assertGreaterEqual(types["paragraphMarkDeleted"], 1)
        self.assertTrue(data["views"]["accepted"])
        self.assertTrue(data["views"]["original"])

        anomalies = {item["kind"] for item in data["anomalies"]}
        self.assertIn("comment_anchor_deleted", anomalies)
        self.assertIn("orphan_comment_anchor", anomalies)
        self.assertIn("revision_date_missing", anomalies)
        self.assertIn("revision_future", anomalies)
        self.assertIn("comment_date_missing", anomalies)
        self.assertIn("hidden_text", anomalies)
        self.assertIn("comment_status_not_resolved", anomalies)
        self.assertNotIn("package_protection", anomalies)

    def test_net_zero_change_is_visible_as_reverted(self) -> None:
        net_zero = self.root / "net-zero.docx"
        make_net_zero(net_zero, self.clean)
        data, output = self.analyze(net_zero, "net-zero")
        self.assertEqual(data["counts"]["changed_clauses"], 0)
        self.assertEqual(data["counts"]["net_zero_clauses"], 1)
        self.assertIn("net_zero_changes", {item["kind"] for item in data["anomalies"]})
        self.assertEqual(
            (output / "accepted.md").read_text(encoding="utf-8").count("保证金为10%。"),
            (output / "original.md").read_text(encoding="utf-8").count("保证金为10%。"),
        )

    def test_include_hidden_and_clean_document(self) -> None:
        data, output = self.analyze(self.docx, "hidden", "--include-hidden")
        accepted = (output / "accepted.md").read_text(encoding="utf-8")
        self.assertIn("内部说明", accepted)
        self.assertIn("样式隐藏的底价信息", accepted)
        clean, clean_output = self.analyze(self.clean, "clean")
        self.assertEqual(clean["counts"]["revisions"], 0)
        self.assertEqual(clean["counts"]["comments"], 0)
        self.assertIn("合同价款为人民币100,000元。", (clean_output / "accepted.md").read_text())

    def test_force_protection_and_output_immutability(self) -> None:
        output = self.root / "force"
        first = run_cli("analyze", str(self.docx), "--out", str(output))
        self.assertTrue(first["ok"])
        second = run_cli("analyze", str(self.docx), "--out", str(output), expect=2)
        self.assertIn("已存在", second["error"])
        third = run_cli("analyze", str(self.docx), "--out", str(output), "--force")
        self.assertTrue(third["ok"])

    def test_macro_package_is_read_but_flagged(self) -> None:
        macro = self.root / "macro.docx"
        with zipfile.ZipFile(self.docx) as source, zipfile.ZipFile(macro, "w") as target:
            for info in source.infolist():
                target.writestr(info, source.read(info.filename))
            target.writestr("word/vbaProject.bin", b"not-executed")
        data, _ = self.analyze(macro, "macro")
        kinds = {item["kind"] for item in data["anomalies"]}
        self.assertIn("package_protection", kinds)
        self.assertTrue(any("宏" in item["message"] for item in data["anomalies"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
