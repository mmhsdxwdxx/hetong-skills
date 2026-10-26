#!/usr/bin/env python3
"""hidden_change.py 的回归测试，全部使用现场构造的虚构 DOCX。"""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SCRIPT = Path(__file__).resolve().parent / "hidden_change.py"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:trackChanges/>
</w:settings>"""

COMMENTS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:comment w:id="0" w:author="Alice" w:date="2026-01-01T00:00:00Z">
<w:p><w:r><w:t>请确认付款比例。</w:t></w:r></w:p>
</w:comment></w:comments>"""


def paragraph(text: str) -> str:
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def make_docx(path: Path, body_xml: str, *, settings: str = SETTINGS,
              media: dict[str, bytes] | None = None, comments: bool = False) -> None:
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{body_xml}<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr></w:body></w:document>"
    )
    files: dict[str, str | bytes] = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": RELS,
        "word/document.xml": document,
        "word/styles.xml": STYLES,
        "word/settings.xml": settings,
        "word/_rels/document.xml.rels": DOC_RELS,
    }
    for name, data in (media or {}).items():
        files[name] = data
    if comments:
        files["word/comments.xml"] = COMMENTS
        files["word/_rels/document.xml.rels"] = DOC_RELS.replace(
            "</Relationships>",
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'
        )
        files["[Content_Types].xml"] = CONTENT_TYPES.replace(
            "</Types>",
            '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'
        )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_cli(*args: str, expect: int = 0) -> dict:
    completed = subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        capture_output=True, text=True, check=False,
    )
    if completed.returncode != expect:
        raise AssertionError(
            f"exit={completed.returncode}, expected={expect}\n"
            f"stdout:\n{completed.stdout[-5000:]}\nstderr:\n{completed.stderr[-5000:]}"
        )
    try:
        return json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise AssertionError(f"输出不是 JSON：{completed.stdout[-2000:]}") from exc


class HiddenChangeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temporary = tempfile.TemporaryDirectory(prefix="hidden-change-test-")
        cls.root = Path(cls.temporary.name)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temporary.cleanup()

    def pair(self, name: str, before_body: str, after_body: str, **after_options) -> tuple[Path, Path]:
        before = self.root / f"{name}-before.docx"
        after = self.root / f"{name}-after.docx"
        make_docx(before, before_body)
        make_docx(after, after_body, **after_options)
        return before, after

    def test_tracked_change_is_not_reported_as_untracked(self) -> None:
        before, after = self.pair(
            "tracked",
            paragraph("付款比例为30%。"),
            '<w:p><w:r><w:t>付款比例为</w:t></w:r>'
            '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>30%</w:delText></w:r></w:del>'
            '<w:ins w:id="2" w:author="A" w:date="2026-01-01T00:01:00Z"><w:r><w:t>40%</w:t></w:r></w:ins>'
            '<w:r><w:t>。</w:t></w:r></w:p>',
        )
        before_hash, after_hash = sha(before), sha(after)
        output = self.root / "tracked-out"
        result = run_cli("compare", "--before", str(before), "--after", str(after), "--out", str(output))
        self.assertTrue(result["ok"])
        self.assertEqual(result["counts"]["untracked_text_changes"], 0)
        self.assertGreaterEqual(result["counts"]["tracked_text_changes"], 1)
        self.assertEqual(before_hash, sha(before))
        self.assertEqual(after_hash, sha(after))
        self.assertTrue((output / "hidden-change-report.md").is_file())

    def test_direct_text_change_is_untracked(self) -> None:
        before, after = self.pair("direct", paragraph("付款比例为30%。"), paragraph("付款比例为40%。"))
        result = run_cli("compare", "--before", str(before), "--after", str(after),
                         "--out", str(self.root / "direct-out"), expect=1)
        self.assertEqual(result["counts"]["untracked_text_changes"], 1)
        self.assertEqual(result["severity"], "高")
        findings = json.loads((self.root / "direct-out" / "hidden-change.json").read_text())
        self.assertTrue(any(item["category"] == "untracked_text_change" for item in findings["findings"]))
        self.assertIn("未留痕", (self.root / "direct-out" / "hidden-change-report.md").read_text())

    def test_mixed_tracked_and_direct_changes(self) -> None:
        before, after = self.pair(
            "mixed",
            paragraph("付款比例为30%，验收期限为5日。"),
            '<w:p><w:r><w:t>付款比例为</w:t></w:r>'
            '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>30%</w:delText></w:r></w:del>'
            '<w:ins w:id="2" w:author="A" w:date="2026-01-01T00:01:00Z"><w:r><w:t>40%</w:t></w:r></w:ins>'
            '<w:r><w:t>，验收期限为7日。</w:t></w:r></w:p>',
        )
        result = run_cli("compare", "--before", str(before), "--after", str(after),
                         "--out", str(self.root / "mixed-out"), expect=1)
        self.assertEqual(result["counts"]["untracked_text_changes"], 1)
        self.assertGreaterEqual(result["counts"]["tracked_text_changes"], 1)
        self.assertEqual(result["severity"], "高")

    def test_net_zero_history_removal_is_flagged(self) -> None:
        before_body = (
            '<w:p><w:r><w:t>保证金为</w:t></w:r>'
            '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>10%</w:delText></w:r></w:del>'
            '<w:ins w:id="2" w:author="A" w:date="2026-01-01T00:01:00Z"><w:r><w:t>10%</w:t></w:r></w:ins>'
            '<w:r><w:t>。</w:t></w:r></w:p>'
        )
        after_body = paragraph("保证金为10%。")
        before, after = self.pair("history", before_body, after_body)
        result = run_cli("compare", "--before", str(before), "--after", str(after),
                         "--out", str(self.root / "history-out"))
        self.assertEqual(result["counts"]["untracked_text_changes"], 0)
        self.assertEqual(result["counts"]["history_removed"], 2)
        findings = json.loads((self.root / "history-out" / "hidden-change.json").read_text())
        self.assertTrue(any(item["category"] == "review_history_removed" for item in findings["findings"]))

    def test_hidden_text_and_media_changes(self) -> None:
        before_body = (
            '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>内部底价A</w:t></w:r>'
            '<w:r><w:t>合同价款为100,000元。</w:t></w:r></w:p>'
        )
        after_body = (
            '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>内部底价B</w:t></w:r>'
            '<w:r><w:t>合同价款为100,000元。</w:t></w:r></w:p>'
        )
        before, after = self.pair(
            "hidden-media", before_body, after_body,
        )
        # 追加不同的媒体文件，模拟未留痕替换
        with zipfile.ZipFile(before, "a") as archive:
            archive.writestr("word/media/image1.png", b"image-A")
        with zipfile.ZipFile(after, "a") as archive:
            archive.writestr("word/media/image1.png", b"image-B")
        result = run_cli("compare", "--before", str(before), "--after", str(after),
                         "--out", str(self.root / "hidden-media-out"))
        self.assertEqual(result["counts"]["untracked_text_changes"], 0)
        self.assertEqual(result["counts"]["hidden_text_changes"], 1)
        self.assertEqual(result["counts"]["media_changes"], 1)
        findings = json.loads((self.root / "hidden-media-out" / "hidden-change.json").read_text())
        categories = {item["category"] for item in findings["findings"]}
        self.assertIn("hidden_text_change", categories)
        self.assertIn("non_text_media_change", categories)

    def test_low_similarity_warning(self) -> None:
        before, after = self.pair("low-sim", paragraph("完全不同的甲合同内容。"), paragraph("无关的乙合同内容，另有条款。"))
        _result = run_cli("compare", "--before", str(before), "--after", str(after),
                          "--out", str(self.root / "low-sim-out"), expect=1)
        findings = json.loads((self.root / "low-sim-out" / "hidden-change.json").read_text())
        self.assertTrue(any(item["category"] == "low_similarity" for item in findings["findings"]))

    def test_comments_removed_is_flagged(self) -> None:
        before_body = (
            '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>付款比例为30%。</w:t></w:r>'
            '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>'
        )
        before = self.root / "comments-before.docx"
        after = self.root / "comments-after.docx"
        make_docx(before, before_body, comments=True)
        make_docx(after, paragraph("付款比例为30%。"))
        result = run_cli("compare", "--before", str(before), "--after", str(after),
                         "--out", str(self.root / "comments-out"))
        self.assertEqual(result["counts"]["untracked_text_changes"], 0)
        self.assertEqual(result["counts"]["comments_removed"], 1)
        findings = json.loads((self.root / "comments-out" / "hidden-change.json").read_text())
        self.assertTrue(any(item["category"] == "comments_removed" for item in findings["findings"]))

    def test_force_protection(self) -> None:
        before, after = self.pair("force", paragraph("甲。"), paragraph("乙。"))
        output = self.root / "force-out"
        run_cli("compare", "--before", str(before), "--after", str(after), "--out", str(output), expect=1)
        second = run_cli("compare", "--before", str(before), "--after", str(after),
                         "--out", str(output), expect=2)
        self.assertIn("已存在", second["error"])
        third = run_cli("compare", "--before", str(before), "--after", str(after),
                        "--out", str(output), "--force", expect=1)
        self.assertTrue(third["ok"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
