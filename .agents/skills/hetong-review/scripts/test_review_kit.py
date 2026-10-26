#!/usr/bin/env python3
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "review_kit.py"


def run_cli(*args, expected=0):
    result = subprocess.run(
        [sys.executable, str(SCRIPT), *map(str, args)],
        capture_output=True,
        text=True,
    )
    if result.returncode != expected:
        raise AssertionError(f"exit={result.returncode}, expected={expected}\nstdout={result.stdout}\nstderr={result.stderr}")
    return result


class ReviewKitTest(unittest.TestCase):
    def test_tax_net_and_gross(self):
        result = json.loads(run_cli("calc-tax", "--net", "100000", "--rate", "6%").stdout)
        self.assertEqual(result, {**result, "net": "100000.00", "tax": "6000.00", "gross": "106000.00"})
        result = json.loads(run_cli("calc-tax", "--gross", "113", "--rate", "0.13").stdout)
        self.assertEqual((result["net"], result["tax"]), ("100.00", "13.00"))

    def test_sum_and_difference(self):
        result = json.loads(run_cli("sum", "--values", "113000", "10600", "--compare", "123000").stdout)
        self.assertEqual(result["total"], "123600.00")
        self.assertEqual(result["difference_total_minus_compare"], "600.00")

    def test_non_finite_rejected(self):
        result = run_cli("sum", "--values", "NaN", expected=2)
        self.assertIn("有限", result.stderr)

    def test_docx_inventory_and_readonly(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.docx"
            document = f'''<?xml version="1.0"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
            <w:p><w:ins><w:r><w:t>新增</w:t></w:r></w:ins><w:del><w:r><w:delText>删除</w:delText></w:r></w:del></w:p>
            <w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>隐藏</w:t></w:r></w:p>
            <w:p><w:txbxContent><w:p><w:r><w:t>文本框</w:t></w:r></w:p></w:txbxContent></w:p>
            <w:tbl><w:tr><w:tc><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>
            </w:body></w:document>'''
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", document)
                archive.writestr("word/header1.xml", '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                archive.writestr("word/comments.xml", '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1"/></w:comments>')
                archive.writestr("word/footnotes.xml", '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"/></w:footnotes>')
                archive.writestr("word/media/image1.png", b"fake")
                archive.writestr("word/embeddings/object.bin", b"fake")
            manifest = Path(directory) / "manifest.json"
            run_cli("fingerprint", path, "-o", manifest)
            data = json.loads(manifest.read_text())
            inventory = data["files"][0]["docx_inventory"]
            self.assertEqual(inventory["counts"]["insertions"], 1)
            self.assertEqual(inventory["counts"]["deletions"], 1)
            self.assertEqual(inventory["counts"]["comments"], 1)
            self.assertEqual(inventory["counts"]["textboxes"], 1)
            self.assertEqual(inventory["counts"]["footnotes"], 1)
            self.assertEqual(inventory["counts"]["hidden_text"], 1)
            self.assertTrue(inventory["has_main_document"])
            self.assertGreaterEqual(inventory["counts"]["nested_tables"], 1)
            self.assertGreaterEqual(len(inventory["warnings"]), 5)
            run_cli("readonly-check", manifest, path)
            with zipfile.ZipFile(path, "a") as archive:
                archive.writestr("word/added.xml", "<x/>")
            failed = run_cli("readonly-check", manifest, path, expected=1)
            self.assertTrue(json.loads(failed.stdout)["changed"])

    def test_numbers_high_recall(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "numbers.txt"
            path.write_text("P0001 100 200\nP0002 违约金为百分之五\nP0003 大写壹万元整\nP0004 付款期限由双方协商\n双方应诚信履行本合同\n")
            result = json.loads(run_cli("numbers", path).stdout)
            self.assertEqual(result["count"], 4)
            self.assertEqual(result["records"][0]["locator"], "P0001")
            self.assertNotIn("诚信", "".join(x["text"] for x in result["records"]))

    def test_quote_verification_against_frozen_source(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "frozen.txt"
            report = Path(directory) / "report.md"
            source.write_text("第三条 金额为100元。\n")
            report.write_text(
                "### R-001\n- **原文或缺口**：\n  > 金额为100元。\n- **依据**：合同原文\n\n"
                "### R-002\n- **原文或缺口**：经核查全文缺失通知条款。\n"
            )
            passed = run_cli("verify-quotes", report, source)
            data = json.loads(passed.stdout)
            self.assertTrue(data["passed"])
            self.assertEqual(data["quote_count"], 1)
            self.assertEqual(data["missing_clause_claims"], 1)
            report.write_text(report.read_text().replace("100元", "1000元"))
            failed = run_cli("verify-quotes", report, source, expected=1)
            self.assertIn("逐字引文", failed.stdout)

    def test_output_cannot_overwrite_input(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "hetong.txt"
            path.write_text("P0001 金额100元\n")
            before = path.read_bytes()
            result = run_cli("numbers", path, "-o", path, expected=2)
            self.assertIn("不得覆盖", result.stderr)
            self.assertEqual(path.read_bytes(), before)
            manifest = Path(directory) / "baseline.json"
            run_cli("fingerprint", path, "-o", manifest)
            result = run_cli("readonly-check", manifest, "-o", path, expected=2)
            self.assertIn("不得覆盖", result.stderr)
            self.assertEqual(path.read_bytes(), before)

    def test_tax_ledger_totals_and_report_gate(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            ledger = directory / "ledger.json"
            checked = directory / "tax-check.json"
            ledger.write_text(json.dumps({
                "items": [
                    {"id": "设备", "net": "100000", "rate": "13%", "gross": "113000"},
                    {"id": "安装", "net": "10000", "rate": "6%", "tax": "600", "gross": "10600"},
                ],
                "stated_total": "123000",
            }, ensure_ascii=False))
            run_cli("tax-check", ledger, "-o", checked)
            result = json.loads(checked.read_text())
            self.assertEqual(result["totals"], {"net": "110000.00", "tax": "13600.00", "gross": "123600.00"})
            self.assertEqual(result["difference_gross_minus_stated"], "600.00")
            self.assertTrue(result["items_match_stated"])
            ledger_data = json.loads(ledger.read_text())
            ledger_data["items"][1]["tax"] = "601"
            ledger.write_text(json.dumps(ledger_data, ensure_ascii=False))
            discrepancy = json.loads(run_cli("tax-check", ledger).stdout)
            self.assertTrue(discrepancy["passed"])
            self.assertFalse(discrepancy["items_match_stated"])
            self.assertEqual(len(discrepancy["discrepancies"]), 1)
            ledger.write_text(json.dumps({**ledger_data, "items": ledger_data["items"][:1] + [{**ledger_data["items"][1], "tax": "600"}]}, ensure_ascii=False))
            report = directory / "report.md"
            text = (ROOT / "assets" / "review-report-template.md").read_text()
            text = text.replace(
                "- **风险等级 / 处置**：[S1阻断/S2重大/S3一般/S4建议] / [签前必改等]",
                "- **风险等级 / 处置**：S2重大 / 签前必改",
            )
            text = text.replace("| S2 重大 | [数量] | [ID 列表或“—”] |", "| S2 重大 | 1 | R-001 |")
            text = text.replace("## 二、审核边界与材料清单", "价税合计：不含税 110,000.00 元＋税额 13,600.00 元＝含税 123,600.00 元。\n\n## 二、审核边界与材料清单")
            report.write_text(text)
            run_cli("check-report", report, "--tax-check", checked)
            report.write_text(text.replace("税额 13,600.00 元", "税额待核"))
            failed = run_cli("check-report", report, "--tax-check", checked, expected=1)
            self.assertIn("合计税额", failed.stdout)

    def test_summary_count_mismatch_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.md"
            text = (ROOT / "assets" / "review-report-template.md").read_text()
            text = text.replace(
                "- **风险等级 / 处置**：[S1阻断/S2重大/S3一般/S4建议] / [签前必改等]",
                "- **风险等级 / 处置**：S2重大 / 签前必改",
            )
            text = text.replace("| S2 重大 | [数量] | [ID 列表或“—”] |", "| S2 重大 | 0 | R-001 |")
            path.write_text(text)
            failed = run_cli("check-report", path, expected=1)
            result = json.loads(failed.stdout)
            self.assertEqual(result["level_counts"]["S2"], 1)
            self.assertTrue(any("统计不一致" in x for x in result["errors"]))

    def test_clause_scheme_table_required_for_final(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "clauses.md"
            path.write_text(
                "### C-001｜对应 R-001\n"
                "- **保护目标与适用立场**：保护乙方\n"
                "- **建议文本**：完整条款正文。\n"
                "- **待确认变量**：无\n"
                "- **联动修改**：无\n"
                "- **法律核验**：纯商业建议，无待核法律命题\n"
            )
            failed = run_cli("check-clauses", path, "--final", expected=1)
            self.assertIn("确认方案一致性", failed.stdout)
            path.write_text(path.read_text() + "\n## 四、确认方案一致性表\n\n| C-ID | 确认方案原句 | 正文逐字对应 | 状态 | 语义差异与处置 |\n|---|---|---|---|---|\n| C-001 | 全部 | 全文 | 照录 | 无 |\n")
            failed = run_cli("check-clauses", path, "--final", expected=1)
            self.assertIn("逐句反查", failed.stdout)
            path.write_text(path.read_text() + "\n### 逐句反查记录\n\n| C-ID／款项 | 草案语义单元 | 对应确认原句 | 新增/收窄/扩大检查 | 处置 |\n|---|---|---|---|---|\n| C-001 | 全部 | 全部 | 无 | 保留 |\n")
            run_cli("check-clauses", path, "--final")

    def test_clause_anchor_gate_checks_suggested_text(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            path = directory / "clauses.md"
            anchors = directory / "anchors.json"
            path.write_text(
                "### C-001｜对应 R-001\n"
                "- **保护目标与适用立场**：保护乙方\n"
                "- **建议文本**：乙方提交验收申请后 5 个工作日内完成验收。\n"
                "- **待确认变量**：无\n"
                "- **联动修改**：无\n"
                "- **法律核验**：纯商业建议，无待核法律命题\n"
            )
            anchors.write_text(json.dumps({"items": [{"c_id": "C-001", "required": ["完整提交", "5 个工作日"], "forbidden": ["附带损失"]}]}, ensure_ascii=False))
            failed = run_cli("check-clauses", path, "--anchors", anchors, expected=1)
            self.assertIn("缺少确认锚点：完整提交", failed.stdout)
            path.write_text(path.read_text().replace("乙方提交", "乙方完整提交"))
            run_cli("check-clauses", path, "--anchors", anchors)
            path.write_text(path.read_text().replace("。", "，并排除附带损失。"))
            failed = run_cli("check-clauses", path, "--anchors", anchors, expected=1)
            self.assertIn("禁止锚点：附带损失", failed.stdout)

    def test_deliverable_templates_pass_structure_gate(self):
        run_cli("check-report", ROOT / "assets" / "review-report-template.md")
        run_cli("check-clauses", ROOT / "assets" / "clause-proposals-template.md")


if __name__ == "__main__":
    unittest.main()
