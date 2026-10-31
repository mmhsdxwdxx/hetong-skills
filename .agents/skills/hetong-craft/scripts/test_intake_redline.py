#!/usr/bin/env python3
"""Meaningful write-gate checks on synthetic DOCX documents only."""
from pathlib import Path
import json
import subprocess
import sys
import unittest

import hetong_intake as intake
import test_hetong_redline as fixtures


class IntakeRedlineGateTests(fixtures.Base):
    def source_and_plan(self, comment=False):
        source = self.write("synthetic.docx", fixtures.p(fixtures.r("名称为旧名称。")))
        edit = {"id": "C1" if comment else "R1", "para": "RL0001", "old": "旧名称"}
        edit.update({"comment": "仅用于虚构测试的批注。"} if comment else {"new": "新名称", "reason": "虚构名称规范化。"})
        code, plan, path = self.plan(source, [edit])
        self.assertEqual(code, 0, plan)
        return source, path

    def receipt(self, source, action, known_extra=None):
        path = Path(self.path("intake.json"))
        known = {"perspective": "buyer", "stage": "unsigned", "action": action, "categories": []}
        known.update(known_extra or {})
        intake.init_state(path, [source], known, "虚构测试：用户明确指定此处理方式，本次不用预设条款。")
        return path

    def apply_with_receipt(self, source, plan, receipt):
        return self.json_cli("apply", source, plan, "-o", self.path("result.docx"), "--intake-state", str(receipt))

    def test_review_only_does_not_write_even_comments(self):
        source, plan = self.source_and_plan(comment=True)
        before = Path(source).read_bytes()
        code, result = self.apply_with_receipt(source, plan, self.receipt(source, "review"))
        self.assertNotEqual(code, 0, result)
        self.assertFalse(Path(self.path("result.docx")).exists())
        self.assertEqual(Path(source).read_bytes(), before)

    def test_comment_permission_does_not_allow_replacement(self):
        source, plan = self.source_and_plan()
        code, result = self.apply_with_receipt(source, plan, self.receipt(source, "comment"))
        self.assertNotEqual(code, 0, result)
        self.assertFalse(Path(self.path("result.docx")).exists())

    def test_comment_permission_allows_actual_comment(self):
        source, plan = self.source_and_plan(comment=True)
        code, result = self.apply_with_receipt(source, plan, self.receipt(source, "comment"))
        self.assertEqual(code, 0, result)
        self.assertTrue(result["intake_receipt"])
        self.assertEqual(result["author"], "SR")
        self.assertEqual(self.local_count(self.read_docx_part(self.path("result.docx")), "ins"), 0)

    def test_different_source_path_cannot_borrow_receipt(self):
        source, plan = self.source_and_plan()
        other = Path(self.path("other.docx"))
        other.write_bytes(Path(source).read_bytes())
        receipt = self.receipt(other, "revise", {"delivery": "tracked"})
        code, result = self.apply_with_receipt(source, plan, receipt)
        self.assertNotEqual(code, 0, result)
        self.assertFalse(Path(self.path("result.docx")).exists())

    def test_revision_permission_passes_real_engine_verification(self):
        source, plan = self.source_and_plan()
        receipt = self.receipt(source, "revise", {"delivery": "tracked"})
        process = subprocess.run([sys.executable, "-I", str(Path(__file__).with_name("hetong_redline.py")), "apply", source, plan,
                                  "-o", self.path("result.docx"), "--intake-state", str(receipt)], capture_output=True, text=True)
        self.assertEqual(process.returncode, 0, process.stdout + process.stderr)
        result = json.loads(process.stdout)
        self.assertTrue(result["verification"]["ok"])
        self.assertEqual(self.local_count(self.read_docx_part(self.path("result.docx")), "ins"), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
