#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""office_kit.py / office_pptx.py 的自动验收测试。

全部材料由本测试现场生成，仅为虚构文本，不读取任何真实合同。测试覆盖：
环境自检、结构检查、无头批量渲染与缓存、CJK 文本层、PPTX 计划预演/提交/
独立验证、跨 run 保护、几何边界、宏/ZIP 攻击拒绝和原件只读。
"""

from __future__ import annotations

import hashlib
import importlib
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
KIT = SCRIPT_DIR / "office_kit.py"


def ensure_runtime() -> None:
    needed = ("docx", "pptx", "openpyxl", "PIL", "pypdfium2")
    if all(importlib.util.find_spec(name) is not None for name in needed):
        return
    sys.path.insert(0, str(SCRIPT_DIR))
    import office_kit

    python = office_kit._runtime_python()
    if python is None:
        raise SystemExit(f"Office 运行环境缺失，先运行：{KIT} setup")
    environment = os.environ.copy()
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    completed = subprocess.run(
        [str(python), str(Path(__file__).resolve()), *sys.argv[1:]],
        env=environment,
        check=False,
    )
    raise SystemExit(completed.returncode)


sys.dont_write_bytecode = True
ensure_runtime()


def load(name: str):
    return importlib.import_module(name)


def file_sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_cli(*arguments: str | Path, expect: int = 0) -> dict:
    environment = os.environ.copy()
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    completed = subprocess.run(
        [sys.executable, str(KIT), *[str(item) for item in arguments]],
        capture_output=True, text=True, encoding="utf-8", check=False, env=environment,
    )
    if completed.returncode != expect:
        raise AssertionError(
            f"CLI exit={completed.returncode}, expected={expect}\n"
            f"stdout:\n{completed.stdout[-4000:]}\nstderr:\n{completed.stderr[-4000:]}"
        )
    try:
        return json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise AssertionError(f"非 JSON 输出：{completed.stdout[-2000:]}") from exc


def build_docx(path: Path) -> None:
    Document = load("docx").Document
    document = Document()
    document.add_heading("虚构设备采购合同", level=1)
    document.add_paragraph("合同编号：DEMO-OFFICE-001")
    document.add_paragraph("含税总价 123,600 元，其中税额 13,600 元。")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "项目"
    table.cell(0, 1).text = "标准"
    table.cell(1, 0).text = "验收期限"
    table.cell(1, 1).text = "5 个工作日"
    document.save(path)


def build_pptx(path: Path, image_path: Path) -> None:
    Presentation = load("pptx").Presentation
    Inches = load("pptx.util").Inches
    Pt = load("pptx.util").Pt
    RGBColor = load("pptx.dml.color").RGBColor
    Image = load("PIL.Image")
    Image.new("RGB", (32, 32), (210, 220, 230)).save(image_path)

    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    title = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), Inches(0.8))
    run = title.text_frame.paragraphs[0].add_run()
    run.text = "虚构采购合同审阅汇报"
    run.font.name = "PingFang SC"
    run.font.size = Pt(32)
    run.font.bold = True

    body = slide.shapes.add_textbox(Inches(0.9), Inches(2.0), Inches(11.2), Inches(2.0))
    frame = body.text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    for text, bold in (("付款比例为", False), ("30%", True), ("，验收期限为5个工作日。", False)):
        run = paragraph.add_run()
        run.text = text
        run.font.name = "PingFang SC"
        run.font.size = Pt(22)
        run.font.bold = bold
    second = frame.add_paragraph()
    run = second.add_run()
    run.text = "重点：价款、验收与责任上限需要联动核对。"
    run.font.name = "PingFang SC"
    run.font.size = Pt(18)
    slide.notes_slide.notes_text_frame.text = "虚构备注：初始版本。"

    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    title = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), Inches(0.8))
    run = title.text_frame.paragraphs[0].add_run()
    run.text = "关键金额与验收节点"
    run.font.name = "PingFang SC"
    run.font.size = Pt(28)
    run.font.bold = True
    table_shape = slide.shapes.add_table(2, 2, Inches(0.9), Inches(1.8), Inches(7.0), Inches(1.6))
    table = table_shape.table
    for row, values in enumerate((("项目", "金额/期限"), ("含税总价", "123,600 元"))):
        for column, value in enumerate(values):
            cell = table.cell(row, column)
            cell.text = value
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "PingFang SC"
                    run.font.size = Pt(16)
    box = slide.shapes.add_textbox(Inches(8.5), Inches(1.8), Inches(3.8), Inches(1.4))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "责任上限需明确基数与例外"
    run.font.name = "PingFang SC"
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = RGBColor(31, 55, 92)
    slide.shapes.add_picture(str(image_path), Inches(0.9), Inches(4.2), Inches(0.8), Inches(0.8))
    presentation.save(path)


def build_xlsx(path: Path) -> None:
    Workbook = load("openpyxl").Workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "付款计划"
    sheet.append(["期次", "比例", "金额"])
    sheet.append(["首付款", 0.3, "=123600*B2"])
    sheet.append(["验收款", 0.7, "=123600*B3"])
    workbook.save(path)


class OfficeFixture(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temporary = tempfile.TemporaryDirectory(prefix="contact-office-test-")
        cls.root = Path(cls.temporary.name)
        cls.docx = cls.root / "虚构采购合同.docx"
        cls.pptx = cls.root / "虚构审阅汇报.pptx"
        cls.xlsx = cls.root / "虚构付款计划.xlsx"
        cls.icon = cls.root / "icon.png"
        build_docx(cls.docx)
        build_pptx(cls.pptx, cls.icon)
        build_xlsx(cls.xlsx)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temporary.cleanup()


class RuntimeTests(OfficeFixture):
    def test_doctor_quick(self) -> None:
        data = run_cli("doctor")
        self.assertTrue(data["ok"])
        self.assertTrue(data["packages"]["python-pptx"])
        self.assertTrue(data["libreoffice"]["version"].startswith("LibreOffice"))

    def test_doctor_full_cjk_render(self) -> None:
        data = run_cli("doctor", "--full", "--timeout", "240")
        self.assertTrue(data["ok"])
        self.assertTrue(data["full_render"]["cjk_pdf_text"])
        self.assertEqual(data["full_render"]["libreoffice_invocations"], 1)


class InspectTests(OfficeFixture):
    def test_docx_pptx_xlsx_inspection(self) -> None:
        data = run_cli("inspect", self.docx, self.pptx, self.xlsx)
        self.assertTrue(all(item["ok"] for item in data["results"]))
        docx, pptx, xlsx = data["results"]
        self.assertEqual(docx["format"], "docx")
        self.assertGreaterEqual(docx["paragraphs"], 3)
        self.assertEqual(docx["tables"], 1)
        self.assertEqual(pptx["slide_count"], 2)
        self.assertEqual(pptx["slides"][0]["notes"], "虚构备注：初始版本。")
        self.assertTrue(any(shape.get("table") for shape in pptx["slides"][1]["shapes"]))
        self.assertEqual(xlsx["formula_count"], 2)
        self.assertEqual(xlsx["sheets"][0]["name"], "付款计划")


class RenderTests(OfficeFixture):
    def test_batch_render_cache_and_cjk(self) -> None:
        output = self.root / "render"
        before = {path: file_sha(path) for path in (self.docx, self.pptx)}
        first = run_cli("render", self.docx, self.pptx, "--out", output, "--dpi", "100")
        self.assertEqual(first["libreoffice_invocations"], 1)
        self.assertFalse(any(entry["cached"] for entry in first["entries"]))
        self.assertEqual(len(first["entries"]), 2)
        all_pages = 0
        for entry in first["entries"]:
            self.assertTrue(Path(entry["pdf"]).is_file())
            self.assertTrue(entry["images"])
            self.assertTrue(entry["contact_sheets"])
            all_pages += len(entry["images"])
            PdfReader = load("pypdf").PdfReader
            text = "\n".join(page.extract_text() or "" for page in PdfReader(entry["pdf"]).pages)
            self.assertGreaterEqual(len({character for character in text if "\u4e00" <= character <= "\u9fff"}), 6)
        self.assertGreaterEqual(all_pages, 3)
        self.assertEqual(before, {path: file_sha(path) for path in (self.docx, self.pptx)})
        second = run_cli("render", self.docx, self.pptx, "--out", output, "--dpi", "100")
        self.assertEqual(second["libreoffice_invocations"], 0)
        self.assertTrue(all(entry["cached"] for entry in second["entries"]))
        tampered = Path(second["entries"][0]["images"][0])
        tampered.write_bytes(b"tampered-cache")
        third = run_cli("render", self.docx, self.pptx, "--out", output, "--dpi", "100")
        self.assertEqual(third["libreoffice_invocations"], 1)
        self.assertFalse(all(entry["cached"] for entry in third["entries"]))

    def test_page_range(self) -> None:
        output = self.root / "render-page-range"
        data = run_cli("render", self.pptx, "--out", output, "--pages", "2", "--dpi", "60")
        self.assertEqual(len(data["entries"][0]["images"]), 1)
        self.assertTrue(data["entries"][0]["images"][0].endswith("page-002.png"))

    def test_macro_and_zip_traversal_rejected(self) -> None:
        macro = self.root / "macro.pptx"
        with zipfile.ZipFile(self.pptx) as source, zipfile.ZipFile(macro, "w") as target:
            for info in source.infolist():
                target.writestr(info, source.read(info.filename))
            target.writestr("vbaProject.bin", b"not-a-real-macro")
        data = run_cli("inspect", macro, expect=1)
        self.assertIn("VBA", data["results"][0]["error"])

        traversal = self.root / "traversal.pptx"
        with zipfile.ZipFile(traversal, "w") as archive:
            archive.writestr("../escape.txt", "bad")
        data = run_cli("inspect", traversal, expect=1)
        self.assertIn("路径不安全", data["results"][0]["error"])


class PptxEditTests(OfficeFixture):
    @classmethod
    def setUpClass(cls) -> None:
        super().setUpClass()
        cls.inspect_data = run_cli("inspect", cls.pptx)["results"][0]
        cls.inspect_source_hash = cls.inspect_data["source_sha256"]
        cls.title_id = next(shape["shape_id"] for shape in cls.inspect_data["slides"][0]["shapes"] if shape.get("text") == "虚构采购合同审阅汇报")
        cls.body_id = next(shape["shape_id"] for shape in cls.inspect_data["slides"][0]["shapes"] if "付款比例" in (shape.get("text") or ""))
        cls.table_id = next(shape["shape_id"] for shape in cls.inspect_data["slides"][1]["shapes"] if shape.get("table"))
        cls.box_id = next(shape["shape_id"] for shape in cls.inspect_data["slides"][1]["shapes"] if "责任上限" in (shape.get("text") or ""))

    def make_plan(self, name: str, operations: list[dict], output_name: str | None = None) -> tuple[Path, Path]:
        output = self.root / "edits" / (output_name or f"{name}.pptx")
        output.parent.mkdir(parents=True, exist_ok=True)
        plan = {
            "schema": "contact-pptx-plan/1",
            "source": str(self.pptx),
            "source_sha256": self.inspect_source_hash,
            "output": str(output),
            "operations": [{"id": f"O-{index:03d}", **operation} for index, operation in enumerate(operations, 1)],
        }
        path = self.root / "edits" / f"{name}-plan.json"
        path.write_text(json.dumps(plan, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return path, output

    def standard_operations(self) -> list[dict]:
        return [
            {"op": "replace_text", "slide": 1, "shape_id": self.body_id, "paragraph": 1, "old": "30%", "new": "40%"},
            {"op": "set_font", "slide": 1, "shape_id": self.title_id, "paragraph": 1, "size_pt": 34, "bold": True},
            {"op": "set_table_cell", "slide": 2, "shape_id": self.table_id, "row": 2, "column": 2, "text": "123,600 元（含税）"},
            {"op": "set_geometry", "slide": 2, "shape_id": self.box_id, "left_in": 8.3, "width_in": 4.0},
            {"op": "set_notes", "slide": 1, "text": "虚构备注：已复核金额与验收节点。\n仅作测试。"},
        ]

    def test_dry_run_commit_verify_and_render(self) -> None:
        plan, output = self.make_plan("standard", self.standard_operations())
        source_hash = file_sha(self.pptx)
        dry = run_cli("pptx", "apply", plan, "--dry-run")
        self.assertFalse(dry["commit"])
        self.assertFalse(output.exists())
        self.assertEqual(source_hash, file_sha(self.pptx))

        result_path = self.root / "edits" / "standard-result.json"
        render_dir = self.root / "edits" / "standard-render"
        committed = run_cli(
            "pptx", "apply", plan, "--commit", "--result", result_path, "--render-dir", render_dir
        )
        self.assertTrue(committed["commit"])
        self.assertTrue(output.is_file())
        self.assertTrue(committed["media_preserved"])
        self.assertTrue(committed["render"]["ok"])
        self.assertEqual(source_hash, file_sha(self.pptx))
        verified = run_cli("pptx", "verify", plan, output, "--result", result_path)
        self.assertTrue(verified["ok"])
        self.assertEqual(verified["slide_count"], 2)
        self.assertEqual(file_sha(output), committed["output_sha256"])

        after = run_cli("inspect", output)["results"][0]
        text = "\n".join((shape.get("text") or "") for slide in after["slides"] for shape in slide["shapes"])
        for slide in after["slides"]:
            for shape in slide["shapes"]:
                for row in shape.get("table", {}).get("cells", []):
                    for cell in row:
                        text += "\n" + "\n".join(paragraph["text"] for paragraph in cell)
        self.assertIn("付款比例为40%", text)
        self.assertIn("123,600 元（含税）", text)
        self.assertEqual(after["slides"][0]["notes"], "虚构备注：已复核金额与验收节点。\n仅作测试。")
        box = next(shape for slide in after["slides"] for shape in slide["shapes"] if shape["shape_id"] == self.box_id and slide["slide"] == 2)
        self.assertAlmostEqual(box["width_in"], 4.0, places=2)

    def test_cross_run_requires_explicit_permission(self) -> None:
        operation = {
            "op": "replace_text", "slide": 1, "shape_id": self.body_id, "paragraph": 1,
            "old": "付款比例为30%", "new": "付款比例调整为40%",
        }
        plan, output = self.make_plan("cross-refused", [operation])
        failed = run_cli("pptx", "apply", plan, "--dry-run", expect=2)
        self.assertIn("allow_cross_run", failed["error"])
        operation["allow_cross_run"] = True
        plan, output = self.make_plan("cross-allowed", [operation])
        data = run_cli("pptx", "apply", plan, "--commit")
        self.assertTrue(data["warnings"])
        after = run_cli("inspect", output)["results"][0]
        body = next(shape for shape in after["slides"][0]["shapes"] if shape["shape_id"] == self.body_id)
        self.assertIn("付款比例调整为40%", body["text"])

    def test_paragraph_style_loss_requires_permission(self) -> None:
        operation = {
            "op": "set_paragraph_text", "slide": 1, "shape_id": self.body_id,
            "paragraph": 1, "text": "付款比例调整为40%，验收期限为5个工作日。",
        }
        plan, _ = self.make_plan("style-refused", [operation])
        failed = run_cli("pptx", "apply", plan, "--dry-run", expect=2)
        self.assertIn("allow_style_loss", failed["error"])
        operation["allow_style_loss"] = True
        plan, output = self.make_plan("style-allowed", [operation])
        data = run_cli("pptx", "apply", plan, "--commit")
        self.assertTrue(any(item.get("style_loss") for item in data["operations"]))
        after = run_cli("inspect", output)["results"][0]
        body = next(shape for shape in after["slides"][0]["shapes"] if shape["shape_id"] == self.body_id)
        self.assertIn("付款比例调整为40%", body["text"])

    def test_geometry_out_of_bounds_rejected(self) -> None:
        operation = {
            "op": "set_geometry", "slide": 2, "shape_id": self.box_id,
            "left_in": 12.0, "width_in": 3.0,
        }
        plan, _ = self.make_plan("bounds", [operation])
        failed = run_cli("pptx", "apply", plan, "--dry-run", expect=2)
        self.assertIn("超出幻灯片", failed["error"])

    def test_sha_mismatch_and_in_place_rejected(self) -> None:
        plan, _ = self.make_plan("sha", self.standard_operations())
        data = json.loads(plan.read_text(encoding="utf-8"))
        data["source_sha256"] = "0" * 64
        bad = self.root / "edits" / "sha-bad.json"
        bad.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        failed = run_cli("pptx", "apply", bad, "--dry-run", expect=2)
        self.assertIn("哈希", failed["error"])

        plan, _ = self.make_plan("in-place", self.standard_operations(), output_name="../虚构审阅汇报.pptx")
        # 实际解析后仍等于源文件时也必须拒绝
        data = json.loads(plan.read_text(encoding="utf-8"))
        data["output"] = str(self.pptx)
        bad = self.root / "edits" / "in-place-bad.json"
        bad.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        failed = run_cli("pptx", "apply", bad, "--dry-run", expect=2)
        self.assertIn("不得等于源文件", failed["error"])

    def test_existing_output_requires_overwrite(self) -> None:
        plan, output = self.make_plan("existing", self.standard_operations())
        output.write_bytes(b"existing")
        failed = run_cli("pptx", "apply", plan, "--commit", expect=2)
        self.assertIn("输出已存在", failed["error"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
