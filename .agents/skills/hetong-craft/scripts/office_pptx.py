#!/usr/bin/env python3
"""PPTX 计划式精修：先预演、后原子另存、再结构验证。

只修改计划明确指向的文本、表格单元格、字体、形状几何和备注；不复制幻灯片、
不处理 SmartArt/图表内部对象，也不自动接受或删除其他内容。运行本模块的 Python
必须来自 Office 隔离环境，入口由 office_kit.py 负责切换。
"""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from office_kit import OfficeKitError, _validate_input, sha256_file

PLAN_SCHEMA = "contact-pptx-plan/1"
RESULT_SCHEMA = "contact-pptx-result/1"
TEXT_OPS = {"replace_text", "set_paragraph_text", "set_table_cell", "set_font"}
SUPPORTED_OPS = TEXT_OPS | {"set_geometry", "set_notes"}


def _load(name: str) -> Any:
    import importlib

    try:
        return importlib.import_module(name)
    except ImportError as exc:
        raise OfficeKitError(f"PPTX 运行环境缺少模块 {name}") from exc


def _pptx() -> Any:
    return _load("pptx")


def _rgb_color(value: str) -> Any:
    if not isinstance(value, str) or not value.startswith("#") or len(value) != 7:
        raise OfficeKitError(f"颜色必须为 #RRGGBB：{value!r}")
    try:
        int(value[1:], 16)
    except ValueError as exc:
        raise OfficeKitError(f"颜色无效：{value!r}") from exc
    return _load("pptx.dml.color").RGBColor.from_string(value[1:].upper())


def _font_snapshot(run: Any) -> dict[str, Any]:
    font = run.font
    try:
        color = str(font.color.rgb) if font.color.type is not None and font.color.rgb is not None else None
    except (AttributeError, ValueError, TypeError):
        color = None
    try:
        size = round(float(font.size.pt), 3) if font.size is not None else None
    except (AttributeError, ValueError, TypeError):
        size = None
    rpr = run._r.rPr
    return {
        "name": font.name,
        "size_pt": size,
        "bold": font.bold,
        "italic": font.italic,
        "underline": font.underline,
        "color": color,
        "rpr_xml": str(rpr.xml) if rpr is not None else None,
    }


def _paragraph_snapshot(paragraph: Any) -> dict[str, Any]:
    return {
        "text": paragraph.text,
        "level": paragraph.level,
        "alignment": str(paragraph.alignment) if paragraph.alignment is not None else None,
        "runs": [_font_snapshot(run) for run in paragraph.runs],
    }


def _text_frame_snapshot(text_frame: Any) -> list[dict[str, Any]]:
    return [_paragraph_snapshot(paragraph) for paragraph in text_frame.paragraphs]


def _table_snapshot(table: Any) -> list[list[list[dict[str, Any]]]]:
    return [
        [_text_frame_snapshot(cell.text_frame) for cell in row.cells]
        for row in table.rows
    ]


def _shape_type_name(shape: Any) -> str:
    shape_type = getattr(shape, "shape_type", None)
    return getattr(shape_type, "name", str(shape_type))


def _iter_shapes(shapes: Any, path: str = "root"):
    for shape in shapes:
        yield shape, path
        if _shape_type_name(shape) == "GROUP":
            yield from _iter_shapes(shape.shapes, f"{path}/group:{shape.shape_id}")


def _find_shape(slide: Any, shape_id: int) -> tuple[Any, str]:
    for shape, path in _iter_shapes(slide.shapes):
        if int(shape.shape_id) == shape_id:
            return shape, path
    raise OfficeKitError(f"第 {slide.slide_id} 页找不到 shape_id={shape_id}")


def _shape_snapshot(shape: Any, path: str) -> dict[str, Any]:
    result: dict[str, Any] = {
        "shape_id": int(shape.shape_id),
        "name": shape.name,
        "type": _shape_type_name(shape),
        "path": path,
        "left": int(shape.left) if shape.left is not None else None,
        "top": int(shape.top) if shape.top is not None else None,
        "width": int(shape.width) if shape.width is not None else None,
        "height": int(shape.height) if shape.height is not None else None,
        "rotation": round(float(shape.rotation), 3) if getattr(shape, "rotation", None) is not None else None,
        "text": None,
        "paragraphs": None,
        "table": None,
    }
    if getattr(shape, "has_text_frame", False):
        result["text"] = shape.text_frame.text
        result["paragraphs"] = _text_frame_snapshot(shape.text_frame)
    if getattr(shape, "has_table", False):
        result["table"] = _table_snapshot(shape.table)
    return result


def _slide_shapes_snapshot(slide: Any) -> dict[str, dict[str, Any]]:
    return {
        f"{slide.slide_id}:{shape.shape_id}": _shape_snapshot(shape, path)
        for shape, path in _iter_shapes(slide.shapes)
    }


def _notes_snapshot(slide: Any, create: bool = False) -> list[dict[str, Any]] | None:
    has_notes = bool(getattr(slide, "has_notes_slide", False))
    if not has_notes and not create:
        return None
    text_frame = slide.notes_slide.notes_text_frame
    return _text_frame_snapshot(text_frame)


def _notes_text(slide: Any) -> str | None:
    has_notes = bool(getattr(slide, "has_notes_slide", False))
    if not has_notes:
        return None
    return slide.notes_slide.notes_text_frame.text


def inspect_pptx(path: Path, max_items: int = 300) -> dict[str, Any]:
    _validate_input(path)
    Presentation = _pptx().Presentation
    presentation = Presentation(str(path))
    slides = []
    truncated = False
    for slide_number, slide in enumerate(presentation.slides, 1):
        shapes = []
        for shape, shape_path in _iter_shapes(slide.shapes):
            item: dict[str, Any] = {
                "shape_id": int(shape.shape_id),
                "name": shape.name,
                "type": _shape_type_name(shape),
                "path": shape_path,
                "left_in": round(shape.left.inches, 3) if shape.left is not None else None,
                "top_in": round(shape.top.inches, 3) if shape.top is not None else None,
                "width_in": round(shape.width.inches, 3) if shape.width is not None else None,
                "height_in": round(shape.height.inches, 3) if shape.height is not None else None,
                "rotation": round(float(shape.rotation), 3) if getattr(shape, "rotation", None) is not None else None,
                "is_placeholder": bool(getattr(shape, "is_placeholder", False)),
            }
            if getattr(shape, "has_text_frame", False):
                item["text"] = shape.text_frame.text
                item["paragraphs"] = _text_frame_snapshot(shape.text_frame)
            if getattr(shape, "has_table", False):
                item["table"] = {
                    "rows": len(shape.table.rows),
                    "columns": len(shape.table.columns),
                    "cells": _table_snapshot(shape.table),
                }
            if getattr(shape, "has_chart", False):
                item["chart"] = {
                    "type": str(shape.chart.chart_type),
                    "has_title": bool(shape.chart.has_title),
                }
            if _shape_type_name(shape) == "PICTURE":
                try:
                    item["picture"] = {
                        "filename": shape.image.filename,
                        "extension": shape.image.ext,
                        "size": [shape.image.size.width, shape.image.size.height],
                    }
                except (AttributeError, ValueError):
                    item["picture"] = {"filename": None, "extension": None}
            shapes.append(item)
        notes = _notes_text(slide)
        hidden = slide._element.get("show") == "0"
        slides.append({
            "slide": slide_number,
            "slide_id": int(slide.slide_id),
            "hidden": hidden,
            "shapes": shapes[:max_items],
            "notes": notes,
        })
        if len(shapes) > max_items:
            truncated = True
    return {
        "format": "pptx",
        "ok": True,
        "source_sha256": sha256_file(path),
        "slide_width_in": round(presentation.slide_width.inches, 3),
        "slide_height_in": round(presentation.slide_height.inches, 3),
        "slide_count": len(presentation.slides),
        "slides": slides[:max_items],
        "truncated": truncated or len(presentation.slides) > max_items,
    }


def _resolve_plan_path(value: Any, base: Path, label: str) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise OfficeKitError(f"计划缺少 {label} 路径")
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = base / path
    return path.resolve()


def _load_plan(plan_path: Path) -> dict[str, Any]:
    plan_path = plan_path.expanduser().resolve()
    if not plan_path.is_file():
        raise OfficeKitError(f"计划不存在：{plan_path}")
    try:
        plan = json.loads(plan_path.read_text())
    except json.JSONDecodeError as exc:
        raise OfficeKitError(f"计划 JSON 无效：{exc}") from exc
    if not isinstance(plan, dict):
        raise OfficeKitError("计划根节点必须是对象")
    if plan.get("schema") not in {PLAN_SCHEMA, "contact-pptx-plan/1.0"}:
        raise OfficeKitError(f"计划 schema 必须为 {PLAN_SCHEMA}")
    base = plan_path.parent
    source = _resolve_plan_path(plan.get("source"), base, "source")
    output = _resolve_plan_path(plan.get("output"), base, "output")
    if source.suffix.lower() != ".pptx":
        raise OfficeKitError("PPTX 精修仅接受 .pptx 源文件")
    if output.suffix.lower() != ".pptx":
        raise OfficeKitError("输出必须为 .pptx")
    if not source.is_file():
        raise OfficeKitError(f"源文件不存在：{source}")
    if output == source:
        raise OfficeKitError("输出不得等于源文件；禁止原地修改")
    expected_sha = plan.get("source_sha256")
    if not isinstance(expected_sha, str) or len(expected_sha) != 64:
        raise OfficeKitError("计划必须包含 64 位 source_sha256")
    actual_sha = sha256_file(source)
    if actual_sha != expected_sha.lower():
        raise OfficeKitError(
            f"源文件哈希与计划不一致：计划 {expected_sha.lower()}，当前 {actual_sha}"
        )
    operations = plan.get("operations")
    if not isinstance(operations, list) or not operations:
        raise OfficeKitError("计划 operations 必须为非空数组")
    if len(operations) > 1000:
        raise OfficeKitError("单次计划最多 1000 个操作")
    seen: set[str] = set()
    for index, operation in enumerate(operations, 1):
        if not isinstance(operation, dict):
            raise OfficeKitError(f"operations[{index}] 必须是对象")
        op_id = operation.get("id")
        if not isinstance(op_id, str) or not op_id:
            raise OfficeKitError(f"operations[{index}] 缺少字符串 id")
        if op_id in seen:
            raise OfficeKitError(f"操作 id 重复：{op_id}")
        seen.add(op_id)
        if operation.get("op") not in SUPPORTED_OPS:
            raise OfficeKitError(f"{op_id}: 不支持的 op={operation.get('op')!r}")
    return {
        "path": plan_path,
        "data": plan,
        "source": source,
        "output": output,
        "source_sha256": actual_sha,
        "operations": operations,
    }


def _slide_at(presentation: Any, number: Any) -> Any:
    if not isinstance(number, int) or isinstance(number, bool):
        raise OfficeKitError("slide 必须为 1-based 整数")
    if number < 1 or number > len(presentation.slides):
        raise OfficeKitError(f"slide={number} 超出 1-{len(presentation.slides)}")
    return presentation.slides[number - 1]


def _shape_at(presentation: Any, operation: dict[str, Any]) -> tuple[Any, Any, str]:
    slide = _slide_at(presentation, operation.get("slide"))
    shape_id = operation.get("shape_id")
    if not isinstance(shape_id, int) or isinstance(shape_id, bool):
        raise OfficeKitError("shape_id 必须为整数")
    shape, path = _find_shape(slide, shape_id)
    return slide, shape, path


def _surface(operation: dict[str, Any], presentation: Any) -> tuple[Any, Any, str, str]:
    slide, shape, shape_path = _shape_at(presentation, operation)
    if "table" in operation or "row" in operation or "column" in operation or operation.get("op") == "set_table_cell":
        if not getattr(shape, "has_table", False):
            raise OfficeKitError(f"shape_id={shape.shape_id} 不是表格")
        row = operation.get("row")
        column = operation.get("column")
        if not isinstance(row, int) or not isinstance(column, int) or row < 1 or column < 1:
            raise OfficeKitError("表格 row/column 必须为 1-based 整数")
        if row > len(shape.table.rows) or column > len(shape.table.columns):
            raise OfficeKitError(
                f"单元格 {row},{column} 超出表格 {len(shape.table.rows)}x{len(shape.table.columns)}"
            )
        cell = shape.table.cell(row - 1, column - 1)
        key = f"{slide.slide_id}:{shape.shape_id}:cell:{row},{column}"
        return slide, cell.text_frame, shape_path, key
    if not getattr(shape, "has_text_frame", False):
        raise OfficeKitError(f"shape_id={shape.shape_id} 没有可编辑文本")
    key = f"{slide.slide_id}:{shape.shape_id}:text"
    return slide, shape.text_frame, shape_path, key


def _validate_paragraph_index(operation: dict[str, Any], text_frame: Any) -> int | None:
    paragraph = operation.get("paragraph")
    if paragraph is None:
        return None
    if not isinstance(paragraph, int) or isinstance(paragraph, bool) or paragraph < 1:
        raise OfficeKitError("paragraph 必须为 1-based 整数")
    if paragraph > len(text_frame.paragraphs):
        raise OfficeKitError(f"paragraph={paragraph} 超出 1-{len(text_frame.paragraphs)}")
    return paragraph - 1


def _replace_in_paragraph(
    paragraph: Any,
    old: str,
    new: str,
    occurrence: int,
    allow_cross_run: bool,
) -> dict[str, Any]:
    if not old:
        raise OfficeKitError("replace_text.old 不得为空")
    if "\n" in old or "\r" in old or "\n" in new or "\r" in new:
        raise OfficeKitError("replace_text 不支持跨行文本；请按段操作")
    runs = list(paragraph.runs)
    texts = [run.text for run in runs]
    concatenated = "".join(texts)
    starts = []
    cursor = 0
    while True:
        found = concatenated.find(old, cursor)
        if found < 0:
            break
        starts.append(found)
        cursor = found + max(1, len(old))
    if not starts:
        return {"matched": False, "count": 0}
    if occurrence < 1 or occurrence > len(starts):
        return {"matched": False, "count": len(starts), "requested": occurrence}
    start = starts[occurrence - 1]
    end = start + len(old)
    spans = []
    position = 0
    for index, text in enumerate(texts):
        spans.append((position, position + len(text), index))
        position += len(text)
    start_span = next((span for span in spans if span[0] <= start < span[1]), None)
    end_span = next((span for span in spans if span[0] < end <= span[1]), None)
    if start_span is None or end_span is None:
        return {"matched": False, "count": len(starts)}
    if start_span[2] == end_span[2]:
        run = runs[start_span[2]]
        offset = start - start_span[0]
        run.text = texts[start_span[2]][:offset] + new + texts[start_span[2]][offset + len(old):]
        return {"matched": True, "count": len(starts), "cross_run": False}
    if not allow_cross_run:
        return {
            "matched": False, "count": len(starts), "cross_run": True,
            "message": "目标跨多个 run；确认格式影响后在操作中设置 allow_cross_run=true",
        }
    first_index, last_index = start_span[2], end_span[2]
    first_offset = start - start_span[0]
    last_offset = end - end_span[0]
    runs[first_index].text = texts[first_index][:first_offset] + new
    for index in range(first_index + 1, last_index):
        runs[index].text = ""
    runs[last_index].text = texts[last_index][last_offset:]
    return {"matched": True, "count": len(starts), "cross_run": True, "runs_touched": last_index - first_index + 1}


def _replace_text(operation: dict[str, Any], presentation: Any) -> dict[str, Any]:
    old = operation.get("old")
    new = operation.get("new")
    if not isinstance(old, str) or not isinstance(new, str):
        raise OfficeKitError("replace_text 需要字符串 old/new")
    occurrence = operation.get("occurrence")
    strict_single = occurrence is None
    if occurrence is None:
        occurrence = 1
    if not isinstance(occurrence, int) or isinstance(occurrence, bool) or occurrence < 1:
        raise OfficeKitError("occurrence 必须为 1-based 整数")
    allow_cross = bool(operation.get("allow_cross_run", False))
    _, text_frame, _, key = _surface(operation, presentation)
    paragraph_index = _validate_paragraph_index(operation, text_frame)
    all_paragraphs = list(text_frame.paragraphs)
    paragraphs = [all_paragraphs[paragraph_index]] if paragraph_index is not None else all_paragraphs
    matches = []
    for item in paragraphs:
        raw = "".join(run.text for run in item.runs)
        count = raw.count(old)
        if count:
            matches.append((item, count))
    total = sum(count for _, count in matches)
    if total == 0:
        raise OfficeKitError(f"未找到 old={old!r}；检查 paragraph/单元格定位与 run 边界")
    if strict_single and total > 1:
        raise OfficeKitError(f"old 在目标中出现 {total} 次；必须用 paragraph 或 occurrence 精确指定")
    remaining = occurrence
    chosen = None
    for paragraph, count in matches:
        if remaining <= count:
            chosen = (paragraph, remaining)
            break
        remaining -= count
    if chosen is None:
        raise OfficeKitError(f"occurrence={occurrence} 超出匹配数 {total}")
    paragraph, local_occurrence = chosen
    if not isinstance(local_occurrence, int):
        raise OfficeKitError("内部匹配计数异常")
    paragraph_number = all_paragraphs.index(paragraph) + 1
    before = paragraph.text
    result = _replace_in_paragraph(paragraph, old, new, local_occurrence, allow_cross)
    if not result.get("matched"):
        if result.get("cross_run"):
            raise OfficeKitError(f"{key}: {result.get('message')}")
        raise OfficeKitError(f"{key}: 替换未执行（匹配数或 run 边界异常）")
    return {
        "key": key, "paragraph": paragraph_number,
        "before": before, "after": paragraph.text,
        "cross_run": result.get("cross_run", False),
        "runs_touched": result.get("runs_touched", 1),
    }


def _set_paragraph_object(
    paragraph: Any,
    text: str,
    allow_style_loss: bool,
) -> tuple[str, bool]:
    if "\n" in text or "\r" in text:
        raise OfficeKitError("单段文本不得包含换行；请分多个操作或使用备注")
    runs = list(paragraph.runs)
    xml = paragraph._p.xml
    special = "a:br" in xml or "a:fld" in xml or "a:hlinkClick" in xml
    styles = {str(run._r.rPr.xml) if run._r.rPr is not None else "" for run in runs}
    style_loss = len(styles) > 1 or special
    if style_loss and not allow_style_loss:
        raise OfficeKitError(
            "整段改写会合并多种字符格式/域/链接；检查后设置 allow_style_loss=true"
        )
    before = paragraph.text
    if runs:
        runs[0].text = text
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        run = paragraph.add_run()
        run.text = text
    return before, style_loss


def _set_paragraph_text(operation: dict[str, Any], presentation: Any) -> dict[str, Any]:
    text = operation.get("text")
    if not isinstance(text, str):
        raise OfficeKitError("set_paragraph_text 需要字符串 text")
    _, text_frame, _, key = _surface(operation, presentation)
    index = _validate_paragraph_index(operation, text_frame)
    if index is None:
        raise OfficeKitError("set_paragraph_text 必须指定 1-based paragraph")
    paragraph = text_frame.paragraphs[index]
    before, style_loss = _set_paragraph_object(
        paragraph, text, bool(operation.get("allow_style_loss", False))
    )
    return {
        "key": key, "paragraph": index + 1, "before": before, "after": paragraph.text,
        "style_loss": style_loss,
    }


def _set_table_cell(operation: dict[str, Any], presentation: Any) -> dict[str, Any]:
    text = operation.get("text")
    if not isinstance(text, str):
        raise OfficeKitError("set_table_cell 需要字符串 text")
    if "row" not in operation or "column" not in operation:
        raise OfficeKitError("set_table_cell 需要 row/column")
    _, text_frame, _, key = _surface(operation, presentation)
    if "\n" in text:
        lines = text.split("\n")
        paragraphs = list(text_frame.paragraphs)
        while len(paragraphs) < len(lines):
            paragraphs.append(text_frame.add_paragraph())
        for index, line in enumerate(lines):
            _set_paragraph_object(
                text_frame.paragraphs[index], line,
                bool(operation.get("allow_style_loss", False)),
            )
        for paragraph in list(text_frame.paragraphs)[len(lines):]:
            paragraph._p.getparent().remove(paragraph._p)
        after = text_frame.text
    else:
        index = _validate_paragraph_index(operation, text_frame) or 0
        _set_paragraph_object(
            text_frame.paragraphs[index], text,
            bool(operation.get("allow_style_loss", False)),
        )
        for paragraph in list(text_frame.paragraphs)[1:]:
            paragraph._p.getparent().remove(paragraph._p)
        after = text_frame.text
    return {"key": key, "after": after}


def _set_font(operation: dict[str, Any], presentation: Any) -> dict[str, Any]:
    _, text_frame, _, key = _surface(operation, presentation)
    paragraph_index = _validate_paragraph_index(operation, text_frame)
    paragraphs = [text_frame.paragraphs[paragraph_index]] if paragraph_index is not None else list(text_frame.paragraphs)
    match = operation.get("match")
    if match is not None and not isinstance(match, str):
        raise OfficeKitError("set_font.match 必须为字符串")
    properties = {}
    if "font_name" in operation:
        if not isinstance(operation["font_name"], str) or not operation["font_name"]:
            raise OfficeKitError("font_name 必须为非空字符串")
        properties["name"] = operation["font_name"]
    if "size_pt" in operation:
        size = operation["size_pt"]
        if not isinstance(size, (int, float)) or not 4 <= float(size) <= 200:
            raise OfficeKitError("size_pt 必须在 4-200 之间")
        properties["size"] = _pptx().util.Pt(float(size))
    for flag in ("bold", "italic"):
        if flag in operation:
            if not isinstance(operation[flag], bool):
                raise OfficeKitError(f"{flag} 必须为布尔值")
            properties[flag] = operation[flag]
    if "color_hex" in operation:
        properties["color"] = _rgb_color(operation["color_hex"])
    if not properties:
        raise OfficeKitError("set_font 至少指定一个字体属性")
    touched = 0
    for paragraph in paragraphs:
        for run in paragraph.runs:
            if match is not None and match not in run.text:
                continue
            for attribute, value in properties.items():
                setattr(run.font, attribute, value)
            touched += 1
    if touched == 0:
        raise OfficeKitError("set_font 未命中任何 run")
    return {"key": key, "runs_touched": touched}


def _set_geometry(operation: dict[str, Any], presentation: Any) -> dict[str, Any]:
    slide, shape, path = _shape_at(presentation, operation)
    if path != "root":
        raise OfficeKitError("set_geometry 只支持顶层形状，不调整组合内部坐标")
    changed = {}
    for field, attribute in (("left_in", "left"), ("top_in", "top"), ("width_in", "width"), ("height_in", "height")):
        if field not in operation:
            continue
        value = operation[field]
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            raise OfficeKitError(f"{field} 必须为数字")
        if field in {"width_in", "height_in"} and float(value) < 0.05:
            raise OfficeKitError(f"{field} 不得小于 0.05 英寸")
        setattr(shape, attribute, _pptx().util.Inches(float(value)))
        changed[field] = round(float(value), 4)
    if not changed:
        raise OfficeKitError("set_geometry 至少指定一个位置或尺寸")
    if not operation.get("allow_out_of_bounds", False):
        left, top = shape.left, shape.top
        right, bottom = left + shape.width, top + shape.height
        if left < 0 or top < 0 or right > presentation.slide_width or bottom > presentation.slide_height:
            raise OfficeKitError("新几何超出幻灯片；确认设计需要后设置 allow_out_of_bounds=true")
    return {"slide": operation["slide"], "shape_id": operation["shape_id"], "geometry": changed}


def _set_notes(operation: dict[str, Any], presentation: Any) -> dict[str, Any]:
    slide = _slide_at(presentation, operation.get("slide"))
    text = operation.get("text")
    if not isinstance(text, str):
        raise OfficeKitError("set_notes 需要字符串 text")
    if "\r" in text:
        raise OfficeKitError("set_notes 只接受 \\n 分行")
    text_frame = slide.notes_slide.notes_text_frame
    lines = text.split("\n") if text else [""]
    while len(text_frame.paragraphs) < len(lines):
        text_frame.add_paragraph()
    for index, line in enumerate(lines):
        _set_paragraph_object(text_frame.paragraphs[index], line, True)
    for paragraph in list(text_frame.paragraphs)[len(lines):]:
        paragraph._p.getparent().remove(paragraph._p)
    return {"slide": operation["slide"], "after": text_frame.text}


def _apply_operation(operation: dict[str, Any], presentation: Any) -> dict[str, Any]:
    op_type = operation["op"]
    try:
        if op_type == "replace_text":
            detail = _replace_text(operation, presentation)
        elif op_type == "set_paragraph_text":
            detail = _set_paragraph_text(operation, presentation)
        elif op_type == "set_table_cell":
            detail = _set_table_cell(operation, presentation)
        elif op_type == "set_font":
            detail = _set_font(operation, presentation)
        elif op_type == "set_geometry":
            detail = _set_geometry(operation, presentation)
        elif op_type == "set_notes":
            detail = _set_notes(operation, presentation)
        else:
            raise OfficeKitError(f"不支持的操作 {op_type}")
    except OfficeKitError as exc:
        raise OfficeKitError(f"{operation.get('id')}: {exc}") from exc
    return {"id": operation["id"], "op": op_type, **detail}


def _package_media(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as archive:
        return {
            name: hashlib.sha256(archive.read(name)).hexdigest()
            for name in archive.namelist()
            if name.startswith("ppt/media/")
        }


def _execute(plan: dict[str, Any]) -> dict[str, Any]:
    source = plan["source"]
    _validate_input(source, for_edit=True)
    Presentation = _pptx().Presentation
    presentation = Presentation(str(source))
    original_slide_count = len(presentation.slides)
    original_width = int(presentation.slide_width)
    original_height = int(presentation.slide_height)
    source_shapes = [_slide_shapes_snapshot(slide) for slide in presentation.slides]
    source_notes = [_notes_snapshot(slide) for slide in presentation.slides]
    target_shape_ids: dict[int, set[int]] = {}
    target_note_slides: set[int] = set()
    operation_results = []
    for operation in plan["operations"]:
        op_type = operation["op"]
        if op_type in TEXT_OPS | {"set_geometry"}:
            slide = _slide_at(presentation, operation.get("slide"))
            shape_id = operation.get("shape_id")
            if not isinstance(shape_id, int):
                raise OfficeKitError(f"{operation['id']}: shape_id 必须为整数")
            target_shape_ids.setdefault(slide.slide_id, set()).add(shape_id)
        elif op_type == "set_notes":
            slide = _slide_at(presentation, operation.get("slide"))
            target_note_slides.add(slide.slide_id)
        operation_results.append(_apply_operation(operation, presentation))
    if len(presentation.slides) != original_slide_count:
        raise OfficeKitError("操作改变了幻灯片数量；本工具不允许增删幻灯片")
    if int(presentation.slide_width) != original_width or int(presentation.slide_height) != original_height:
        raise OfficeKitError("操作改变了幻灯片尺寸")
    target_snapshots: dict[str, Any] = {}
    for slide in presentation.slides:
        ids = target_shape_ids.get(slide.slide_id, set())
        for shape, path in _iter_shapes(slide.shapes):
            if shape.shape_id in ids:
                target_snapshots[f"{slide.slide_id}:{shape.shape_id}"] = _shape_snapshot(shape, path)
        if slide.slide_id in target_note_slides:
            target_snapshots[f"{slide.slide_id}:notes"] = _notes_snapshot(slide, create=True)
    return {
        "presentation": presentation,
        "operation_results": operation_results,
        "source_shapes": source_shapes,
        "source_notes": source_notes,
        "target_snapshots": target_snapshots,
        "target_shape_ids": {str(key): sorted(value) for key, value in target_shape_ids.items()},
        "target_note_slides": sorted(target_note_slides),
        "source_media": _package_media(source),
        "slide_count": original_slide_count,
        "slide_width": original_width,
        "slide_height": original_height,
    }


def _verify_saved(plan: dict[str, Any], execution: dict[str, Any], output: Path) -> dict[str, Any]:
    _validate_input(output, for_edit=True)
    with zipfile.ZipFile(output) as archive:
        if archive.testzip() is not None:
            raise OfficeKitError("输出 ZIP 校验失败")
    Presentation = _pptx().Presentation
    presentation = Presentation(str(output))
    if len(presentation.slides) != execution["slide_count"]:
        raise OfficeKitError("输出幻灯片数量不一致")
    if int(presentation.slide_width) != execution["slide_width"] or int(presentation.slide_height) != execution["slide_height"]:
        raise OfficeKitError("输出幻灯片尺寸不一致")
    changed = []
    for slide_number, slide in enumerate(presentation.slides, 1):
        shapes = _slide_shapes_snapshot(slide)
        source = execution["source_shapes"][slide_number - 1]
        target_ids = set(execution["target_shape_ids"].get(str(slide.slide_id), []))
        for key, source_fingerprint in source.items():
            shape_id = int(key.split(":", 1)[1])
            if shape_id in target_ids:
                continue
            if shapes.get(key) != source_fingerprint:
                raise OfficeKitError(f"未授权形状发生变化：slide {slide_number}, shape {shape_id}")
        notes_source = execution["source_notes"][slide_number - 1]
        if slide.slide_id not in execution["target_note_slides"] and _notes_snapshot(slide) != notes_source:
            raise OfficeKitError(f"未授权的备注变化：slide {slide_number}")
    for key, expected in execution["target_snapshots"].items():
        slide_part, selector = key.split(":", 1)
        slide_id = int(slide_part)
        slide = next((item for item in presentation.slides if item.slide_id == slide_id), None)
        if slide is None:
            raise OfficeKitError(f"目标页消失：{slide_id}")
        if selector == "notes":
            actual = _notes_snapshot(slide)
        else:
            shape_id = int(selector)
            shape, path = _find_shape(slide, shape_id)
            actual = _shape_snapshot(shape, path)
        if actual != expected:
            raise OfficeKitError(f"目标内容与预演结果不一致：{key}")
        changed.append(key)
    if _package_media(output) != execution["source_media"]:
        raise OfficeKitError("媒体文件集合或内容发生变化；本工具不应改动图片/音视频")
    return {
        "slide_count": len(presentation.slides),
        "media_preserved": True,
        "targets_verified": changed,
        "output_sha256": sha256_file(output),
        "output_size": output.stat().st_size,
    }


def _save_atomic(presentation: Any, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.tmp-{os.getpid()}")
    try:
        presentation.save(str(temporary))
        if not temporary.is_file() or temporary.stat().st_size == 0:
            raise OfficeKitError("保存后文件为空")
        os.replace(temporary, output)
    finally:
        if temporary.exists():
            temporary.unlink()


def _render_summary(output: Path, render_dir: Path) -> dict[str, Any]:
    import subprocess
    import sys

    completed = subprocess.run(
        [sys.executable, str(Path(__file__).resolve().parent / "office_kit.py"),
         "render", str(output), "--out", str(render_dir), "--no-cache"],
        capture_output=True, text=True, check=False,
    )
    if completed.returncode != 0:
        return {"ok": False, "error": (completed.stderr or completed.stdout)[-3000:]}
    try:
        data = json.loads(completed.stdout)
    except json.JSONDecodeError:
        return {"ok": False, "error": "渲染输出不是 JSON"}
    entry = data.get("entries", [{}])[0] if data.get("entries") else {}
    return {
        "ok": True,
        "pdf": entry.get("pdf"),
        "pages": len(entry.get("images", [])),
        "contact_sheets": entry.get("contact_sheets", []),
    }


def apply_plan(
    plan_path: Path,
    *,
    commit: bool,
    overwrite: bool,
    result_path: Path | None = None,
    render_dir: Path | None = None,
) -> dict[str, Any]:
    plan = _load_plan(plan_path)
    source = plan["source"]
    output = plan["output"]
    if commit and output.exists() and not overwrite:
        raise OfficeKitError(f"输出已存在；确认后可加 --overwrite：{output}")
    if not commit and overwrite:
        raise OfficeKitError("--overwrite 只能与 --commit 同用")
    execution = _execute(plan)
    presentation = execution["presentation"]
    if commit:
        if sha256_file(source) != plan["source_sha256"]:
            raise OfficeKitError("提交前源文件发生变化；重新检查并生成计划")
        _save_atomic(presentation, output)
        try:
            if sha256_file(source) != plan["source_sha256"]:
                raise OfficeKitError("提交过程中源文件发生变化")
            verified = _verify_saved(plan, execution, output)
        except Exception:
            if output.is_file():
                output.unlink()
            raise
        result = {
            "schema": RESULT_SCHEMA,
            "ok": True,
            "commit": True,
            "source": str(source),
            "source_sha256": plan["source_sha256"],
            "output": str(output),
            "output_sha256": verified["output_sha256"],
            "output_size": verified["output_size"],
            "slide_count": verified["slide_count"],
            "media_preserved": verified["media_preserved"],
            "targets_verified": verified["targets_verified"],
            "operations": execution["operation_results"],
            "warnings": [],
        }
        if any(item.get("cross_run") for item in execution["operation_results"]):
            result["warnings"].append("存在跨 run 替换，新文字沿用起始 run 的字符格式")
        if any(item.get("style_loss") for item in execution["operation_results"]):
            result["warnings"].append("存在整段样式合并，须在渲染结果中复核粗体/颜色/链接")
        if render_dir is not None:
            result["render"] = _render_summary(output, render_dir)
    else:
        with tempfile.TemporaryDirectory(prefix="contact-pptx-dryrun-") as temporary:
            temporary_output = Path(temporary) / "preview.pptx"
            presentation.save(str(temporary_output))
            verified = _verify_saved(plan, execution, temporary_output)
        result = {
            "schema": RESULT_SCHEMA,
            "ok": True,
            "commit": False,
            "source": str(source),
            "source_sha256": plan["source_sha256"],
            "output": str(output),
            "preview_sha256": verified["output_sha256"],
            "slide_count": verified["slide_count"],
            "media_preserved": verified["media_preserved"],
            "targets_verified": verified["targets_verified"],
            "operations": execution["operation_results"],
            "warnings": ["当前为预演，未写入目标文件；确认后加 --commit"],
        }
    if result_path is not None:
        result_path = result_path.expanduser().resolve()
        result_path.parent.mkdir(parents=True, exist_ok=True)
        result_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    return result


def verify_output(
    plan_path: Path,
    output: Path,
    *,
    result_path: Path | None = None,
) -> dict[str, Any]:
    plan = _load_plan(plan_path)
    output = output.expanduser().resolve()
    if not output.is_file():
        raise OfficeKitError(f"输出不存在：{output}")
    execution = _execute(plan)
    verified = _verify_saved(plan, execution, output)
    result = {
        "schema": RESULT_SCHEMA,
        "ok": True,
        "source": str(plan["source"]),
        "source_sha256": plan["source_sha256"],
        "output": str(output),
        **verified,
    }
    if result_path is not None:
        result_path = result_path.expanduser().resolve()
        result_path.parent.mkdir(parents=True, exist_ok=True)
        result_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    return result
