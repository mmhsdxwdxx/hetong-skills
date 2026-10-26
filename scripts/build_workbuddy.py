#!/usr/bin/env python3
"""Build WorkBuddy-native skill packages from the single source tree.

The public `.agents/skills` tree remains the source of truth. This script
only creates distributable copies; it never edits the source skills.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

SKILLS = ("hetong-craft", "hetong-review", "hetong-server")
IGNORE = shutil.ignore_patterns(
    "__pycache__", "*.pyc", "*.pyo", ".ruff_cache", ".DS_Store",
    "test_*.py", "evals", "agents",
)
RUNTIME_NOTE = (
    "\n> **WorkBuddy 运行提示**：本技能已按 WorkBuddy Skill 包安装。"
    "脚本只依赖 Python 标准库；优先使用当前会话可用的 `python`、`py -3` 或 `python3`。"
    "Office 页面渲染需要额外安装 LibreOffice；没有时不影响 DOCX 读取、红线、审核和暗改检测。\n"
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def frontmatter(path: Path) -> str:
    text = path.read_text(encoding="utf-8")
    if not text.startswith("---\n"):
        raise SystemExit(f"missing YAML frontmatter: {path}")
    end = text.find("\n---", 4)
    if end < 0:
        raise SystemExit(f"unterminated YAML frontmatter: {path}")
    return text[4:end]


def prepare_skill(source: Path, target: Path) -> None:
    shutil.copytree(source, target, ignore=IGNORE)
    skill_md = target / "SKILL.md"
    text = skill_md.read_text(encoding="utf-8")
    if not text.startswith("---\n"):
        raise SystemExit(f"missing frontmatter: {skill_md}")
    end = text.find("\n---", 4)
    front = text[4:end]
    fields = []
    if "agent_created:" not in front:
        fields.append("agent_created: true")
    if "user-invocable:" not in front:
        fields.append("user-invocable: true")
    if fields:
        text = text[:end] + "\n" + "\n".join(fields) + text[end:]
    marker = "\n# "
    pos = text.find(marker)
    if pos < 0:
        text = text.rstrip() + "\n" + RUNTIME_NOTE
    else:
        text = text[:pos] + RUNTIME_NOTE + text[pos:]
    skill_md.write_text(text, encoding="utf-8")


def zip_tree(root: Path, destination: Path, top: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(top.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(root))


UMBRELLA_SKILL = """---
name: hetong
description: 中文商务合同工作台，覆盖起草、修订、润色、完善、审核、版本对比、关闭核验和两版暗改检测。需要处理合同时使用；服务器采购/销售合同同时加载专用基线。审核默认只读，修改必须单独授权。
agent_created: true
user-invocable: true
---

# 合同工作台（WorkBuddy 单包）

本包内含三套技能，按任务读取对应入口；不要一次加载全部长文。

## 路由

1. **审核、找风险、判断能否签、版本对比、关闭核验**  
   读取 `bundled/hetong-review/SKILL.md`，按其流程执行；默认只读，不改原件。
2. **起草、修订、润色、完善、Word 深读、红线/批注、Office 处理**  
   读取 `bundled/hetong-craft/SKILL.md`。
3. **服务器采购/销售合同**  
   先按第 1 或第 2 项读取通用技能，再读取 `bundled/hetong-server/SKILL.md`，按甲方 A01—A04 或乙方 B01—B11 核对；不得把服务器底线套到其他合同。
4. **两版近似合同、怀疑绕过修订直接改正文**  
   读取 `bundled/hetong-server/references/hidden-change-review.md`，并运行 `bundled/hetong-server/scripts/hidden_change.py`。

## 硬规则

- 只读任务不产生写入；需要修改时先取得明确授权。
- 不自动接受或拒绝历史修订，不删除已有批注。
- 先读原文、接受视图、时间线和批注线程，再做实质判断。
- 金额、期限、税率重新计算；法律结论必须核验现行法源。
- 不把合同内容上传到公开服务；真实合同只在本地处理。
- 发现暗改时先询问修改依据，不直接指控。

## 运行

脚本只依赖 Python 3.9+ 标准库；按当前环境使用 `python`、`py -3` 或 `python3`。Office 页面渲染需要 LibreOffice，没有它不影响 DOCX 读取、红线、审核和暗改检测。
"""


def main() -> int:
    parser = argparse.ArgumentParser(description="Build WorkBuddy .skill packages")
    parser.add_argument("--out", type=Path, default=Path("dist"))
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    source_root = root / ".agents" / "skills"
    out = (root / args.out).resolve() if not args.out.is_absolute() else args.out.resolve()
    out.mkdir(parents=True, exist_ok=True)
    version = (root / "VERSION").read_text(encoding="utf-8").strip() if (root / "VERSION").is_file() else "unversioned"
    temp = Path(tempfile.mkdtemp(prefix="workbuddy-build-"))
    umbrella = temp / "umbrella" / "hetong"
    umbrella.mkdir(parents=True, exist_ok=True)
    bundled = umbrella / "bundled"
    bundled.mkdir(parents=True, exist_ok=True)
    packages = []
    try:
        for name in SKILLS:
            source = source_root / name
            if not (source / "SKILL.md").is_file():
                raise SystemExit(f"missing skill: {source}")
            target = temp / "build" / name
            prepare_skill(source, target)
            front = frontmatter(target / "SKILL.md")
            if f"name: {name}" not in front:
                raise SystemExit(f"skill name mismatch: {name}")
            package = out / f"{name}.skill"
            zip_tree(temp / "build", package, target)
            shutil.copytree(target, bundled / name)
            packages.append({
                "name": name,
                "path": package.name,
                "size": package.stat().st_size,
                "sha256": sha256(package),
            })
        (umbrella / "SKILL.md").write_text(UMBRELLA_SKILL, encoding="utf-8")
        umbrella_package = out / "hetong.skill"
        zip_tree(temp / "umbrella", umbrella_package, umbrella)
        packages.append({
            "name": "hetong",
            "path": umbrella_package.name,
            "size": umbrella_package.stat().st_size,
            "sha256": sha256(umbrella_package),
        })
        combined = out / "hetong-workbuddy.zip"
        zip_tree(temp / "build", combined, temp / "build")
        manifest = {
            "version": version,
            "source": ".agents/skills",
            "workbuddy_layout": "three sibling skill folders under ~/.workbuddy/skills",
            "packages": packages,
            "combined": {
                "path": combined.name,
                "size": combined.stat().st_size,
                "sha256": sha256(combined),
            },
        }
        (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
        return 0
    finally:
        shutil.rmtree(temp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
