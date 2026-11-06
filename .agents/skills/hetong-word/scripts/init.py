#!/usr/bin/env python3
"""Capability-scoped, opt-in initialization. Detection never installs components."""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

import office_kit

HERE = Path(__file__).resolve().parent
GENERATE = ("python-docx", "lxml", "typing_extensions")
RENDER = GENERATE + ("pillow", "pypdf", "pypdfium2")
MODULES = {"python-docx": "docx", "lxml": "lxml", "typing_extensions": "typing_extensions",
           "pillow": "PIL", "pypdf": "pypdf", "pypdfium2": "pypdfium2",
           "defusedxml": "defusedxml", "et_xmlfile": "et_xmlfile", "openpyxl": "openpyxl",
           "python-pptx": "pptx", "xlsxwriter": "xlsxwriter"}


def requirements(profile: str) -> list[str]:
    pins = dict(line.strip().split("==", 1) for line in
                (HERE / "office-requirements.lock").read_text().splitlines()
                if line.strip() and not line.startswith("#"))
    names = () if profile == "core" else GENERATE if profile == "generate" else RENDER if profile == "render" else tuple(pins)
    return [f"{name}=={pins[name]}" for name in names]


def probe(python: str, packages: list[str]) -> dict:
    modules = [MODULES[p.split("==", 1)[0]] for p in packages]
    code = ("import sys,importlib,json\nerrors={}\n"
            "for name in " + repr(modules) + ":\n"
            " try: importlib.import_module(name)\n"
            " except Exception as exc: errors[name]=type(exc).__name__+': '+str(exc)\n"
            "print(json.dumps({'version':list(sys.version_info[:3]),"
            "'missing':list(errors),'import_errors':errors}))")
    try:
        run = subprocess.run([python, "-c", code], capture_output=True, text=True, timeout=20)
        if run.returncode:
            return {"usable": False, "error": run.stderr[-1000:]}
        data = json.loads(run.stdout)
        data["usable"] = data["version"] >= [3, 10]
        return data
    except (OSError, subprocess.SubprocessError, ValueError) as exc:
        return {"usable": False, "error": str(exc)}


def choose_python(args, packages: list[str]) -> tuple[str, dict]:
    if args.python:
        candidates = [args.python]
    elif args.venv:
        candidates = [str(office_kit._venv_python(Path(args.venv).expanduser()))]
    else:
        candidates = [sys.executable]
        existing = office_kit._runtime_python()
        if existing:
            candidates.append(str(existing))
    attempts = [(p, probe(p, packages)) for p in dict.fromkeys(candidates)]
    for item in attempts:
        if item[1]["usable"] and not item[1].get("missing"):
            return item
    for item in attempts:
        if item[1]["usable"]:
            return item
    return attempts[0]


def install_libreoffice() -> list[str]:
    if os.name == "nt" and shutil.which("winget"):
        command = [shutil.which("winget"), "install", "--id", "TheDocumentFoundation.LibreOffice",
                   "--exact", "--source", "winget", "--accept-package-agreements", "--accept-source-agreements"]
    elif sys.platform == "darwin" and shutil.which("brew"):
        command = [shutil.which("brew"), "install", "--cask", "libreoffice"]
    elif sys.platform.startswith("linux") and os.geteuid() == 0 and shutil.which("apt-get"):
        command = [shutil.which("apt-get"), "install", "-y", "libreoffice"]
    elif sys.platform.startswith("linux") and os.geteuid() == 0 and shutil.which("dnf"):
        command = [shutil.which("dnf"), "install", "-y", "libreoffice"]
    else:
        raise RuntimeError("缺少已支持的包管理器或安装权限；请安装 LibreOffice 后重新检测，不自动 sudo 或安装包管理器。")
    subprocess.run(command, check=True, stdout=sys.stderr)
    return command


def initialize(args) -> dict:
    if args.check_only and (args.install_python_deps or args.install_libreoffice):
        raise ValueError("--check-only 不能与安装开关组合")
    if args.install_libreoffice and args.profile not in ("render", "office"):
        raise ValueError("只有 render/office 能力需要 LibreOffice；请显式选择相应 profile")
    packages = requirements(args.profile)
    python, status = choose_python(args, packages)
    actions = []
    if args.install_python_deps and packages and (not status["usable"] or status.get("missing")):
        if args.python:
            raise ValueError("显式 --python 仅用于复用；安装依赖请用 --venv 指定隔离目标")
        base = probe(sys.executable, [])
        if not base["usable"]:
            raise ValueError("创建环境需要 Python 3.10+；请先安装 Python")
        venv = Path(args.venv).expanduser().resolve() if args.venv else office_kit._user_data_root() / "venv"
        if not venv.exists():
            subprocess.run([sys.executable, "-m", "venv", str(venv)], check=True, stdout=sys.stderr)
            actions.append("created_venv")
        python = str(office_kit._venv_python(venv))
        subprocess.run([python, "-m", "pip", "install", "--disable-pip-version-check", *packages],
                       check=True, stdout=sys.stderr)
        actions.append("installed_python_dependencies")
        status = probe(python, packages)
    needs_lo = args.profile in ("render", "office")
    lo = None
    if needs_lo:
        if args.libreoffice:
            lo = Path(args.libreoffice).expanduser().resolve()
            if not lo.is_file():
                raise ValueError("显式 LibreOffice 路径无效，不回退到其他安装")
        else:
            lo = office_kit._lo_path()
        if not lo and args.install_libreoffice:
            install_libreoffice()
            actions.append("installed_libreoffice")
            lo = office_kit._lo_path()
    lo_version = office_kit._lo_version(lo) if lo else None
    missing = list(status.get("missing", []))
    if not status["usable"]:
        missing.insert(0, "Python >= 3.10")
    if needs_lo and not lo_version:
        missing.append("LibreOffice")
    return {"schema": "hetong-word-init/1", "ok": not missing,
            "profile": args.profile, "python": python, "python_probe": status,
            "required_packages": packages, "missing": missing, "actions": actions,
            "libreoffice": {"required": needs_lo, "path": str(lo) if lo else None, "version": lo_version},
            "native_word_wps_roundtrip": "not performed"}


def parser():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--profile", choices=("core", "generate", "render", "office"), default="core")
    p.add_argument("--python")
    p.add_argument("--venv")
    p.add_argument("--libreoffice")
    p.add_argument("--check-only", action="store_true")
    p.add_argument("--install-python-deps", action="store_true")
    p.add_argument("--install-libreoffice", action="store_true")
    return p


def main() -> int:
    try:
        result = initialize(parser().parse_args())
    except (ValueError, RuntimeError, OSError, subprocess.SubprocessError, office_kit.OfficeKitError) as exc:
        result = {"schema": "hetong-word-init/1", "ok": False, "error": str(exc)}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
