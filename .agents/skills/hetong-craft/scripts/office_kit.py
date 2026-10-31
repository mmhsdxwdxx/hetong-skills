#!/usr/bin/env python3
"""仓库级 Office 工具：结构检查、PPTX 精修、批量无头渲染。

本文件自身只依赖 Python 标准库；除 ``setup`` 外的命令会自动切换到隔离的
Office 虚拟环境。核心写入引擎不调用 Word/WPS，LibreOffice 仅用于批量渲染。
"""

from __future__ import annotations

import argparse
import contextlib
import hashlib
import html
import importlib.util
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Any, cast

KIT_VERSION = "1.1.0"
SCHEMA = "contact-office-kit/1"

# Windows 控制台默认编码可能不是 UTF-8；命令行输出统一为 UTF-8。
for _stream in (sys.stdout, sys.stderr):
    try:
        cast(Any, _stream).reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass
SCRIPT_DIR = Path(__file__).resolve().parent
REQUIREMENTS = SCRIPT_DIR / "office-requirements.lock"
REQUIRED_MODULES = ("docx", "pptx", "openpyxl", "lxml", "defusedxml", "PIL", "pypdf", "pypdfium2")
MACRO_SUFFIXES = {".docm", ".dotm", ".pptm", ".potm", ".xlsm", ".xltm"}
OOXML_SUFFIXES = {".docx", ".dotx", ".pptx", ".potx", ".xlsx", ".xltx"}
RENDER_SUFFIXES = OOXML_SUFFIXES | {".pdf"}
MAX_ZIP_ENTRIES = 50_000
MAX_UNCOMPRESSED = 512 * 1024 * 1024


class OfficeKitError(RuntimeError):
    def __init__(self, message: str, code: int = 2):
        super().__init__(message)
        self.code = code


def _module(name: str) -> Any:
    import importlib

    try:
        return importlib.import_module(name)
    except ImportError as exc:
        raise OfficeKitError(
            f"缺少 Python 模块 {name}；运行 {Path(__file__)} setup 修复隔离环境",
            code=2,
        ) from exc


def _json_dump(data: Any, path: Path | None = None) -> None:
    text = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
    if path is None:
        print(text, end="")
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp-{os.getpid()}")
    temporary.write_text(text, encoding="utf-8")
    os.replace(temporary, path)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def safe_stem(name: str, limit: int = 48) -> str:
    cleaned = re.sub(r"[^0-9A-Za-z._\-\u4e00-\u9fff]+", "_", name).strip("._")
    return (cleaned or "document")[:limit]


def _is_windows() -> bool:
    return os.name == "nt"


def _user_data_root() -> Path:
    if _is_windows():
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        if base:
            return Path(base) / "contact-office"
        return Path.home() / "AppData" / "Local" / "contact-office"
    return Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share")) / "contact-office"


def _user_cache_root() -> Path:
    if _is_windows():
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        if base:
            return Path(base) / "contact-office"
        return Path.home() / "AppData" / "Local" / "contact-office"
    return Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / "contact-office"


def _venv_python(venv: Path) -> Path:
    if _is_windows():
        return venv / "Scripts" / "python.exe"
    return venv / "bin" / "python"


def _libreoffice_hint() -> str:
    if _is_windows():
        return "安装 LibreOffice for Windows 后重试（默认路径 C:\\Program Files\\LibreOffice\\program\\soffice.exe）"
    return "安装 LibreOffice 后重试：brew install --cask libreoffice"


def _runtime_python() -> Path | None:
    candidates: list[Path] = []
    configured = os.environ.get("CONTACT_OFFICE_VENV")
    if configured:
        configured_path = Path(configured).expanduser()
        candidates.append(configured_path if configured_path.is_file() else _venv_python(configured_path))
    candidates.append(_venv_python(_user_data_root() / "venv"))
    candidates.append(_venv_python(SCRIPT_DIR.parent / ".office-venv"))
    for candidate in candidates:
        try:
            usable = candidate.is_file() and (_is_windows() or os.access(candidate, os.X_OK))
        except OSError:
            usable = False
        if usable:
            return candidate.absolute()
    return None


def _runtime_ready() -> bool:
    return all(importlib.util.find_spec(module) is not None for module in REQUIRED_MODULES)


def ensure_runtime(argv: list[str]) -> None:
    if len(argv) > 0 and argv[0] == "setup":
        return
    if _runtime_ready():
        return
    configured = _runtime_python()
    if configured is not None and Path(sys.executable).absolute() != configured:
        environment = os.environ.copy()
        environment["CONTACT_OFFICE_REEXEC"] = "1"
        completed = subprocess.run(
            [str(configured), str(Path(__file__).resolve()), *argv],
            env=environment,
            check=False,
        )
        raise SystemExit(completed.returncode)
    raise OfficeKitError(
        "Office Python 运行环境不完整。先运行："
        + ("python " if _is_windows() else "python3 ")
        + f"{Path(__file__).resolve()} setup",
        code=2,
    )


def _package_versions() -> dict[str, str | None]:
    from importlib import metadata

    mapping = {
        "python-docx": "python-docx",
        "python-pptx": "python-pptx",
        "openpyxl": "openpyxl",
        "lxml": "lxml",
        "defusedxml": "defusedxml",
        "Pillow": "Pillow",
        "pypdf": "pypdf",
        "pypdfium2": "pypdfium2",
    }
    result: dict[str, str | None] = {}
    for label, distribution in mapping.items():
        try:
            result[label] = metadata.version(distribution)
        except metadata.PackageNotFoundError:
            result[label] = None
    return result


def _lo_path() -> Path | None:
    configured = os.environ.get("CONTACT_LIBREOFFICE")
    if configured:
        path = Path(configured).expanduser()
        if not path.is_file():
            raise OfficeKitError("CONTACT_LIBREOFFICE 必须指向已存在的 LibreOffice 可执行文件")
        return path.resolve()
    candidates: list[str | None] = [shutil.which("soffice"), shutil.which("soffice.exe")]
    if _is_windows():
        for env_name in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
            base = os.environ.get(env_name)
            if base:
                candidates.append(str(Path(base) / "LibreOffice" / "program" / "soffice.exe"))
    else:
        candidates.extend([
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/usr/bin/libreoffice",
            "/usr/local/bin/libreoffice",
        ])
    for candidate in candidates:
        if candidate:
            path = Path(candidate)
            try:
                usable = path.is_file() and (_is_windows() or os.access(path, os.X_OK))
            except OSError:
                usable = False
            if usable:
                return path
    return None


def _lo_version(path: Path | None = None) -> str | None:
    binary = path or _lo_path()
    if binary is None:
        return None
    try:
        completed = subprocess.run(
            [str(binary), "--version"],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    text = (completed.stdout or completed.stderr).strip()
    return text or None


def setup_runtime(args: argparse.Namespace) -> int:
    venv = Path(args.venv).expanduser().resolve() if args.venv else (_user_data_root() / "venv")
    if not REQUIREMENTS.is_file():
        raise OfficeKitError(f"缺少版本锁定文件：{REQUIREMENTS}")
    if not venv.exists():
        import venv as venv_module

        venv_module.EnvBuilder(with_pip=True, clear=False).create(venv)
    python = _venv_python(venv)
    if not python.is_file():
        raise OfficeKitError(f"虚拟环境创建失败：{python}")
    command = [str(python), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(REQUIREMENTS)]
    if args.index_url:
        command.extend(["--index-url", args.index_url])
    completed = subprocess.run(command, check=False)
    if completed.returncode != 0:
        raise OfficeKitError("Python Office 依赖安装失败；检查网络或使用 --index-url", code=2)
    probe = subprocess.run(
        [str(python), "-c", "import docx,pptx,openpyxl,lxml,defusedxml,PIL,pypdf,pypdfium2"],
        capture_output=True,
        text=True,
        check=False,
    )
    if probe.returncode != 0:
        raise OfficeKitError(f"安装后导入失败：{probe.stderr.strip()}", code=2)
    version_probe = subprocess.run(
        [str(python), "-c", (
            "import importlib.metadata as m,json; "
            "print(json.dumps({n:m.version(n) for n in "
            "['python-docx','python-pptx','openpyxl','lxml','defusedxml','Pillow','pypdf','pypdfium2']}))"
        )],
        capture_output=True,
        text=True,
        check=False,
    )
    try:
        packages = json.loads(version_probe.stdout)
    except json.JSONDecodeError as exc:
        raise OfficeKitError(f"无法读取安装版本：{version_probe.stderr.strip()}") from exc
    _json_dump({
        "ok": True,
        "venv": str(venv),
        "python": str(python),
        "packages": packages,
        "next": "安装 LibreOffice 后运行 doctor --full：" + _libreoffice_hint(),
    })
    return 0


def _zip_inventory(path: Path) -> dict[str, Any]:
    if path.suffix.lower() in MACRO_SUFFIXES:
        raise OfficeKitError(f"拒绝宏/可执行文档：{path.name}；先另存为无宏标准 OOXML")
    try:
        archive = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as exc:
        raise OfficeKitError(f"不是有效的 OOXML/ZIP 文件：{path.name}（{exc}）") from exc
    with archive:
        infos = archive.infolist()
        if len(infos) > MAX_ZIP_ENTRIES:
            raise OfficeKitError(f"ZIP 条目过多：{len(infos)}")
        total = 0
        external = 0
        macros = False
        for info in infos:
            name = info.filename
            parts = Path(name.replace("\\", "/")).parts
            if name.startswith("/") or ".." in parts or "\\" in name or re.match(r"^[A-Za-z]:", name):
                raise OfficeKitError(f"ZIP 路径不安全：{name}")
            mode = info.external_attr >> 16
            if stat.S_ISLNK(mode):
                raise OfficeKitError(f"拒绝 ZIP 符号链接：{name}")
            if info.flag_bits & 0x1:
                raise OfficeKitError("拒绝加密 ZIP 条目")
            total += info.file_size
            if total > MAX_UNCOMPRESSED:
                raise OfficeKitError(f"解压总量超过限制：{MAX_UNCOMPRESSED} 字节")
            if (
                info.file_size > 1024 * 1024
                and info.compress_size > 0
                and info.file_size / info.compress_size > 250
            ):
                raise OfficeKitError(f"ZIP 压缩比异常：{name}")
            if name.lower().endswith("vbaproject.bin"):
                macros = True
        if macros:
            raise OfficeKitError("检测到 VBA 宏工程；拒绝编辑或渲染")
        if "[Content_Types].xml" in archive.namelist():
            content = archive.read("[Content_Types].xml").decode("utf-8", errors="ignore")
            if re.search(r"macroEnabled", content, re.I):
                raise OfficeKitError("Content Types 标记为 macroEnabled；拒绝处理")
        for name in archive.namelist():
            if name.lower().endswith(".rels"):
                text = archive.read(name).decode("utf-8", errors="ignore")
                external += len(re.findall(r'TargetMode\s*=\s*"External"', text, re.I))
    return {
        "entries": len(infos),
        "uncompressed_bytes": total,
        "external_relationships": external,
    }


def _validate_input(path: Path, *, for_edit: bool = False) -> dict[str, Any]:
    suffix = path.suffix.lower()
    if suffix in MACRO_SUFFIXES:
        raise OfficeKitError(f"拒绝宏/可执行文档：{path.name}")
    if suffix in OOXML_SUFFIXES:
        inventory = _zip_inventory(path)
        if for_edit and suffix not in {".pptx", ".potx"}:
            raise OfficeKitError("PPTX 精修仅接受 .pptx/.potx")
        return inventory
    if suffix == ".pdf" and not for_edit:
        PdfReader = _module("pypdf").PdfReader

        try:
            reader = PdfReader(str(path))
            if reader.is_encrypted:
                raise OfficeKitError(f"PDF 已加密，需先获授权解密：{path.name}")
            if len(reader.pages) == 0:
                raise OfficeKitError(f"PDF 没有页面：{path.name}")
        except OfficeKitError:
            raise
        except Exception as exc:
            raise OfficeKitError(f"无法读取 PDF：{path.name}（{exc}）") from exc
        return {"pages": len(reader.pages), "external_relationships": 0}
    if for_edit:
        raise OfficeKitError("精修仅接受 .pptx/.potx")
    raise OfficeKitError(
        f"不支持的输入类型：{suffix or '无后缀'}。请先另存为 "
        ".docx/.pptx/.xlsx 或 PDF；旧版 .doc/.ppt/.xls 不直接处理"
    )


def _parse_pages(spec: str | None, count: int) -> list[int]:
    if not spec:
        return list(range(1, count + 1))
    selected: list[int] = []
    for token in spec.split(","):
        token = token.strip()
        if not token:
            continue
        if "-" in token:
            left, right = token.split("-", 1)
            start, end = int(left), int(right)
            if start > end:
                raise OfficeKitError(f"页码范围倒置：{token}")
            selected.extend(range(start, end + 1))
        else:
            selected.append(int(token))
    result = list(dict.fromkeys(selected))
    if not result or any(page < 1 or page > count for page in result):
        raise OfficeKitError(f"页码范围超出 1-{count}：{spec}")
    return result


def _flock(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    if _is_windows():
        import importlib

        msvcrt = cast(Any, importlib.import_module("msvcrt"))
        handle = path.open("a+b")
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"\0")
            handle.flush()
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_LOCK, 1)
        return handle
    import fcntl

    handle = path.open("a+")
    fcntl.flock(handle.fileno(), fcntl.LOCK_EX)
    return handle


def _lo_profile() -> Path:
    profile = _user_cache_root() / "libreoffice-profile"
    profile.mkdir(parents=True, exist_ok=True, mode=0o700)
    return profile


def _font_directories() -> list[Path]:
    home = Path.home()
    if _is_windows():
        candidates = [
            Path(os.environ.get("WINDIR", "C:\\Windows")) / "Fonts",
            home / "AppData" / "Local" / "Microsoft" / "Windows" / "Fonts",
        ]
    else:
        candidates = [
            Path("/System/Library/Fonts"),
            Path("/System/Library/Fonts/Supplemental"),
            Path("/Library/Fonts"),
            home / "Library" / "Fonts",
            home / ".local" / "share" / "fonts",
            home / ".fonts",
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
            Path("/Applications/LibreOffice.app/Contents/Resources/fonts/truetype"),
        ]
    return [path for path in candidates if path.is_dir()]


def _font_signature(directories: list[Path]) -> str:
    digest = hashlib.sha256()
    for directory in directories:
        try:
            entries = sorted(directory.rglob("*"))
        except OSError:
            continue
        for path in entries:
            if path.suffix.lower() not in {".ttf", ".otf", ".ttc", ".otc", ".pfb"}:
                continue
            try:
                stat_info = path.stat()
            except OSError:
                continue
            digest.update(f"{path}:{stat_info.st_size}:{stat_info.st_mtime_ns}".encode())
    return digest.hexdigest()[:24]


def _font_environment() -> tuple[Path | None, str]:
    directories = _font_directories()
    if _is_windows():
        # Windows LibreOffice 使用系统字体，不套用 Unix fontconfig 配置。
        return None, _font_signature(directories)
    data = _user_data_root()
    config_dir = data / "fontconfig"
    cache_dir = _user_cache_root() / "fontconfig-cache"
    config_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
    cache_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
    lines = [
        '<?xml version="1.0"?>',
        '<!DOCTYPE fontconfig SYSTEM "fonts.dtd">',
        '<fontconfig>',
        *[f"  <dir>{html.escape(str(path))}</dir>" for path in directories],
        f"  <cachedir>{html.escape(str(cache_dir))}</cachedir>",
        '</fontconfig>',
    ]
    config = config_dir / "fonts.conf"
    content = "\n".join(lines) + "\n"
    if not config.is_file() or config.read_text(encoding="utf-8") != content:
        config.write_text(content, encoding="utf-8")
    digest = hashlib.sha256(content.encode())
    digest.update(_font_signature(directories).encode())
    return config, digest.hexdigest()[:24]


def _convert_batch(inputs: list[Path], out_dir: Path, timeout: int) -> dict[str, Any]:
    binary = _lo_path()
    if binary is None:
        raise OfficeKitError(f"未找到 LibreOffice。{_libreoffice_hint()}")
    out_dir.mkdir(parents=True, exist_ok=True)
    profile = _lo_profile()
    font_config, font_signature = _font_environment()
    lock = _flock(profile.parent / "libreoffice.lock")
    started = time.monotonic()
    try:
        command = [
            str(binary), "--headless", "--nologo", "--nodefault", "--nofirststartwizard",
            "--norestore", "--nolockcheck", f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--convert-to", "pdf", "--outdir", str(out_dir), *[str(path) for path in inputs],
        ]
        environment = os.environ.copy()
        if _is_windows():
            # A copied Unix setting must not override Windows system font discovery.
            environment.pop("FONTCONFIG_FILE", None)
            environment.pop("FONTCONFIG_PATH", None)
        if font_config is not None:
            environment["FONTCONFIG_FILE"] = str(font_config)
        if sys.platform.startswith("linux"):
            environment.setdefault("SAL_USE_VCLPLUGIN", "svp")
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if _is_windows() else 0
        completed = subprocess.run(
            command, capture_output=True, text=True, timeout=timeout,
            check=False, env=environment, creationflags=creationflags,
        )
    except subprocess.TimeoutExpired as exc:
        raise OfficeKitError(f"LibreOffice 转换超时（{timeout}s）") from exc
    finally:
        lock.close()
    duration = round(time.monotonic() - started, 3)
    expected = {path.stem: out_dir / f"{path.stem}.pdf" for path in inputs}
    missing = [str(source) for source, target in expected.items() if not target.is_file()]
    if completed.returncode != 0 or missing:
        tail = ((completed.stdout or "") + "\n" + (completed.stderr or ""))[-4000:].strip()
        raise OfficeKitError(
            f"LibreOffice 转换失败（exit={completed.returncode}, missing={missing}）：{tail}"
        )
    return {
        "duration_seconds": duration,
        "returncode": completed.returncode,
        "stdout": (completed.stdout or "")[-4000:],
        "stderr": (completed.stderr or "")[-4000:],
        "font_signature": font_signature,
    }


def _pdf_metadata(pdf_path: Path) -> dict[str, Any]:
    PdfReader = _module("pypdf").PdfReader

    reader = PdfReader(str(pdf_path))
    pages = []
    for index, page in enumerate(reader.pages, 1):
        box = page.mediabox
        pages.append({
            "page": index,
            "width_pt": round(float(box.width), 3),
            "height_pt": round(float(box.height), 3),
        })
    return {"page_count": len(reader.pages), "pages": pages}


def _render_pdf_images(
    pdf_path: Path,
    target_dir: Path,
    dpi: int,
    pages: list[int],
) -> tuple[list[Path], list[Path], list[dict[str, Any]]]:
    pdfium = _module("pypdfium2")
    Image = _module("PIL.Image")
    ImageDraw = _module("PIL.ImageDraw")
    ImageOps = _module("PIL.ImageOps")

    document = pdfium.PdfDocument(str(pdf_path))
    try:
        if len(document) == 0:
            raise OfficeKitError("PDF 没有页面")
        selected = _parse_pages(",".join(str(page) for page in pages), len(document))
        images: list[Path] = []
        page_info: list[dict[str, Any]] = []
        scale = dpi / 72.0
        for page_number in selected:
            page = document[page_number - 1]
            width_pt, height_pt = page.get_size()
            bitmap = page.render(scale=scale)
            image = bitmap.to_pil()
            target = target_dir / f"page-{page_number:03d}.png"
            image.save(target, "PNG", optimize=True)
            images.append(target)
            page_info.append({
                "page": page_number,
                "width_px": image.width,
                "height_px": image.height,
                "width_pt": round(width_pt, 3),
                "height_pt": round(height_pt, 3),
            })
    finally:
        document.close()

    sheets: list[Path] = []
    for sheet_index in range(0, len(images), 9):
        chunk = images[sheet_index:sheet_index + 9]
        cell_width, cell_height, margin, label_height = 420, 330, 18, 26
        sheet = Image.new("RGB", (cell_width * 3, (cell_height + label_height) * 3), "#e8ebef")
        draw = ImageDraw.Draw(sheet)
        for local_index, image_path in enumerate(chunk):
            row, column = divmod(local_index, 3)
            with Image.open(image_path) as source:
                thumb = ImageOps.contain(source.convert("RGB"), (cell_width - 2 * margin, cell_height - 2 * margin))
            x = column * cell_width + (cell_width - thumb.width) // 2
            y = row * (cell_height + label_height) + label_height + (cell_height - thumb.height) // 2
            sheet.paste(thumb, (x, y))
            page_label = selected[sheet_index + local_index]
            draw.text(
                (column * cell_width + margin, row * (cell_height + label_height) + 6),
                f"P.{page_label}", fill="#222222",
            )
        target = target_dir / f"contact-sheet-{len(sheets) + 1:02d}.png"
        sheet.save(target, "PNG", optimize=True)
        sheets.append(target)
    return images, sheets, page_info


def _atomic_copy(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.tmp-{os.getpid()}")
    shutil.copy2(source, temporary)
    os.replace(temporary, target)


def _manifest_entry(
    source: Path,
    source_sha: str,
    output_dir: Path,
    pdf: Path,
    images: list[Path],
    sheets: list[Path],
    pages: list[dict[str, Any]],
    dpi: int,
    page_spec: str | None,
    lo_version: str | None,
    font_signature: str,
    inventory: dict[str, Any],
    duration: float,
    cached: bool,
) -> dict[str, Any]:
    cache_payload = json.dumps({
        "source_sha256": source_sha, "dpi": dpi, "page_spec": page_spec or "all",
        "lo_version": lo_version, "kit_version": KIT_VERSION, "schema": SCHEMA,
        "font_signature": font_signature,
    }, sort_keys=True).encode()
    return {
        "input": str(source),
        "input_sha256": source_sha,
        "kind": source.suffix.lower().lstrip("."),
        "output_dir": str(output_dir),
        "pdf": str(pdf),
        "pdf_sha256": sha256_file(pdf),
        "pages": pages,
        "images": [str(path) for path in images],
        "contact_sheets": [str(path) for path in sheets],
        "image_sha256": {str(path): sha256_file(path) for path in images},
        "contact_sheet_sha256": {str(path): sha256_file(path) for path in sheets},
        "dpi": dpi,
        "page_spec": page_spec,
        "cache_key": sha256_bytes(cache_payload)[:24],
        "libreoffice_version": lo_version,
        "inventory": inventory,
        "duration_seconds": round(duration, 3),
        "cached": cached,
    }


def render_documents(args: argparse.Namespace) -> dict[str, Any]:
    inputs = [Path(item).expanduser().resolve() for item in args.inputs]
    if not inputs:
        raise OfficeKitError("至少提供一个输入文件")
    for path in inputs:
        if not path.is_file():
            raise OfficeKitError(f"输入不存在：{path}")
    if args.dpi < 40 or args.dpi > 300:
        raise OfficeKitError("DPI 必须在 40-300 之间")
    output_root = Path(args.out).expanduser().resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    manifest_path = output_root / "render-manifest.json"
    old_manifest: dict[str, Any] = {}
    if manifest_path.is_file():
        try:
            old_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            old_manifest = {}
    old_by_key = {
        entry.get("cache_key"): entry
        for entry in old_manifest.get("entries", [])
        if isinstance(entry, dict) and entry.get("cache_key")
    }
    lo = _lo_path()
    lo_version = _lo_version(lo)
    _, font_signature = _font_environment()
    started_all = time.monotonic()
    stage = Path(tempfile.mkdtemp(prefix=".office-stage-", dir=output_root))
    render_jobs: list[dict[str, Any]] = []
    entries: list[dict[str, Any]] = []
    warnings: list[str] = []
    try:
        for index, source in enumerate(inputs):
            source_sha = sha256_file(source)
            inventory = _validate_input(source)
            cache_payload = json.dumps({
                "source_sha256": source_sha, "dpi": args.dpi,
                "page_spec": args.pages or "all", "lo_version": lo_version,
                "kit_version": KIT_VERSION, "schema": SCHEMA,
                "font_signature": font_signature,
            }, sort_keys=True).encode()
            cache_key = sha256_bytes(cache_payload)[:24]
            old = old_by_key.get(cache_key)
            if old and not args.no_cache:
                needed = [Path(old.get("pdf", "")), *[Path(p) for p in old.get("images", [])],
                          *[Path(p) for p in old.get("contact_sheets", [])]]
                hashes = {**old.get("image_sha256", {}), **old.get("contact_sheet_sha256", {})}
                cached_files_valid = all(path.is_file() for path in needed)
                if cached_files_valid and old.get("pdf_sha256"):
                    cached_files_valid = sha256_file(Path(old["pdf"])) == old["pdf_sha256"]
                if cached_files_valid:
                    cached_files_valid = all(
                        path.is_file() and hashes.get(str(path)) == sha256_file(path)
                        for path in [*[Path(p) for p in old.get("images", [])],
                                     *[Path(p) for p in old.get("contact_sheets", [])]]
                    )
                if cached_files_valid:
                    cached = dict(old)
                    cached["cached"] = True
                    cached["duration_seconds"] = 0.0
                    entries.append(cached)
                    continue
            staged = stage / f"{index:04d}-{safe_stem(source.stem)}{source.suffix.lower()}"
            shutil.copy2(source, staged)
            render_jobs.append({
                "source": source, "source_sha": source_sha, "staged": staged,
                "inventory": inventory, "cache_key": cache_key,
                "output_dir": output_root / f"{safe_stem(source.stem)}-{source_sha[:8]}",
            })
        if any(job["source"].suffix.lower() != ".pdf" and not lo for job in render_jobs):
            raise OfficeKitError(f"需要 LibreOffice 渲染 Office 文件。{_libreoffice_hint()}")
        conversion_root = stage / "pdf"
        conversion_root.mkdir()
        office_jobs = [job for job in render_jobs if job["source"].suffix.lower() != ".pdf"]
        batch_info: dict[str, Any] = {"duration_seconds": 0.0}
        if office_jobs:
            batch_info = _convert_batch(
                [job["staged"] for job in office_jobs], conversion_root, args.timeout
            )
        for job in render_jobs:
            job_started = time.monotonic()
            source = job["source"]
            if source.suffix.lower() == ".pdf":
                pdf = job["staged"]
            else:
                pdf = conversion_root / f"{job['staged'].stem}.pdf"
            if not pdf.is_file():
                raise OfficeKitError(f"缺少转换结果：{pdf}")
            metadata = _pdf_metadata(pdf)
            selected = _parse_pages(args.pages, metadata["page_count"])
            temp_dir = stage / f"pages-{len(entries):04d}"
            temp_dir.mkdir()
            images, sheets, page_info = _render_pdf_images(pdf, temp_dir, args.dpi, selected)
            output_dir = job["output_dir"]
            output_dir.mkdir(parents=True, exist_ok=True)
            target_pdf = output_dir / "source.pdf"
            _atomic_copy(pdf, target_pdf)
            target_images = []
            for image in images:
                target = output_dir / image.name
                _atomic_copy(image, target)
                target_images.append(target)
            target_sheets = []
            for sheet in sheets:
                target = output_dir / sheet.name
                _atomic_copy(sheet, target)
                target_sheets.append(target)
            for stale in output_dir.glob("page-*.png"):
                if stale.name not in {path.name for path in target_images}:
                    stale.unlink()
            for stale in output_dir.glob("contact-sheet-*.png"):
                if stale.name not in {path.name for path in target_sheets}:
                    stale.unlink()
            duration = time.monotonic() - job_started
            entry = _manifest_entry(
                source, job["source_sha"], output_dir, target_pdf, target_images,
                target_sheets, page_info, args.dpi, args.pages, lo_version,
                font_signature, job["inventory"], duration, False,
            )
            entry["libreoffice_batch_seconds"] = batch_info["duration_seconds"]
            entry["font_signature"] = font_signature
            entry["warnings"] = []
            if job["inventory"].get("external_relationships"):
                warning = (
                    f"{source.name} 含 {job['inventory']['external_relationships']} 个外部关系；"
                    "渲染不会主动更新链接，但交付前须确认外链可信"
                )
                entry["warnings"].append(warning)
                warnings.append(warning)
            (output_dir / "render.json").write_text(
                json.dumps(entry, ensure_ascii=False, indent=2) + "\n"
            )
            entries.append(entry)
        merged = list(old_by_key.values())
        fresh_keys = {entry["cache_key"] for entry in entries}
        merged = [entry for entry in merged if entry.get("cache_key") not in fresh_keys]
        merged.extend(entries)
        merged.sort(key=lambda entry: (entry.get("input", ""), entry.get("cache_key", "")))
        manifest = {
            "schema": SCHEMA,
            "kit_version": KIT_VERSION,
            "libreoffice_version": lo_version,
            "font_signature": font_signature,
            "libreoffice_invocations": 1 if office_jobs else 0,
            "batch_conversion_seconds": batch_info["duration_seconds"],
            "total_duration_seconds": round(time.monotonic() - started_all, 3),
            "entries": merged,
            "warnings": warnings,
        }
        _json_dump(manifest, manifest_path)
        console = dict(manifest)
        console["entries"] = entries
        return console
    finally:
        shutil.rmtree(stage, ignore_errors=True)


def _inspect_docx(path: Path, max_items: int) -> dict[str, Any]:
    Document = _module("docx").Document

    document = Document(str(path))
    paragraphs = [
        {"index": index, "style": paragraph.style.name if paragraph.style else None,
         "text": paragraph.text}
        for index, paragraph in enumerate(document.paragraphs, 1)
    ]
    tables = []
    for table_index, table in enumerate(document.tables, 1):
        rows = []
        for row in table.rows[: min(len(table.rows), max_items)]:
            rows.append([cell.text for cell in row.cells[: min(len(row.cells), 20)]])
        tables.append({
            "index": table_index, "rows": len(table.rows), "columns": len(table.columns),
            "preview": rows,
        })
    sections = []
    for index, section in enumerate(document.sections, 1):
        sections.append({
            "index": index,
            "page_width_in": round(section.page_width.inches, 3) if section.page_width else None,
            "page_height_in": round(section.page_height.inches, 3) if section.page_height else None,
            "left_margin_in": round(section.left_margin.inches, 3) if section.left_margin else None,
            "right_margin_in": round(section.right_margin.inches, 3) if section.right_margin else None,
        })
    return {
        "format": "docx",
        "source_sha256": sha256_file(path),
        "paragraphs": len(paragraphs),
        "tables": len(tables),
        "sections": sections,
        "inline_shapes": len(document.inline_shapes),
        "paragraph_items": paragraphs[:max_items],
        "table_items": tables[:max_items],
        "truncated": len(paragraphs) > max_items or len(tables) > max_items,
    }


def _inspect_xlsx(path: Path, max_items: int) -> dict[str, Any]:
    load_workbook = _module("openpyxl").load_workbook

    workbook = load_workbook(str(path), read_only=False, data_only=False, keep_links=False)
    formulas = 0
    sheets = []
    try:
        for sheet in workbook.worksheets:
            sheet_formulas = 0
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        sheet_formulas += 1
            formulas += sheet_formulas
            sheets.append({
                "name": sheet.title,
                "max_row": sheet.max_row,
                "max_column": sheet.max_column,
                "freeze_panes": sheet.freeze_panes,
                "merged_ranges": [str(item) for item in sheet.merged_cells.ranges],
                "tables": len(sheet.tables),
                "charts": len(sheet._charts),
                "images": len(sheet._images),
                "formulas": sheet_formulas,
            })
    finally:
        workbook.close()
    return {
        "format": "xlsx",
        "source_sha256": sha256_file(path),
        "sheets": sheets[:max_items],
        "formula_count": formulas,
        "defined_names": len(workbook.defined_names),
        "truncated": len(sheets) > max_items,
    }


def _inspect_pdf(path: Path, max_items: int) -> dict[str, Any]:
    PdfReader = _module("pypdf").PdfReader

    reader = PdfReader(str(path))
    pages = []
    for page in reader.pages[:max_items]:
        box = page.mediabox
        pages.append({
            "width_pt": round(float(box.width), 3),
            "height_pt": round(float(box.height), 3),
        })
    return {
        "format": "pdf",
        "source_sha256": sha256_file(path),
        "page_count": len(reader.pages),
        "pages": pages,
        "encrypted": reader.is_encrypted,
        "truncated": len(reader.pages) > max_items,
    }


def inspect_file(path: Path, max_items: int) -> dict[str, Any]:
    suffix = path.suffix.lower()
    if suffix in OOXML_SUFFIXES:
        _validate_input(path)
    if suffix == ".docx":
        return _inspect_docx(path, max_items)
    if suffix == ".pptx":
        inspect_pptx = _module("office_pptx").inspect_pptx

        return inspect_pptx(path, max_items)
    if suffix == ".xlsx":
        return _inspect_xlsx(path, max_items)
    if suffix == ".pdf":
        return _inspect_pdf(path, max_items)
    raise OfficeKitError(
        f"inspect 暂不支持 {suffix or '无后缀'}；当前支持 .docx/.pptx/.xlsx/.pdf"
    )


def doctor(args: argparse.Namespace) -> dict[str, Any]:
    current_lo = _lo_path()
    current_lo_version = _lo_version(current_lo)
    font_config, font_signature = _font_environment()
    result: dict[str, Any] = {
        "ok": False,
        "kit_version": KIT_VERSION,
        "python": sys.executable,
        "venv": sys.prefix if sys.prefix != sys.base_prefix else os.environ.get("VIRTUAL_ENV"),
        "packages": _package_versions(),
        "libreoffice": {"path": str(current_lo) if current_lo else None, "version": current_lo_version},
        "fonts": {"config": str(font_config) if font_config is not None else None, "directories": len(_font_directories()), "signature": font_signature},
        "pdfium": {},
        "full_render": None,
    }
    try:
        pdfium = _module("pypdfium2")
        PdfWriter = _module("pypdf").PdfWriter

        artifacts = getattr(args, "artifacts", None)
        if artifacts:
            target = Path(artifacts).expanduser().resolve()
            if target.exists() and any(target.iterdir()):
                raise OfficeKitError("检查样张目录非空；请使用新的目录，不覆盖旧证据")
            target.mkdir(parents=True, exist_ok=True)
            workspace = contextlib.nullcontext(str(target))
            result["artifacts"] = str(target)
        else:
            workspace = tempfile.TemporaryDirectory(prefix="contact-office-doctor-")
        with workspace as temporary:
            pdf_path = Path(temporary) / "blank.pdf"
            writer = PdfWriter()
            writer.add_blank_page(width=72, height=72)
            with pdf_path.open("wb") as stream:
                writer.write(stream)
            document = pdfium.PdfDocument(str(pdf_path))
            try:
                page = document[0]
                bitmap = page.render(scale=1)
                image = bitmap.to_pil()
                result["pdfium"] = {"ok": True, "image": [image.width, image.height]}
            finally:
                document.close()
            if args.full:
                Document = _module("docx").Document
                pptx_module = _module("pptx")
                Presentation = pptx_module.Presentation
                pptx_util = _module("pptx.util")
                Inches, Pt = pptx_util.Inches, pptx_util.Pt

                docx_path = Path(temporary) / "doctor.docx"
                document = Document()
                document.add_heading("合同 Office 渲染健康检查", level=1)
                document.add_paragraph("中文正文：付款比例为 30%，验收期限为 5 个工作日。")
                table = document.add_table(rows=2, cols=2)
                table.cell(0, 0).text = "项目"
                table.cell(0, 1).text = "金额"
                table.cell(1, 0).text = "设备"
                table.cell(1, 1).text = "100,000 元"
                document.save(docx_path)

                pptx_path = Path(temporary) / "doctor.pptx"
                presentation = Presentation()
                presentation.slide_width = Inches(13.333)
                presentation.slide_height = Inches(7.5)
                slide = presentation.slides.add_slide(presentation.slide_layouts[5])
                slide.shapes.title.text = "合同汇报"
                box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(10.5), Inches(1.2))
                paragraph = box.text_frame.paragraphs[0]
                run = paragraph.add_run()
                run.text = "付款比例 30%，验收期限 5 个工作日。"
                run.font.size = Pt(24)
                presentation.save(pptx_path)

                out = Path(temporary) / "render"
                completed = subprocess.run(
                    [sys.executable, str(Path(__file__).resolve()), "render", str(docx_path),
                     str(pptx_path), "--out", str(out), "--dpi", "100"],
                    capture_output=True, text=True, encoding="utf-8", timeout=max(args.timeout, 120), check=False,
                )
                if completed.returncode != 0:
                    detail = ((completed.stderr or "") + "\n" + (completed.stdout or ""))[-2500:].strip()
                    raise OfficeKitError(f"完整渲染失败：{detail}")
                manifest = json.loads(completed.stdout)
                docx_entry = next((entry for entry in manifest["entries"] if entry.get("kind") == "docx"), None)
                if docx_entry is None:
                    raise OfficeKitError("完整渲染未返回 DOCX 结果")
                PdfReader = _module("pypdf").PdfReader
                pdf_text = "\n".join(page.extract_text() or "" for page in PdfReader(docx_entry["pdf"]).pages)
                cjk_chars = [char for char in pdf_text if "\u4e00" <= char <= "\u9fff"]
                cjk_ok = len(cjk_chars) >= 10 and len(set(cjk_chars)) >= 6 and "合同" in pdf_text
                result["full_render"] = {
                    "ok": True,
                    "files": len(manifest["entries"]),
                    "libreoffice_invocations": manifest["libreoffice_invocations"],
                    "pages": sum(len(entry["images"]) for entry in manifest["entries"]),
                    "cjk_pdf_text": cjk_ok,
                    "cjk_unique_glyphs": len(set(cjk_chars)),
                }
                if not cjk_ok:
                    raise OfficeKitError("渲染 PDF 的中文文本层缺失或不可提取；Windows检查系统中文字体，macOS/Linux检查CJK字体及进程fontconfig")
        packages_ok = all(result["packages"].values())
        result["ok"] = bool(packages_ok and current_lo and result["pdfium"].get("ok"))
        if args.full:
            result["ok"] = bool(result["ok"] and result["full_render"]["ok"])
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"
    return result


def pptx_command(args: argparse.Namespace) -> int:
    pptx_module = _module("office_pptx")
    apply_plan = pptx_module.apply_plan
    verify_output = pptx_module.verify_output

    if args.pptx_command == "apply":
        result = apply_plan(
            Path(args.plan), commit=args.commit, overwrite=args.overwrite,
            result_path=Path(args.result) if args.result else None,
            render_dir=Path(args.render_dir) if args.render_dir else None,
        )
        _json_dump(result)
        return 0 if result.get("ok") else 2
    if args.pptx_command == "verify":
        result = verify_output(
            Path(args.plan), Path(args.output),
            result_path=Path(args.result) if args.result else None,
        )
        _json_dump(result)
        return 0 if result.get("ok") else 2
    raise OfficeKitError("未知 pptx 子命令")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="contact 仓库 Office 工具：检查、PPTX 计划式精修、批量无头渲染"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    setup = sub.add_parser("setup", help="创建隔离 Python 环境并安装锁定依赖")
    setup.add_argument("--venv")
    setup.add_argument("--index-url")
    setup.set_defaults(handler=setup_runtime)

    doc = sub.add_parser("doctor", help="检查 Office 运行环境")
    doc.add_argument("--full", action="store_true", help="生成中文 DOCX/PPTX 并完整渲染一次")
    doc.add_argument("--timeout", type=int, default=180)
    doc.add_argument("--artifacts", help="保留虚构检查样张及预览；目录须为空或不存在")
    doc.add_argument("--out", help="另存JSON检查报告")
    def doctor_handler(args):
        result = doctor(args)
        if args.out:
            _json_dump(result, Path(args.out).expanduser().resolve())
        _json_dump(result)
        return 0 if result["ok"] else 1
    doc.set_defaults(handler=doctor_handler)

    render = sub.add_parser("render", help="批量无头转换并生成页面图/联系表")
    render.add_argument("inputs", nargs="+")
    render.add_argument("--out", required=True, help="输出目录（放任务产出目录，不放技能目录）")
    render.add_argument("--dpi", type=int, default=120)
    render.add_argument("--pages", help="如 1,3-5；默认全部页面/幻灯片")
    render.add_argument("--no-cache", action="store_true")
    render.add_argument("--timeout", type=int, default=300)
    def render_handler(args):
        _json_dump(render_documents(args))
        return 0
    render.set_defaults(handler=render_handler)

    inspect = sub.add_parser("inspect", help="输出 DOCX/PPTX/XLSX/PDF 结构清单")
    inspect.add_argument("inputs", nargs="+")
    inspect.add_argument("--max-items", type=int, default=300)
    def inspect_handler(args):
        results = []
        failed = False
        for item in args.inputs:
            path = Path(item).expanduser().resolve()
            try:
                data = inspect_file(path, args.max_items)
            except Exception as exc:
                data = {"input": str(path), "ok": False, "error": f"{type(exc).__name__}: {exc}"}
                failed = True
            data["ok"] = not data.get("error")
            data["input"] = str(path)
            results.append(data)
        _json_dump({"schema": SCHEMA, "results": results})
        return 1 if failed else 0
    inspect.set_defaults(handler=inspect_handler)

    pptx = sub.add_parser("pptx", help="PPTX 计划式精修与验证")
    psub = pptx.add_subparsers(dest="pptx_command", required=True)
    apply_parser = psub.add_parser("apply", help="预演/提交 PPTX 编辑计划")
    apply_parser.add_argument("plan")
    group = apply_parser.add_mutually_exclusive_group()
    group.add_argument("--dry-run", action="store_true", help="默认：只验证和预演")
    group.add_argument("--commit", action="store_true", help="写入新文件；不覆盖源文件")
    apply_parser.add_argument("--overwrite", action="store_true", help="允许替换已存在的目标文件")
    apply_parser.add_argument("--result", help="把执行结果另存 JSON")
    apply_parser.add_argument("--render-dir", help="提交后同时无头渲染到该目录")
    verify_parser = psub.add_parser("verify", help="独立核对计划与输出")
    verify_parser.add_argument("plan")
    verify_parser.add_argument("output")
    verify_parser.add_argument("--result", help="apply 保存的结果 JSON，用于复核后置条件")
    pptx.set_defaults(handler=lambda args: pptx_command(args))

    return parser


def main(argv: list[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    args = parser.parse_args(arguments)
    try:
        ensure_runtime(arguments)
        result = args.handler(args)
        return int(result)
    except OfficeKitError as exc:
        _json_dump({"schema": SCHEMA, "ok": False, "error": str(exc), "code": exc.code},
                   None)
        return exc.code
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        _json_dump({
            "schema": SCHEMA, "ok": False,
            "error": f"内部错误（{type(exc).__name__}）：{exc}", "code": 2,
        })
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
