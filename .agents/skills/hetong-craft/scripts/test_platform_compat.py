#!/usr/bin/env python3
"""Windows/macOS 平台路径与 LibreOffice 发现的离线回归。

这些测试只验证平台选择逻辑，不要求本机安装 Windows 或 LibreOffice。
"""

from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

SCRIPT = Path(__file__).resolve().parent / "office_kit.py"
_spec = importlib.util.spec_from_file_location("office_kit_platform_test", SCRIPT)
if _spec is None or _spec.loader is None:
    raise ImportError("无法加载 office_kit.py")
office_kit = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = office_kit
_spec.loader.exec_module(office_kit)


class PlatformTests(unittest.TestCase):
    def test_venv_python_windows_and_posix(self) -> None:
        with mock.patch.object(office_kit, "_is_windows", return_value=True):
            self.assertTrue(
                str(office_kit._venv_python(Path("C:/work/venv"))).replace("\\", "/").endswith(
                    "Scripts/python.exe"
                )
            )
        with mock.patch.object(office_kit, "_is_windows", return_value=False):
            self.assertTrue(
                str(office_kit._venv_python(Path("/work/venv"))).replace("\\", "/").endswith("bin/python")
            )

    def test_user_roots_use_localappdata_on_windows(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            with mock.patch.dict(os.environ, {"LOCALAPPDATA": temporary}, clear=False), mock.patch.object(office_kit, "_is_windows", return_value=True):
                self.assertEqual(
                    office_kit._user_data_root(),
                    Path(temporary) / "contact-office",
                )
                self.assertEqual(
                    office_kit._user_cache_root(),
                    Path(temporary) / "contact-office",
                )

    def test_windows_libreoffice_path(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            binary = root / "LibreOffice" / "program" / "soffice.exe"
            binary.parent.mkdir(parents=True)
            binary.write_bytes(b"fake")
            with mock.patch.dict(os.environ, {"PROGRAMFILES": str(root)}, clear=False), mock.patch.object(office_kit, "_is_windows", return_value=True), mock.patch.object(office_kit.shutil, "which", return_value=None):
                found = office_kit._lo_path()
            self.assertEqual(found, binary)

    def test_windows_fonts_use_system_fonts_without_fontconfig(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            fonts = Path(temporary) / "Fonts"
            fonts.mkdir()
            (fonts / "simsun.ttc").write_bytes(b"fake-font")
            with mock.patch.object(office_kit, "_is_windows", return_value=True), mock.patch.object(office_kit, "_font_directories", return_value=[fonts]):
                config, signature = office_kit._font_environment()
            self.assertIsNone(config)
            self.assertEqual(len(signature), 24)

    def test_explicit_libreoffice_path_with_spaces_and_unicode(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            binary = Path(temporary) / "办公 软件" / "soffice.exe"
            binary.parent.mkdir()
            binary.write_bytes(b"fixture")
            with mock.patch.dict(os.environ, {"CONTACT_LIBREOFFICE": str(binary)}):
                self.assertEqual(office_kit._lo_path(), binary.resolve())

    def test_invalid_explicit_libreoffice_does_not_fall_back(self) -> None:
        with mock.patch.dict(os.environ, {"CONTACT_LIBREOFFICE": "/missing/soffice.exe"}):
            with self.assertRaises(office_kit.OfficeKitError):
                office_kit._lo_path()

    def test_windows_converter_drops_unix_font_variables(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "中文 文件.docx"
            source.write_bytes(b"fixture")
            output = root / "output"
            def convert(command, **kwargs):
                self.assertNotIn("FONTCONFIG_FILE", kwargs["env"])
                self.assertNotIn("FONTCONFIG_PATH", kwargs["env"])
                self.assertIn(str(source), command)
                (output / "中文 文件.pdf").write_bytes(b"fixture")
                return type("Result", (), {"returncode": 0, "stdout": "", "stderr": ""})()
            with mock.patch.dict(os.environ, {"FONTCONFIG_FILE": "/mac/fonts.conf", "FONTCONFIG_PATH": "/mac/fonts"}), \
                 mock.patch.object(office_kit, "_is_windows", return_value=True), \
                 mock.patch.object(office_kit, "_lo_path", return_value=Path("C:/Office/soffice.exe")), \
                 mock.patch.object(office_kit, "_lo_profile", return_value=root / "profile"), \
                 mock.patch.object(office_kit, "_font_environment", return_value=(None, "fixture")), \
                 mock.patch.object(office_kit, "_flock", return_value=mock.Mock()), \
                 mock.patch.object(office_kit.subprocess, "run", side_effect=convert):
                office_kit._convert_batch([source], output, 10)

    def test_unsupported_platform_does_not_crash_import(self) -> None:
        # 模块导入不应依赖 fcntl/msvcrt；平台相关模块只在函数调用时加载。
        self.assertTrue(callable(office_kit._flock))
        self.assertTrue(callable(office_kit._lo_path))


if __name__ == "__main__":
    unittest.main(verbosity=2)
