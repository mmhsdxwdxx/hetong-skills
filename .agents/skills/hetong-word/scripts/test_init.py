#!/usr/bin/env python3
"""Capability/permission regressions; all documents and environments are synthetic."""
import importlib.util
import contextlib
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
spec = importlib.util.spec_from_file_location("word_init", HERE / "init.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class InitTests(unittest.TestCase):
    def args(self, *args):
        return module.parser().parse_args(args)

    def test_core_works_in_empty_venv_without_lo_and_writes_nothing(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            environment = root / "isolated"
            subprocess.run([sys.executable, "-m", "venv", "--without-pip", str(environment)], check=True)
            before = sorted(str(p.relative_to(root)) for p in root.rglob("*"))
            env = dict(os.environ, CONTACT_LIBREOFFICE="/missing/soffice", PYTHONDONTWRITEBYTECODE="1")
            run = subprocess.run([sys.executable, str(HERE / "init.py"), "--venv", str(environment)],
                                 cwd=root, env=env, capture_output=True, text=True)
            self.assertEqual(run.returncode, 0, run.stderr + run.stdout)
            data = json.loads(run.stdout)
            self.assertEqual(data["required_packages"], [])
            self.assertEqual(data["actions"], [])
            self.assertFalse(data["libreoffice"]["required"])
            self.assertEqual(before, sorted(str(p.relative_to(root)) for p in root.rglob("*")))

    def test_generate_missing_does_not_install(self):
        with tempfile.TemporaryDirectory() as directory:
            environment = Path(directory) / "isolated"
            subprocess.run([sys.executable, "-m", "venv", "--without-pip", str(environment)], check=True)
            data = module.initialize(self.args("--profile", "generate", "--venv", str(environment)))
            self.assertFalse(data["ok"])
            self.assertIn("docx", data["missing"])
            self.assertEqual(data["actions"], [])
            self.assertFalse(data["libreoffice"]["required"])

    def test_check_only_rejects_both_install_flags(self):
        for flag in ("--install-python-deps", "--install-libreoffice"):
            with self.subTest(flag=flag), self.assertRaises(ValueError):
                module.initialize(self.args("--profile", "render", "--check-only", flag))

    def test_invalid_explicit_python_does_not_fallback(self):
        data = module.initialize(self.args("--python", "/missing/python"))
        self.assertFalse(data["ok"])
        self.assertEqual(data["python"], "/missing/python")

    def test_present_but_broken_module_is_not_ready(self):
        with tempfile.TemporaryDirectory() as directory:
            Path(directory, "broken_probe.py").write_text("raise ImportError('missing native dependency')\n")
            with mock.patch.dict(os.environ, {"PYTHONPATH": directory}), \
                 mock.patch.dict(module.MODULES, {"probe-dist": "broken_probe"}):
                data = module.probe(sys.executable, ["probe-dist==0"])
            self.assertEqual(data["missing"], ["broken_probe"])
            self.assertIn("missing native dependency", data["import_errors"]["broken_probe"])

    def test_invalid_explicit_lo_does_not_fallback(self):
        with self.assertRaises(ValueError), mock.patch.object(module.office_kit, "_lo_path") as find:
            module.initialize(self.args("--profile", "render", "--libreoffice", "/missing/soffice"))
        find.assert_not_called()

    def test_invalid_environment_lo_returns_structured_error(self):
        output = io.StringIO()
        with mock.patch.object(sys, "argv", ["init.py", "--profile", "render"]), \
             mock.patch.dict(os.environ, {"CONTACT_LIBREOFFICE": "/missing/soffice"}), \
             mock.patch.object(module, "choose_python", return_value=(sys.executable, {"usable": True, "missing": []})), \
             contextlib.redirect_stdout(output):
            self.assertEqual(module.main(), 2)
        self.assertFalse(json.loads(output.getvalue())["ok"])

    def test_render_requires_lo_without_installing(self):
        with mock.patch.object(module, "choose_python", return_value=(sys.executable, {"usable": True, "missing": []})), \
             mock.patch.object(module.office_kit, "_lo_path", return_value=None), \
             mock.patch.object(module, "install_libreoffice") as install:
            data = module.initialize(self.args("--profile", "render"))
        self.assertFalse(data["ok"])
        self.assertEqual(data["missing"], ["LibreOffice"])
        install.assert_not_called()

    def test_lo_install_is_explicit_and_only_runs_when_missing(self):
        for existing in (None, Path("/existing/soffice")):
            with self.subTest(existing=existing), \
                 mock.patch.object(module, "choose_python", return_value=(sys.executable, {"usable": True, "missing": []})), \
                 mock.patch.object(module.office_kit, "_lo_path", side_effect=[existing, Path("/installed/soffice")]), \
                 mock.patch.object(module.office_kit, "_lo_version", return_value="LibreOffice test"), \
                 mock.patch.object(module, "install_libreoffice") as install:
                data = module.initialize(self.args("--profile", "render", "--install-libreoffice"))
                self.assertTrue(data["ok"])
                self.assertEqual(install.call_count, 0 if existing else 1)
                self.assertEqual(data["actions"], [] if existing else ["installed_libreoffice"])

    def test_generate_excludes_render_and_office_packages(self):
        pins = module.requirements("generate")
        self.assertTrue(any(p.startswith("python-docx==") for p in pins))
        self.assertFalse(any(p.startswith(("pypdf", "pillow", "python-pptx", "openpyxl")) for p in pins))

    def test_install_only_selected_dependencies_into_isolated_venv(self):
        with tempfile.TemporaryDirectory() as directory:
            environment = Path(directory) / "isolated"
            environment.mkdir()
            args = self.args("--profile", "generate", "--venv", str(environment), "--install-python-deps")
            with mock.patch.object(module, "choose_python", return_value=(sys.executable, {"usable": True, "missing": ["docx"]})), \
                 mock.patch.object(module, "probe", return_value={"usable": True, "missing": []}), \
                 mock.patch.object(module.subprocess, "run") as run:
                data = module.initialize(args)
            command = run.call_args.args[0]
            self.assertEqual(command[0], str(module.office_kit._venv_python(environment.resolve())))
            self.assertEqual(command[5:], module.requirements("generate"))
            self.assertNotIn("--upgrade", command)
            self.assertEqual(data["actions"], ["installed_python_dependencies"])

    def test_existing_environment_is_reused_without_install(self):
        with mock.patch.object(module, "choose_python", return_value=(sys.executable, {"usable": True, "missing": []})), \
             mock.patch.object(module.subprocess, "run") as run:
            data = module.initialize(self.args("--profile", "generate", "--install-python-deps"))
        self.assertTrue(data["ok"])
        self.assertEqual(data["actions"], [])
        run.assert_not_called()

    def test_render_has_no_pptx_or_spreadsheet_dependency(self):
        self.assertFalse(any(p.startswith(("python-pptx", "openpyxl")) for p in module.requirements("render")))
        self.assertEqual(module.office_kit._command_modules(["render", "example.docx"]), ("pypdf", "pypdfium2", "PIL"))
        self.assertEqual(module.office_kit._command_modules(["inspect", "example.docx"]), ("docx",))

    def test_core_cannot_request_irrelevant_lo_install(self):
        with self.assertRaises(ValueError):
            module.initialize(self.args("--install-libreoffice"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
