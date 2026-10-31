#!/usr/bin/env python3
"""Exercise current packages, checksums and shared resources with no host mutation."""
import hashlib
import json
import posixpath
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILLS = {"hetong-craft", "hetong-review", "hetong-fast-xqy-v1"}


class WorkBuddyPackages(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temporary = tempfile.TemporaryDirectory()
        cls.out = Path(cls.temporary.name)
        subprocess.run([sys.executable, str(ROOT / "scripts/build_workbuddy.py"),
                        "--out", str(cls.out)], check=True, capture_output=True)
        cls.manifest = json.loads((cls.out / "manifest.json").read_text())

    @classmethod
    def tearDownClass(cls):
        cls.temporary.cleanup()

    def test_current_architecture(self):
        self.assertEqual({p["name"] for p in self.manifest["packages"]}, SKILLS | {"hetong"})
        self.assertEqual(self.manifest["version"], (ROOT / "VERSION").read_text().strip())
        with zipfile.ZipFile(self.out / "hetong.skill") as archive:
            self.assertIsNone(archive.testzip())
            names = set(archive.namelist())
            for skill in SKILLS:
                self.assertIn(f"hetong/bundled/{skill}/SKILL.md", names)
            self.assertFalse(any("/hetong-server/" in n for n in names))
            self.assertIn("hetong/bundled/hetong-craft/references/contract-timeline.md", names)
            self.assertIn("hetong/bundled/hetong-craft/assets/contract-timeline-template.html", names)

    def test_checksums_and_sizes(self):
        for item in self.manifest["packages"] + [self.manifest["combined"]]:
            data = (self.out / item["path"]).read_bytes()
            self.assertEqual(len(data), item["size"])
            self.assertEqual(hashlib.sha256(data).hexdigest(), item["sha256"])

    def test_combined_relative_links_resolve(self):
        with zipfile.ZipFile(self.out / "hetong-workbuddy.zip") as archive:
            names = set(archive.namelist())
            for name in names:
                self.assertFalse(name.startswith("/") or ".." in name.split("/"))
                self.assertFalse("__pycache__" in name or name.endswith((".pyc", ".DS_Store")))
                if not name.endswith(".md"):
                    continue
                body = re.sub(r"```[\s\S]*?```", "", archive.read(name).decode())
                for link in re.findall(r"\]\(([^)]+)\)", body):
                    link = link.split("#", 1)[0]
                    if not link or "://" in link or link.startswith("/") or "<" in link:
                        continue
                    target = posixpath.normpath(posixpath.join(posixpath.dirname(name), link))
                    self.assertTrue(target in names or any(n.startswith(target.rstrip("/") + "/") for n in names),
                                    (name, link, target))

    def test_source_payload_preserved(self):
        with zipfile.ZipFile(self.out / "hetong-workbuddy.zip") as archive:
            for name in archive.namelist():
                if name.endswith("/SKILL.md") and name.count("/") == 1:
                    continue  # Only entry metadata and the host note are intentionally adapted.
                self.assertEqual(archive.read(name), (ROOT / ".agents/skills" / name).read_bytes(), name)


if __name__ == "__main__":
    unittest.main(verbosity=2)
