#!/usr/bin/env python3
"""Install the four-skill payload into a fresh, local temporary project."""
from pathlib import Path
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SKILLS = ('hetong-craft', 'hetong-review', 'hetong-fast-xqy-v1', 'hetong-word')


class InstallTests(unittest.TestCase):
    def test_install_keeps_existing_files_and_has_complete_word_runtime(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder)
            existing = target / '.agents/skills/custom/SKILL.md'
            existing.parent.mkdir(parents=True)
            existing.write_text('user-owned fixture')
            subprocess.run(['sh', str(ROOT / 'scripts/install.sh'), str(target)], check=True,
                           capture_output=True, text=True)
            self.assertEqual(existing.read_text(), 'user-owned fixture')
            installed = target / '.agents/skills'
            self.assertEqual((installed / 'VERSION').read_text(), (ROOT / 'VERSION').read_text())
            for name in SKILLS:
                self.assertTrue((installed / name / 'SKILL.md').is_file())
            subprocess.run(['python3', str(installed / 'hetong-word/scripts/init.py'), '--help'],
                           check=True, capture_output=True, text=True)
            subprocess.run(['python3', str(installed / 'hetong-craft/scripts/hetong_redline.py'), '--help'],
                           check=True, capture_output=True, text=True)
            self.assertTrue((installed / 'hetong-word/scripts/office-requirements.lock').is_file())


if __name__ == '__main__':
    unittest.main(verbosity=2)
