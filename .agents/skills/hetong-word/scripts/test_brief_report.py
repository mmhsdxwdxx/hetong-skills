#!/usr/bin/env python3
"""Real PDF page counts and stale-render protection, using fictional fixtures."""
import json
from pathlib import Path
import tempfile
import unittest
from pypdf import PdfWriter
import verify_brief_report as brief


class BriefTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / 'fictional.docx'
        self.source.write_bytes(b'fictional report source')
        self.pdf = self.root / 'report.pdf'
        self.manifest = self.root / 'render.json'

    def tearDown(self):
        self.temp.cleanup()

    def render(self, pages):
        writer = PdfWriter()
        for _ in range(pages): writer.add_blank_page(width=595, height=842)
        with self.pdf.open('wb') as stream: writer.write(stream)
        self.manifest.write_text(json.dumps({'entries': [{
            'input': str(self.source), 'input_sha256': brief.sha(self.source),
            'pdf': str(self.pdf), 'pdf_sha256': brief.sha(self.pdf)}]}))

    def test_two_pages_pass(self):
        self.render(2)
        self.assertEqual(brief.verify(self.source, self.manifest)['pages'], 2)

    def test_three_pages_rejected(self):
        self.render(3)
        with self.assertRaisesRegex(ValueError, '最多2页'):
            brief.verify(self.source, self.manifest)

    def test_source_changed_requires_rerender(self):
        self.render(1); self.source.write_bytes(b'changed report')
        with self.assertRaisesRegex(ValueError, '重新渲染'):
            brief.verify(self.source, self.manifest)

    def test_pdf_changed_is_rejected(self):
        self.render(1); self.pdf.write_bytes(self.pdf.read_bytes() + b'changed')
        with self.assertRaisesRegex(ValueError, 'PDF与渲染清单不一致'):
            brief.verify(self.source, self.manifest)


if __name__ == '__main__':
    unittest.main(verbosity=2)
