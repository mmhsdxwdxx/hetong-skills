#!/usr/bin/env python3
"""Safety and editable Word-format regression tests using fictional input only."""
from copy import deepcopy
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile
import xml.etree.ElementTree as ET

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('xqy_report', HERE / 'build_report.py')
report = importlib.util.module_from_spec(spec)
spec.loader.exec_module(report)
FIXTURE = json.loads((HERE.parent / 'assets/example-report.json').read_text())
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


class ReportTests(unittest.TestCase):
    def test_legacy_input_without_stance_still_builds(self):
        data = deepcopy(FIXTURE); data.pop('stance')
        for item in report.report_items(data): item.pop('fixed_clauses', None)
        self.assertIn('1、基本情况', report.markdown(report.validate(data)))

    def test_buyer_highlight_preserves_analysis_and_live_source_ids(self):
        data = report.validate(deepcopy(FIXTURE))
        text = report.markdown(data)
        self.assertIn('甲方固定条款重点〔M03／A03、A09；缺失〕', text)
        self.assertIn(data['sections'][1]['fixed_clauses'][0]['analysis'], text)
        self.assertLess(text.index('甲方固定条款重点'), text.index('4、产品配套资料'))

    def test_seller_uses_seller_catalog_and_rejects_buyer_code(self):
        data = json.loads((HERE.parent / 'assets/example-report-seller.json').read_text())
        text = report.markdown(report.validate(data))
        self.assertIn('乙方固定条款重点〔B07；冲突〕', text)
        self.assertIn('乙方固定条款重点〔B04；部分符合〕', text)
        data['sections'][1]['fixed_clauses'][0]['id'] = 'M03'
        with self.assertRaises(report.ReportError): report.validate(data)

    def test_fixed_matches_need_explicit_stance_and_known_code(self):
        data = deepcopy(FIXTURE); data.pop('stance')
        with self.assertRaises(report.ReportError): report.validate(data)
        data = deepcopy(FIXTURE); data['sections'][1]['fixed_clauses'][0]['id'] = 'M99'
        with self.assertRaises(report.ReportError): report.validate(data)

    def test_fixed_matches_respect_categories_and_explicit_none(self):
        for categories in ([], ['一般条款']):
            data = deepcopy(FIXTURE); data['stance']['categories'] = categories
            with self.assertRaises(report.ReportError): report.validate(data)
        data = deepcopy(FIXTURE); data['stance']['categories'] = ['必要条款']
        report.validate(data)

    def test_party_name_reversal_uses_actual_trade_role(self):
        data = json.loads((HERE.parent / 'assets/example-report-seller.json').read_text())
        data['stance']['party'] = '甲方'; data['perspective'] = '甲方（实际销售方）'
        text = report.markdown(report.validate(data))
        self.assertIn('审核立场：甲方（销售方）', text)
        self.assertIn('甲方立场·销售方固定条款重点〔B07；冲突〕', text)
        self.assertNotIn('M03', text)

    def test_fixed_match_requires_substantive_text_and_no_duplicate(self):
        data = deepcopy(FIXTURE); data['sections'][1]['fixed_clauses'][0]['analysis'] = ''
        with self.assertRaises(report.ReportError): report.validate(data)
        data = deepcopy(FIXTURE)
        data['sections'][1]['fixed_clauses'] *= 2
        with self.assertRaises(report.ReportError): report.validate(data)

    def test_fixed_match_marked_in_word_with_bold_label_only(self):
        with ZipFile(io.BytesIO(report.word(deepcopy(FIXTURE)))) as z:
            tree = ET.fromstring(z.read('word/document.xml'))
            paragraphs = [p for p in tree.findall('.//w:p', NS) if '甲方固定条款重点' in ''.join(p.itertext())]
            self.assertEqual(len(paragraphs), 1)
            runs = paragraphs[0].findall('w:r', NS)
            self.assertIsNotNone(runs[0].find('w:rPr/w:b', NS))
            self.assertIsNone(runs[1].find('w:rPr/w:b', NS))

    def test_complete_fixture_and_missing_topic(self):
        report.validate(deepcopy(FIXTURE))
        bad = deepcopy(FIXTURE); bad['sections'].pop(3)
        with self.assertRaises(report.ReportError): report.validate(bad)

    def test_duplicate_or_reordered_topics_rejected(self):
        for ids in ([2, 3, 4, 5, 6, 7, 7], [3, 2, 4, 5, 6, 7, 8]):
            bad = deepcopy(FIXTURE)
            for s, n in zip(bad['sections'], ids): s['id'] = n
            with self.assertRaises(report.ReportError): report.validate(bad)

    def test_invalid_status_date_and_injected_url_rejected(self):
        bad = deepcopy(FIXTURE); bad['sections'][0]['status'] = '已自动批准'
        with self.assertRaises(report.ReportError): report.validate(bad)
        bad = deepcopy(FIXTURE); bad['review_date'] = '2000-02-31'
        with self.assertRaises(report.ReportError): report.validate(bad)
        for url in ('javascript:alert(1)', 'https://name:secret@example.invalid'):
            bad = deepcopy(FIXTURE)
            bad['verification']['sources'] = [{'title': '测试', 'url': url, 'checked_on': '2000-01-01'}]
            with self.assertRaises(report.ReportError): report.validate(bad)

    def test_dry_run_creates_nothing(self):
        with tempfile.TemporaryDirectory() as root:
            src = Path(root) / 'source.json'; src.write_text(json.dumps(FIXTURE))
            dest = Path(root) / 'not-created' / 'out.docx'
            result = report.build(src, dest, dry_run=True)
            self.assertTrue(result['dry_run'])
            self.assertFalse(dest.parent.exists())

    def test_build_pair_and_no_overwrite(self):
        with tempfile.TemporaryDirectory() as root:
            src = Path(root) / 'source.json'; src.write_text(json.dumps(FIXTURE))
            dest = Path(root) / 'out.docx'
            original = src.read_bytes()
            result = report.build(src, dest)
            for ext in ('.docx', '.md', '.build.json'):
                self.assertTrue(dest.with_suffix(ext).is_file())
            self.assertEqual(src.read_bytes(), original)
            before = dest.read_bytes()
            with self.assertRaises(report.ReportError): report.build(src, dest)
            self.assertEqual(dest.read_bytes(), before)
            self.assertFalse(result['substantive_review_verified'])
            self.assertFalse(result['visual_review_verified'])
            self.assertEqual(result['fixed_clause_catalog']['role'], 'buyer')
            self.assertEqual(result['fixed_clause_references'][0]['source_ids'], ['A03', 'A09'])

    def test_sidecar_collision_fails_without_partial_output(self):
        with tempfile.TemporaryDirectory() as root:
            src = Path(root) / 'source.json'; src.write_text(json.dumps(FIXTURE))
            dest = Path(root) / 'out.docx'; dest.with_suffix('.md').write_text('用户原稿')
            with self.assertRaises(report.ReportError): report.build(src, dest)
            self.assertFalse(dest.exists())
            self.assertEqual(dest.with_suffix('.md').read_text(), '用户原稿')

    def test_protect_original_directory_and_explicit_input(self):
        with tempfile.TemporaryDirectory() as root:
            src = Path(root) / 'source.json'; src.write_text(json.dumps(FIXTURE))
            with self.assertRaises(report.ReportError): report.build(src, Path(root) / '原件' / 'new.docx')
            dest = Path(root) / 'original.docx'
            data = deepcopy(FIXTURE); data['input_files'] = [str(dest)]
            src.write_text(json.dumps(data))
            with self.assertRaises(report.ReportError): report.build(src, dest)
            data['input_files'] = ['original.docx']
            src.write_text(json.dumps(data))
            with self.assertRaises(report.ReportError): report.build(src, dest)

    def test_word_contains_exact_plain_text_and_only_expected_external_link(self):
        data = deepcopy(FIXTURE)
        text = '测试 <原文> & 金额 0123.45，保留【待确认：比例】\n第二行'
        data['sections'][0]['observation'] = text
        data['verification']['sources'] = [{'title': '演示资料', 'url': 'https://example.invalid/a?x=1&y=2', 'checked_on': '2000-01-01'}]
        blob = report.word(report.validate(data))
        with ZipFile(io.BytesIO(blob)) as z:
            tree = ET.fromstring(z.read('word/document.xml'))
            xml_text = ''.join(tree.itertext())
            self.assertIn(text.replace('\n', ''), xml_text)
            self.assertIn('未确认', xml_text)
            self.assertNotIn('<原文>', z.read('word/document.xml').decode())
            rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
            external = [n.attrib['Target'] for n in rels if n.attrib.get('TargetMode') == 'External']
            self.assertEqual(external, [data['verification']['sources'][0]['url']])
            self.assertEqual(len(tree.findall('.//w:ins', NS)), 0)
            self.assertEqual(len(tree.findall('.//w:del', NS)), 0)
        self.assertIn(text, report.markdown(data))

    def test_word_fonts_headings_pagination_and_footer_are_editable(self):
        with ZipFile(io.BytesIO(report.word(deepcopy(FIXTURE)))) as z:
            styles = ET.fromstring(z.read('word/styles.xml'))
            normal = styles.find("w:style[@w:styleId='Normal']", NS)
            rf = normal.find('w:rPr/w:rFonts', NS)
            for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                self.assertEqual(rf.attrib['{' + NS['w'] + '}' + attr], '宋体')
            self.assertEqual(normal.find('w:rPr/w:sz', NS).attrib['{' + NS['w'] + '}val'], '22')
            tree = ET.fromstring(z.read('word/document.xml'))
            self.assertEqual(len(tree.findall('.//w:br', NS)), 0)
            self.assertEqual(len(tree.findall('.//w:drawing', NS)), 0)
            self.assertEqual(len(styles.findall('.//w:pBdr', NS)), 0)
            fields = ET.fromstring(z.read('word/footer1.xml')).findall('.//w:fldSimple', NS)
            self.assertEqual([f.attrib['{' + NS['w'] + '}instr'] for f in fields], ['PAGE', 'NUMPAGES'])
            headings = []
            for p in tree.findall('.//w:p', NS):
                st = p.find('w:pPr/w:pStyle', NS)
                if st is not None and st.attrib.get('{' + NS['w'] + '}val') == 'Heading1':
                    headings.append(''.join(p.itertext()))
            self.assertEqual([x.split('、')[0] for x in headings[:8]], list('12345678'))

    def test_not_applicable_remains_visible_without_fabricated_sample(self):
        data = deepcopy(FIXTURE)
        data['sections'][2] = {'id': 4, 'status': '不适用', 'location': '服务交付条款', 'observation': '无实物设备，SN与物流资料不适用。'}
        data['extras'] = []
        text = report.markdown(report.validate(data))
        self.assertIn('不适用', text)
        self.assertIn('无实物设备', text)
        self.assertNotIn('框架外明显经济及法律问题', text)


if __name__ == '__main__':
    unittest.main(verbosity=2)
