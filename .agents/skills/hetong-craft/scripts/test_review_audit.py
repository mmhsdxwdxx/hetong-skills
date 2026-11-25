"""Behavioral regressions for the read-only auditor using isolated ZIP fixtures."""
import importlib.util
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

_spec = importlib.util.spec_from_file_location(
    "audit_docx_review", Path(__file__).with_name("audit_docx_review.py")
)
if _spec is None or _spec.loader is None:
    raise ImportError("无法加载 audit_docx_review.py")
_module = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = _module
_spec.loader.exec_module(_module)
audit = _module.audit

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'


def run(text):
    return '<w:r><w:t>' + text + '</w:t></w:r>'


def span(text, cid='1'):
    return ('<w:commentRangeStart w:id="' + cid + '"/>' + text +
            '<w:commentRangeEnd w:id="' + cid + '"/>' +
            '<w:r><w:commentReference w:id="' + cid + '"/></w:r>')


def comment(cid='1', text='期限应从实际到货起算。', author='YYX'):
    return f'<w:comment w:id="{cid}" w:author="{author}"><w:p>{run(text)}</w:p></w:comment>'


def rev(text, rid='2', author='YYX', kind='ins'):
    content = run(text) if kind == 'ins' else '<w:r><w:delText>' + text + '</w:delText></w:r>'
    return f'<w:{kind} w:id="{rid}" w:author="{author}">{content}</w:{kind}>'


class ReviewAuditTests(unittest.TestCase):
    def setUp(self):
        self.tmp = TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)

    def package(self, name, body, comments='', rel=True, content_type=True, extras=None):
        parts = {'word/document.xml': f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>',
                 '[Content_Types].xml': f'<Types xmlns="{CT}">' +
                    ('<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' if comments and content_type else '') + '</Types>',
                 'word/_rels/document.xml.rels': f'<Relationships xmlns="{R}">' +
                    ('<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' if comments and rel else '') + '</Relationships>'}
        if comments:
            parts['word/comments.xml'] = f'<w:comments xmlns:w="{W}">{comments}</w:comments>'
        parts.update(extras or {})
        path = self.root / (name + '.zip')
        with ZipFile(path, 'w') as z:
            for part, data in parts.items():
                z.writestr(part, data)
        return path

    def test_real_anchor_across_runs_and_revision(self):
        before = self.package('before', '<w:p>' + run('验收收到通知后开始。') + '</w:p>')
        content = run('验收') + span(rev('收到通知后', '2', kind='del') +
                   rev('实际到货后', '3') + run('开始。'))
        final = self.package('final', '<w:p>' + content + '</w:p>', comment())
        original_bytes = final.read_bytes()
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        anchor = result['comments'][0]['anchors'][0]
        self.assertEqual(anchor['text'], '实际到货后开始。')
        self.assertEqual(anchor['markup_text'], '收到通知后实际到货后开始。')
        self.assertEqual(result['summary']['new_revisions'], 2)
        stats = result['summary']['redline_stats']
        self.assertEqual(stats['status'], 'ok')
        self.assertEqual(stats['baseline_view']['visible_text_chars'], 10)
        self.assertEqual(stats['inserted']['chars'], 5)
        self.assertEqual(stats['deleted']['chars'], 5)
        self.assertEqual(stats['inserted']['ratio'], 0.5)
        self.assertEqual(stats['deleted']['ratio'], 0.5)
        self.assertEqual(final.read_bytes(), original_bytes)

    def test_existing_authors_and_id_history_are_preserved(self):
        body = '<w:p>' + span(run('旧意见'), '0') + rev('旧修订', '7', 'OLD') + '</w:p>'
        before = self.package('before', body, comment('0', '旧批注', 'OLD'))
        final = self.package('final', body + '<w:p>' + span(run('新内容')) + '</w:p>',
                             comment('0', '旧批注', 'OLD') + comment())
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        self.assertEqual(result['summary']['new_comments'], 1)
        self.assertEqual(result['summary']['new_revisions'], 0)
        stats = result['summary']['redline_stats']
        self.assertEqual(stats['status'], 'ok')
        self.assertEqual(stats['inserted']['chars'], 0)
        self.assertEqual(stats['deleted']['chars'], 0)

    def test_same_author_does_not_hide_historical_rewrite(self):
        before = self.package('before', '<w:p>' + rev('原文', '7') + '</w:p>')
        final = self.package('final', '<w:p>' + rev('偷改', '7') + '</w:p>')
        result = audit(final, before, 'YYX')
        self.assertTrue(any('历史修订' in e for e in result['errors']))
        self.assertEqual(result['summary']['redline_stats']['status'], 'unreliable')
        self.assertIsNone(result['summary']['redline_stats']['inserted']['chars'])

    def test_same_author_historical_revision_is_not_counted_as_new(self):
        before = self.package('before', '<w:p>' + rev('历史插入', '7', 'YYX') + '</w:p>')
        final = self.package('final', '<w:p>' + rev('历史插入', '7', 'YYX') +
                             rev('本轮新增', '8', 'YYX') + '</w:p>')
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        self.assertEqual(result['summary']['redline_stats']['inserted']['chars'], 4)

    def test_redline_stats_cover_table_header_footnotes_and_moves(self):
        extras_before = {
            'word/header1.xml': f'<w:hdr xmlns:w="{W}"><w:p>{run("页眉")}</w:p></w:hdr>',
            'word/footer1.xml': f'<w:ftr xmlns:w="{W}"><w:p>{run("页脚")}</w:p></w:ftr>',
            'word/footnotes.xml': f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p>{run("脚注")}</w:p></w:footnote></w:footnotes>',
            'word/endnotes.xml': f'<w:endnotes xmlns:w="{W}"><w:endnote w:id="1"><w:p>{run("尾注")}</w:p></w:endnote></w:endnotes>',
        }
        before_body = ('<w:p>' + run('基线') + rev('历史插入', '20', 'OLD') +
                       rev('历史删除', '21', 'OLD', kind='del') +
                       rev('历史移出', '28', 'OLD', kind='moveFrom') +
                       rev('历史移入', '29', 'OLD', kind='moveTo') + '</w:p>' +
                       '<w:tbl><w:tr><w:tc><w:p>' + run('表格') + '</w:p></w:tc></w:tr></w:tbl>')
        before = self.package('before', before_body, extras=extras_before)
        extras_final = dict(extras_before)
        extras_final['word/header1.xml'] = (
            f'<w:hdr xmlns:w="{W}"><w:p>{run("页眉")}{rev("头", "24")}</w:p></w:hdr>')
        final_body = ('<w:p>' + run('基线') + rev('历史插入', '20', 'OLD') +
                      rev('历史删除', '21', 'OLD', kind='del') +
                      rev('历史移出', '28', 'OLD', kind='moveFrom') +
                      rev('历史移入', '29', 'OLD', kind='moveTo') + rev('新增A1!', '22') +
                      rev('删', '23', kind='del') + rev('移出', '24', kind='moveFrom') +
                      rev('移入', '25', kind='moveTo') +
                      '<w:r><w:rPr><w:rPrChange w:id="27" w:author="YYX"><w:rPr/></w:rPrChange></w:rPr>' +
                      '<w:t>格式</w:t></w:r></w:p>' +
                      '<w:tbl><w:tr><w:tc><w:p>' + run('表格') +
                      rev('表内新增', '26') + '</w:p></w:tc></w:tr></w:tbl>')
        final = self.package('final', final_body, extras=extras_final)
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        stats = result['summary']['redline_stats']
        self.assertEqual(stats['status'], 'ok')
        # 基线：正文2 + 历史插入4 + 历史moveTo4 + 表格2 + 页眉2 + 页脚2 + 脚注2 + 尾注2。
        self.assertEqual(stats['baseline_view']['visible_text_chars'], 20)
        self.assertEqual(stats['inserted']['chars'], 10)
        self.assertEqual(stats['deleted']['chars'], 1)
        self.assertEqual(stats['move_from']['chars'], 2)
        self.assertEqual(stats['move_to']['chars'], 2)
        self.assertEqual(stats['inserted']['ratio'], 10 / 20)
        self.assertEqual(stats['deleted']['ratio'], 1 / 20)

    def test_structural_table_row_and_cell_revisions_are_unreliable(self):
        row = ('<w:tr><w:trPr><w:ins w:id="50" w:author="OLD"/></w:trPr>'
               '<w:tc><w:p>' + run('整行文字') + '</w:p></w:tc></w:tr>')
        cell = ('<w:tr><w:tc><w:tcPr><w:cellDel w:id="51" w:author="OLD"/></w:tcPr>'
                '<w:p>' + run('单元格文字') + '</w:p></w:tc></w:tr>')
        body = '<w:tbl>' + row + cell + '</w:tbl>'
        before = self.package('before', body)
        final = self.package('final', body)
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        stats = result['summary']['redline_stats']
        self.assertEqual(stats['status'], 'unreliable')
        self.assertIsNone(stats['baseline_view']['visible_text_chars'])
        self.assertIsNone(stats['inserted']['chars'])
        self.assertIn('row_insert_delete_structure_not_expandable', stats['reasons'])
        self.assertIn('cell_insert_delete_structure_not_expandable', stats['reasons'])

    def test_redline_stats_exceed_100_percent_without_rejection(self):
        before = self.package('before', '<w:p>' + run('基线') + '</w:p>')
        final = self.package('final', '<w:p>' + run('基线') + rev('新增超过基线', '30') + '</w:p>')
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        stats = result['summary']['redline_stats']
        self.assertEqual(stats['status'], 'ok')
        self.assertGreater(stats['inserted']['ratio'], 1)

    def test_redline_stats_mark_story_structure_change_unreliable(self):
        before = self.package('before', '<w:p>' + run('基线') + '</w:p>')
        final = self.package('final', '<w:p>' + run('基线') + '</w:p>', extras={
            'word/header1.xml': f'<w:hdr xmlns:w="{W}"><w:p>{rev("新增页眉", "33")}</w:p></w:hdr>'})
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        stats = result['summary']['redline_stats']
        self.assertEqual(stats['status'], 'unreliable')
        self.assertIn('story_parts_changed', stats['reasons'])
        self.assertIsNone(stats['inserted']['chars'])

    def test_redline_stats_without_baseline_are_null(self):
        final = self.package('final', '<w:p>' + run('正文') + rev('新增', '31') + '</w:p>')
        stats = audit(final)['summary']['redline_stats']
        self.assertEqual(stats['status'], 'no_baseline')
        self.assertIsNone(stats['baseline_view']['visible_text_chars'])
        self.assertIsNone(stats['inserted']['chars'])
        self.assertIsNone(stats['inserted']['ratio'])
        self.assertIsNone(stats['move_from']['chars'])

    def test_redline_stats_zero_baseline_ratio_is_null(self):
        before = self.package('before', '<w:p>' + run('   \t') + '</w:p>')
        final = self.package('final', '<w:p>' + run('   \t') + rev('新增', '32') + '</w:p>')
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        stats = result['summary']['redline_stats']
        self.assertEqual(stats['status'], 'baseline_empty')
        self.assertEqual(stats['baseline_view']['visible_text_chars'], 0)
        self.assertEqual(stats['inserted']['chars'], 2)
        self.assertIsNone(stats['inserted']['ratio'])

    def test_redline_stats_do_not_double_count_nested_text_revisions(self):
        before = self.package('before', '<w:p>' + run('基线') + '</w:p>')
        nested = ('<w:ins w:id="40" w:author="YYX">' + run('外') +
                  '<w:ins w:id="41" w:author="YYX">' + run('内') +
                  '</w:ins>' + run('层') + '</w:ins>')
        final = self.package('final', '<w:p>' + run('基线') + nested + '</w:p>')
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        self.assertEqual(result['summary']['redline_stats']['inserted']['chars'], 3)

    def test_missing_and_duplicate_anchor_rejected(self):
        for body in [run('无锚点'), span(run('甲')) + span(run('乙'))]:
            with self.subTest(body=body):
                final = self.package('final', '<w:p>' + body + '</w:p>', comment())
                self.assertTrue(audit(final)['errors'])

    def test_relationship_and_content_type_required(self):
        for rel, content_type in [(False, True), (True, False)]:
            with self.subTest(rel=rel):
                final = self.package('final', '<w:p>' + span(run('到货')) + '</w:p>',
                                     comment(), rel=rel, content_type=content_type)
                self.assertTrue(audit(final)['errors'])

    def test_wrong_new_author_rejected(self):
        final = self.package('final', '<w:p>' + span(rev('到货', author='wrong')) + '</w:p>',
                             comment(author='wrong'))
        self.assertEqual(sum('作者不符' in e for e in audit(final, author='YYX')['errors']), 2)

    def test_comments_only_allows_run_split_but_not_text_change(self):
        before = self.package('before', '<w:p>' + run('实际到货后开始。') + '</w:p>')
        final = self.package('final', '<w:p>' + run('实际') + span(run('到货后')) + run('开始。') + '</w:p>', comment())
        self.assertFalse(audit(final, before, 'YYX', True)['errors'])
        changed = self.package('changed', '<w:p>' + span(run('付款后开始。')) + '</w:p>', comment())
        self.assertTrue(audit(changed, before, 'YYX', True)['errors'])

    def test_cross_paragraph_and_image_are_review_warnings_not_failures(self):
        body = ('<w:p><w:commentRangeStart w:id="1"/>' + run('前半句') +
                '</w:p><w:p>' + run('后半句') + '<w:commentRangeEnd w:id="1"/>' +
                '<w:r><w:commentReference w:id="1"/></w:r></w:p>')
        final = self.package('final', body, comment())
        result = audit(final)
        self.assertFalse(result['errors'])
        self.assertTrue(result['warnings'])
        image = self.package('image', '<w:p>' + span('<w:r><w:drawing/></w:r>') + '</w:p>', comment())
        result = audit(image)
        self.assertFalse(result['errors'])
        self.assertTrue(result['comments'][0]['anchors'][0]['has_drawing'])

    def test_duplicate_revision_id_rejected(self):
        final = self.package('final', '<w:p>' + rev('甲') + rev('乙') + '</w:p>')
        self.assertTrue(any('修订 ID 重复' in e for e in audit(final)['errors']))

    def test_historical_comment_removal_rejected(self):
        before = self.package('before', '<w:p>' + span(run('历史')) + '</w:p>', comment())
        final = self.package('final', '<w:p>' + run('历史') + '</w:p>')
        self.assertTrue(any('历史批注' in e for e in audit(final, before)['errors']))

    def test_application_backup_is_opaque(self):
        final = self.package('final', '<w:p>' + run('正文') + '</w:p>', extras={'tbak/modified.xml': 'opaque backup'})
        self.assertFalse(audit(final)['errors'])

    def test_unsafe_zip_path_rejected(self):
        final = self.package('unsafe-path', '<w:p>' + run('正文') + '</w:p>',
                             extras={'../escape.xml': '<x/>'})
        with self.assertRaises(ValueError):
            audit(final)

    def test_dtd_and_entity_rejected_before_parse(self):
        final = self.package('dtd', '<w:p>' + run('正文') + '</w:p>',
                             extras={'word/document.xml': '<!DOCTYPE x [<!ENTITY y "z">]><x/>'})
        with self.assertRaises(ValueError):
            audit(final)

    def test_new_private_referral_split_across_runs_is_blocked(self):
        before = self.package('before', '<w:p>' + run('付款期限。') + '</w:p>')
        body = '<w:p>' + span(run('付款期限。')) + '</w:p>'
        text = '<w:comment w:id="1" w:author="YYX"><w:p>' + run('详见《内部') + run('审核报告.docx》。') + '</w:p></w:comment>'
        final = self.package('final', body, text)
        result = audit(final, before, 'YYX')
        self.assertTrue(result['review_privacy']['new_content_blocked'])
        self.assertTrue(any('隐私铁律' in e for e in result['errors']))

    def test_new_revision_private_referral_is_blocked(self):
        before = self.package('before', '<w:p>' + run('付款期限。') + '</w:p>')
        final = self.package('final', '<w:p>' + run('付款期限。') + rev('详见内部审核报告。') + '</w:p>')
        result = audit(final, before, 'YYX')
        self.assertTrue(result['review_privacy']['new_content_blocked'])

    def test_historical_referral_is_reported_without_rewriting(self):
        body = '<w:p>' + span(run('付款期限。')) + '</w:p>'
        before = self.package('before', body, comment(text='详见《旧审核报告.docx》。'))
        final = self.package('final', body, comment(text='详见《旧审核报告.docx》。'))
        saved = final.read_bytes()
        result = audit(final, before, 'YYX')
        self.assertFalse(result['errors'])
        self.assertFalse(result['review_privacy']['new_content_blocked'])
        self.assertEqual(result['review_privacy']['findings'][0]['scope'], 'historical')
        self.assertTrue(any('历史/无基线' in w for w in result['warnings']))
        self.assertEqual(saved, final.read_bytes())

    def test_no_baseline_does_not_claim_referral_is_new(self):
        body = '<w:p>' + span(run('付款期限。')) + '</w:p>'
        final = self.package('final', body, comment(text='详见《旧审核报告.docx》。'))
        result = audit(final)
        self.assertFalse(result['review_privacy']['has_baseline'])
        self.assertFalse(result['review_privacy']['new_content_blocked'])
        self.assertEqual(result['review_privacy']['findings'][0]['scope'], 'unknown')


if __name__ == '__main__':
    unittest.main()
