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

    def test_same_author_does_not_hide_historical_rewrite(self):
        before = self.package('before', '<w:p>' + rev('原文', '7') + '</w:p>')
        final = self.package('final', '<w:p>' + rev('偷改', '7') + '</w:p>')
        self.assertTrue(any('历史修订' in e for e in audit(final, before, 'YYX')['errors']))

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


if __name__ == '__main__':
    unittest.main()
