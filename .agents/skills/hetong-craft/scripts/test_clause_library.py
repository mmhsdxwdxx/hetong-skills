#!/usr/bin/env python3
"""Markdown source, category selection and portable-library regression tests."""
import json
from pathlib import Path
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parent))
from buyer_clause_select import load_catalog, select_categories, render_modules
from clause_library import index_from_markdown
import hetong_intake


class ClauseLibraryTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory(prefix='hetong-library-')
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        original = Path(__file__).resolve().parent.parent / 'references'
        self.refs = self.root / 'references'
        self.refs.mkdir()
        shutil.copytree(original / 'clauses', self.refs / 'clauses')
        for side in ('buyer', 'seller'):
            shutil.copy2(original / (side + '-clause-catalog.json'), self.refs)
        self.buyer = self.refs / 'buyer-clause-catalog.json'
        self.seller = self.refs / 'seller-clause-catalog.json'

    def test_indexes_never_duplicate_clause_body(self):
        for path in (self.buyer, self.seller):
            data = json.loads(path.read_text())
            self.assertEqual(data['schema_version'], 2)
            for item in data['modules']:
                self.assertNotIn('essential', item)
                self.assertNotIn('example', item)

    def test_seller_user_categories_and_exact_selection(self):
        data, sha = load_catalog(self.seller)
        self.assertEqual(len(data['modules']), 11)
        modules, selection = select_categories(data, ['必要条款'], sha)
        self.assertEqual({x['id'] for x in modules}, {f'B{i:02}' for i in range(1, 12)} - {'B10'})
        self.assertEqual(selection['library_side'], 'seller')
        self.assertNotIn('B10', selection['selected_module_ids'])
        modules, selection = select_categories(data, ['其他条款'], sha)
        self.assertEqual([x['id'] for x in modules], ['B10'])
        self.assertIn('# 乙方', render_modules(modules, selection))

    def test_text_only_update_is_used_without_editing_index(self):
        data, before = load_catalog(self.seller)
        index_before = self.seller.read_bytes()
        p = self.refs / 'clauses/seller/其他条款.md'
        p.write_text(p.read_text() + '\n仅用于隔离测试的正文补充。\n')
        updated, after = load_catalog(self.seller)
        self.assertNotEqual(before, after)
        self.assertEqual(index_before, self.seller.read_bytes())
        selected, selection = select_categories(updated, ['其他条款'], after)
        self.assertIn('仅用于隔离测试的正文补充。', render_modules(selected, selection))

    def test_metadata_update_requires_explicit_reindex(self):
        p = self.refs / 'clauses/seller/其他条款.md'
        p.write_text(p.read_text().replace('## B10 ', '## B10 测试改名：', 1))
        with self.assertRaisesRegex(ValueError, 'reindex'):
            load_catalog(self.seller)
        current = json.loads(self.seller.read_text())
        self.seller.write_text(json.dumps(index_from_markdown(current, self.seller), ensure_ascii=False))
        updated, _ = load_catalog(self.seller)
        self.assertTrue(next(x for x in updated['modules'] if x['id'] == 'B10')['title'].startswith('测试改名：'))

    def test_other_party_text_does_not_change_buyer_fingerprint(self):
        _, before = load_catalog(self.buyer)
        p = self.refs / 'clauses/seller/其他条款.md'
        p.write_text(p.read_text() + '\n测试。\n')
        _, after = load_catalog(self.buyer)
        self.assertEqual(before, after)

    def test_markdown_text_change_invalidates_an_existing_intake_receipt(self):
        source = self.root / 'fictional-contract.txt'
        source.write_text('仅用于受理指纹测试。')
        state = self.root / 'receipt.json'
        with patch.object(hetong_intake, 'DEFAULT_CATALOG', self.buyer):
            hetong_intake.init_state(state, [source], {'perspective': 'buyer', 'stage': 'unsigned', 'action': 'review', 'categories': []}, '虚构测试：不调用预设条款。')
            self.assertEqual(hetong_intake.gate(state)['status'], 'submitted')
            p = self.refs / 'clauses/buyer/一般条款.md'
            p.write_text(p.read_text() + '\n测试正文修改。\n')
            with self.assertRaisesRegex(ValueError, 'catalog changed'):
                hetong_intake.gate(state)

    def test_missing_category_file_is_not_silent_omission(self):
        (self.refs / 'clauses/seller/其他条款.md').unlink()
        with self.assertRaises(ValueError):
            load_catalog(self.seller)

    def test_duplicate_clause_is_rejected(self):
        p = self.refs / 'clauses/seller/其他条款.md'
        body = p.read_text()
        p.write_text(body + '\n' + body[body.index('## B10 '):])
        with self.assertRaisesRegex(ValueError, 'Duplicate clause'):
            load_catalog(self.seller)

    def test_category_mismatch_is_rejected(self):
        p = self.refs / 'clauses/seller/其他条款.md'
        p.write_text(p.read_text().replace('**条款种类**：其他条款', '**条款种类**：必要条款', 1))
        with self.assertRaisesRegex(ValueError, 'category'):
            load_catalog(self.seller)

    def test_path_escape_is_rejected(self):
        index = json.loads(self.seller.read_text())
        index['library_dir'] = '../outside'
        self.seller.write_text(json.dumps(index))
        with self.assertRaisesRegex(ValueError, 'inside'):
            load_catalog(self.seller)

    def test_archive_keeps_references_and_clauses_portable(self):
        expected = {side: load_catalog(self.refs / (side + '-clause-catalog.json'))[1] for side in ('buyer', 'seller')}
        archive = self.root / 'library.zip'
        with ZipFile(archive, 'w') as z:
            for path in self.refs.rglob('*'):
                if path.is_file():
                    z.write(path, path.relative_to(self.root))
        extracted = self.root / 'unpacked'
        with ZipFile(archive) as z:
            self.assertIsNone(z.testzip())
            z.extractall(extracted)
        for side in ('buyer', 'seller'):
            self.assertEqual(load_catalog(extracted / 'references' / (side + '-clause-catalog.json'))[1], expected[side])


if __name__ == '__main__':
    unittest.main(verbosity=2)
