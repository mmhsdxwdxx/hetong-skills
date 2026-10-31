import copy
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location('buyer_select', HERE / 'buyer_clause_select.py')
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class BuyerClauseSelectionTests(unittest.TestCase):
    def setUp(self):
        self.catalog, self.digest = MODULE.load_catalog(MODULE.DEFAULT_CATALOG)

    def test_source_coverage_and_same_category_merges(self):
        self.assertEqual(len(self.catalog['sources']), 24)
        self.assertEqual(len(self.catalog['modules']), 20)
        groups = {m['id']: m['source_ids'] for m in self.catalog['modules'] if len(m['source_ids']) > 1}
        self.assertEqual(groups, {'M01': ['A01', 'A07'], 'M03': ['A03', 'A09'], 'M06': ['A06', 'A21', 'A22']})

    def test_each_category_has_expected_modules(self):
        expected = {'保交条款': 3, '必要条款': 11, '一般条款': 4, '其他条款': 1, '共管条款': 1}
        for category, count in expected.items():
            with self.subTest(category=category):
                modules, result = MODULE.select_categories(self.catalog, [category], self.digest)
                self.assertEqual(len(modules), count)
                self.assertTrue(all(m['category'] == category for m in modules))
                self.assertEqual(result['selected_categories'], [category])

    def test_multiple_categories_without_implicit_dependencies(self):
        modules, result = MODULE.select_categories(self.catalog, ['保交条款', '共管条款'])
        self.assertEqual({m['id'] for m in modules}, {'M04', 'M15', 'M16', 'M17'})
        self.assertIn('必要条款', result['unselected_categories'])
        self.assertTrue(result['related_unselected'])
        self.assertNotIn('A01', result['selected_source_ids'])

    def test_duplicate_choices_are_deduplicated(self):
        modules, result = MODULE.select_categories(self.catalog, ['保交条款', '保交条款'])
        self.assertEqual(len(modules), 3)
        self.assertEqual(result['selected_categories'], ['保交条款'])

    def test_explicit_empty_is_empty(self):
        modules, result = MODULE.select_categories(self.catalog, [])
        self.assertEqual(modules, [])
        self.assertEqual(result['selected_source_ids'], [])
        self.assertIn('明确不调用预设条款', MODULE.render_modules(modules, result))

    def test_missing_choice_is_not_all(self):
        proc = subprocess.run([sys.executable, str(HERE / 'buyer_clause_select.py'), 'select'], capture_output=True, text=True)
        self.assertNotEqual(proc.returncode, 0)
        self.assertNotIn('## M', proc.stdout)

    def test_unknown_choice_writes_no_artifact(self):
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / 'selection.md'
            proc = subprocess.run([sys.executable, str(HERE / 'buyer_clause_select.py'), 'select', '--category', '不存在的种类', '--out', str(output)], capture_output=True, text=True)
            self.assertEqual(proc.returncode, 2)
            self.assertFalse(output.exists())

    def test_existing_output_and_secondary_output_are_preserved(self):
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / 'selection.md'
            sidecar = Path(temp) / 'selection.json'
            output.write_text('existing')
            proc = subprocess.run([sys.executable, str(HERE / 'buyer_clause_select.py'), 'select', '--all', '--out', str(output), '--selection-json', str(sidecar)], capture_output=True, text=True)
            self.assertEqual(proc.returncode, 2)
            self.assertEqual(output.read_text(), 'existing')
            self.assertFalse(sidecar.exists())

    def _load_modified(self, change):
        data = copy.deepcopy(self.catalog)
        change(data)
        with tempfile.TemporaryDirectory() as temp:
            file = Path(temp) / 'catalog.json'
            file.write_text(json.dumps(data, ensure_ascii=False))
            return MODULE.load_catalog(file)

    def test_cross_category_merge_rejected(self):
        with self.assertRaisesRegex(ValueError, 'Cross-category'):
            self._load_modified(lambda data: data['modules'][0].update(category='保交条款'))

    def test_missing_source_coverage_rejected(self):
        with self.assertRaisesRegex(ValueError, 'coverage'):
            self._load_modified(lambda data: data['modules'].pop())

    def test_new_category_is_data_driven(self):
        def change(data):
            data['categories'].append('项目自定')
            data['sources'].append({'id': 'A25', 'title': '虚构测试', 'category': '项目自定'})
            data['modules'].append({'id': 'M21', 'source_ids': ['A25'], 'title': '虚构测试', 'category': '项目自定', 'essential': '测试要求', 'example': '测试范例', 'action': 'clause_example'})
        catalog, _ = self._load_modified(change)
        modules, _ = MODULE.select_categories(catalog, ['项目自定'])
        self.assertEqual([m['id'] for m in modules], ['M21'])

    def test_deletion_directives_keep_distinct_categories(self):
        modules, result = MODULE.select_categories(self.catalog, ['其他条款'])
        self.assertEqual([m['source_ids'] for m in modules], [['A02']])
        self.assertEqual(modules[0]['action'], 'delete_instruction')
        self.assertIn('第二部分：删除执行指引', MODULE.render_modules(modules, result))
        self.assertNotIn('A24', result['selected_source_ids'])

    def test_selection_record_is_not_commercial_consent(self):
        _, result = MODULE.select_categories(self.catalog, ['必要条款'], self.digest)
        self.assertTrue(result['user_scope_evidence_required'])
        self.assertEqual(result['selection_status'], 'categories_specified_not_commercial_confirmation')
        self.assertEqual(result['catalog_sha256'], self.digest)

    def test_preserved_timing_and_goods_requirements(self):
        modules = {m['id']: m for m in self.catalog['modules']}
        for mid, terms in {
            'M01': ['【3】个工作日', '向该受益人交付保函', '仅提交开函申请'],
            'M03': ['十四（14）日', '第【30】日', '不得仅因发货或开票自动减额'],
            'M06': ['服务器', '发货前', '优先顺序仅适用于货物信息', 'SN清单本身不替代'],
            'M15': ['第七至第十个工作日', '万分之七', '十（10）个自然日'],
            'M16': ['【2】个工作日', '可用额度不足'],
            'M17': ['两（2）周', '满四（4）周', '满七（7）周'],
            'M18': ['【100000】', '【解除期限起算事件】'],
        }.items():
            for term in terms:
                self.assertIn(term, modules[mid]['example'])

    def test_legacy_preview_is_an_index_not_a_second_copy(self):
        preview = (HERE.parent / 'references/buyer-consolidated-clauses.md').read_text()
        for category in self.catalog['categories']:
            self.assertIn('clauses/buyer/' + category + '.md', preview)
        self.assertNotIn(self.catalog['modules'][0]['example'], preview)


if __name__ == '__main__':
    unittest.main()
