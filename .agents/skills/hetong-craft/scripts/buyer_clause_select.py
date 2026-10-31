#!/usr/bin/env python3
"""List and select buyer-clause categories without treating selection as deal consent."""
import argparse
from collections import Counter
import hashlib
import json
import os
from pathlib import Path
import sys

SCRIPT_ROOT = Path(__file__).resolve().parent
if str(SCRIPT_ROOT) not in sys.path:
    sys.path.insert(0, str(SCRIPT_ROOT))
from clause_library import load_markdown_catalog, index_from_markdown
DEFAULT_CATALOG = SCRIPT_ROOT.parent / 'references/buyer-clause-catalog.json'
SELLER_CATALOG = SCRIPT_ROOT.parent / 'references/seller-clause-catalog.json'


def load_catalog(path):
    raw = Path(path).read_bytes()
    data = json.loads(raw)
    digest = hashlib.sha256(raw).hexdigest()
    if data.get('schema_version') == 2:
        data, digest = load_markdown_catalog(path)
    if data.get('schema_version') != 1:
        raise ValueError('Unsupported catalog schema')
    categories = data.get('categories', [])
    if not categories or any(not isinstance(x, str) or not x.strip() for x in categories) or len(set(categories)) != len(categories):
        raise ValueError('Invalid or duplicated category names')
    sources = data.get('sources', [])
    source_map = {x['id']: x for x in sources}
    if len(source_map) != len(sources) or not sources:
        raise ValueError('Invalid or duplicated source IDs')
    if any(x['category'] not in categories for x in sources):
        raise ValueError('Source uses unknown category')
    seen = []
    module_ids = []
    for module in data.get('modules', []):
        module_ids.append(module['id'])
        if module['category'] not in categories or not module['source_ids']:
            raise ValueError('Invalid module category or source IDs')
        if any(x not in source_map for x in module['source_ids']):
            raise ValueError('Unknown source ID')
        if any(source_map[x]['category'] != module['category'] for x in module['source_ids']):
            raise ValueError('Cross-category merging is not allowed')
        if any(not isinstance(module.get(k), str) or not module[k].strip() for k in ['id', 'title', 'essential', 'example']):
            raise ValueError('Missing clause content')
        if module.get('action') not in {'clause_example', 'delete_instruction'}:
            raise ValueError('Unknown action kind')
        seen.extend(module['source_ids'])
    if not module_ids or len(set(module_ids)) != len(module_ids):
        raise ValueError('Invalid or duplicated module IDs')
    if Counter(seen) != Counter(source_map.keys()):
        raise ValueError('Source coverage is not one-to-one')
    for relation in data.get('relations', []):
        if any(x not in source_map for x in relation['source_ids']):
            raise ValueError('Unknown related source ID')
    return data, digest


def list_categories(data):
    return [{
        'name': category,
        'module_count': sum(x['category'] == category for x in data['modules']),
        'source_count': sum(x['category'] == category for x in data['sources']),
    } for category in data['categories']]


def select_categories(data, categories, digest=''):
    if not isinstance(categories, list):
        raise ValueError('Categories must be an explicit list')
    unknown = set(categories) - set(data['categories'])
    if unknown:
        raise ValueError('Unknown categories: ' + ', '.join(sorted(unknown)))
    chosen = [x for x in data['categories'] if x in set(categories)]
    modules = [x for x in data['modules'] if x['category'] in chosen]
    selected_sources = {x for module in modules for x in module['source_ids']}
    related = []
    for relation in data.get('relations', []):
        ids = set(relation['source_ids'])
        if ids & selected_sources and ids - selected_sources:
            related.append({**relation, 'unselected_source_ids': [x for x in relation['source_ids'] if x not in selected_sources]})
    selection = {
        'catalog_id': data['catalog_id'], 'catalog_sha256': digest,
        'library_side': data.get('library_side', 'buyer'),
        'selection_status': 'categories_specified_not_commercial_confirmation',
        'user_scope_evidence_required': True,
        'selected_categories': chosen,
        'unselected_categories': [x for x in data['categories'] if x not in chosen],
        'selected_module_ids': [x['id'] for x in modules],
        'selected_source_ids': [x['id'] for x in data['sources'] if x['id'] in selected_sources],
        'related_unselected': related,
    }
    return modules, selection


def render_modules(modules, selection, full=False):
    party = '乙方' if selection.get('library_side') == 'seller' else '甲方'
    lines = [f'# {party}整合条款库' if full else f'# {party}条款分类筛选结果', '',
             '条款范例及删除指引；采用范围和具体条件以本次任务为准。', '',
             '类别选择仅限定预设条款的核对与拟用范围，不确认商业条件，不授权删除原合同中未选类别的条款。', '']
    if full:
        lines += ['完整资料供维护查阅；实际任务应先选类别，再读取筛选结果。', '']
    else:
        lines += ['选中类别：' + ('、'.join(selection['selected_categories']) or '明确不调用预设条款'), '']
    for module in modules:
        lines += [f"## {module['id']} {module['title']}", '',
                  f"**条款种类**：{module['category']}",
                  f"**来源编号**：{'、'.join(module['source_ids'])}", '',
                  '### 第一部分：实质要求', '', module['essential'], '',
                  '### 第二部分：' + ('删除执行指引' if module['action'] == 'delete_instruction' else '范例条款'), '',
                  module['example'], '']
    if not full and selection['related_unselected']:
        lines += ['## 关联提示（不自动增加类别）', '']
        for item in selection['related_unselected']:
            lines.append(f"- {'、'.join(item['source_ids'])}：{item['reason']}")
        lines.append('')
    return '\n'.join(lines)


def write_outputs(outputs):
    paths = [Path(p).resolve() for p, _ in outputs]
    if len(set(paths)) != len(paths):
        raise ValueError('Output paths must be distinct')
    if any(p.exists() for p in paths):
        raise ValueError('Output already exists; choose a new task output path')
    for (path, content), resolved in zip(outputs, paths):
        resolved.parent.mkdir(parents=True, exist_ok=True)
        with resolved.open('x', encoding='utf-8') as stream:
            stream.write(content)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--catalog', type=Path)
    parser.add_argument('--side', choices=('buyer', 'seller'), default='buyer')
    sub = parser.add_subparsers(dest='command', required=True)
    sub.add_parser('list')
    sub.add_parser('validate')
    reindex = sub.add_parser('reindex')
    reindex.add_argument('--write', action='store_true', help='Write the derived index after Markdown metadata changes')
    select = sub.add_parser('select')
    scope = select.add_mutually_exclusive_group(required=True)
    scope.add_argument('--category', action='append')
    scope.add_argument('--all', action='store_true')
    scope.add_argument('--none', action='store_true')
    select.add_argument('--out', type=Path)
    select.add_argument('--selection-json', type=Path)
    render = sub.add_parser('render-all')
    render.add_argument('--out', type=Path)
    args = parser.parse_args(argv)
    try:
        catalog_path = args.catalog or (SELLER_CATALOG if args.side == 'seller' else DEFAULT_CATALOG)
        if args.command == 'reindex':
            current = json.loads(catalog_path.read_text(encoding='utf-8'))
            if current.get('schema_version') != 2:
                raise ValueError('Reindex requires a Markdown-backed library')
            result = index_from_markdown(current, catalog_path)
            text = json.dumps(result, ensure_ascii=False, indent=2) + '\n'
            if args.write:
                temporary = catalog_path.with_name(catalog_path.name + '.tmp')
                temporary.write_text(text, encoding='utf-8')
                os.replace(temporary, catalog_path)
            print(text, end='')
            return 0
        data, digest = load_catalog(catalog_path)
        if args.command == 'validate':
            print(json.dumps({'ok': True, 'library_side': data.get('library_side', args.side), 'catalog_sha256': digest, 'clauses': len(data['modules'])}, ensure_ascii=False))
            return 0
        if args.command == 'list':
            print(json.dumps({'catalog_id': data['catalog_id'], 'catalog_sha256': digest,
        'library_side': data.get('library_side', 'buyer'), 'categories': list_categories(data)}, ensure_ascii=False, indent=2))
            return 0
        requested = data['categories'] if args.command == 'render-all' or args.all else (args.category or [])
        modules, selection = select_categories(data, requested, digest)
        text = render_modules(modules, selection, full=args.command == 'render-all')
        outputs = [(args.out, text)] if args.out else []
        if args.command == 'select' and args.selection_json:
            outputs.append((args.selection_json, json.dumps(selection, ensure_ascii=False, indent=2) + '\n'))
        write_outputs(outputs)
        if not args.out:
            print(text)
        return 0
    except (ValueError, KeyError, TypeError, OSError) as exc:
        print('buyer_clause_select: ' + str(exc), file=sys.stderr)
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
