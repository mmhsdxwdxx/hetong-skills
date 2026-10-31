#!/usr/bin/env python3
"""Read category Markdown as clause text and derive the metadata-only index."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re


def _inside(base, relative):
    value = Path(relative)
    if value.is_absolute() or '..' in value.parts:
        raise ValueError('Library paths must remain inside the skill reference directory')
    resolved = (base / value).resolve()
    try:
        resolved.relative_to(base.resolve())
    except ValueError:
        raise ValueError('Library path escapes the skill reference directory') from None
    return resolved


def read_category_file(path):
    text = Path(path).read_text(encoding='utf-8')
    header = text.split('\n## ', 1)[0]
    category = re.search(r'^\*\*条款种类\*\*：([^\n]+)$', header, re.M)
    if not category:
        raise ValueError('Category Markdown is missing its category header: ' + str(path))
    category = category[1].strip()
    items = []
    for block in re.finditer(r'^## ([A-Z]\d+) ([^\n]+)\n(.*?)(?=^## |\Z)', text, re.M | re.S):
        ident, title, body = block.groups()
        item_category = re.search(r'^\*\*条款种类\*\*：([^\n]+)$', body, re.M)
        sources = re.search(r'^\*\*来源编号\*\*：([^\n]+)$', body, re.M)
        parts = re.search(r'^### 第一部分：实质要求\n\n(.*?)\n\n### 第二部分：(范例条款|删除执行指引)\n\n(.*)', body, re.M | re.S)
        if not item_category or item_category[1].strip() != category or not sources or not parts:
            raise ValueError('Invalid two-part clause or category: ' + ident)
        source_ids = [s for s in re.split(r'[、,\s]+', sources[1].strip()) if s]
        if not source_ids or len(source_ids) != len(set(source_ids)):
            raise ValueError('Invalid source IDs: ' + ident)
        essential, kind, example = parts.groups()
        if not essential.strip() or not example.strip():
            raise ValueError('Empty clause text: ' + ident)
        items.append({'id': ident, 'title': title.strip(), 'category': category,
                      'source_ids': source_ids, 'essential': essential.strip(),
                      'example': example.strip(),
                      'action': 'delete_instruction' if kind == '删除执行指引' else 'clause_example'})
    if not items:
        raise ValueError('Empty category files are not needed: ' + str(path))
    if len(re.findall(r'^## ', text, re.M)) != len(items):
        raise ValueError('Unrecognized clause heading: ' + str(path))
    return category, items


def read_library(index, index_path):
    base = Path(index_path).resolve().parent
    directory = _inside(base, index['library_dir'])
    if not directory.is_dir():
        raise ValueError('Missing Markdown clause directory: ' + str(directory))
    files, modules, categories, ids = [], [], set(), set()
    for path in sorted(directory.glob('*.md')):
        _inside(directory, path.name)
        category, items = read_category_file(path)
        if category in categories:
            raise ValueError('More than one file for category: ' + category)
        categories.add(category)
        relative = path.relative_to(base).as_posix()
        files.append({'path': relative, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()})
        for item in items:
            if item['id'] in ids:
                raise ValueError('Duplicate clause ID: ' + item['id'])
            ids.add(item['id'])
            modules.append({**item, 'file': relative})
    if not modules:
        raise ValueError('Clause library has no category Markdown files')
    modules.sort(key=lambda x: (x['id'][0], int(x['id'][1:])))
    return modules, files


def index_from_markdown(index, index_path, modules=None):
    if modules is None:
        modules, _ = read_library(index, index_path)
    old_sources = {x['id']: x for x in index.get('sources', [])}
    categories = list(index.get('categories', []))
    sources = []
    seen = set()
    for item in modules:
        if item['category'] not in categories:
            categories.append(item['category'])
        for ident in item['source_ids']:
            if ident in seen:
                raise ValueError('Source ID belongs to multiple clauses: ' + ident)
            seen.add(ident)
            sources.append({'id': ident, 'title': old_sources.get(ident, {}).get('title', item['title']),
                            'category': item['category']})
    sources.sort(key=lambda x: (x['id'][0], int(x['id'][1:])))
    if any(ident not in seen for relation in index.get('relations', []) for ident in relation['source_ids']):
        raise ValueError('Relation refers to a missing source ID')
    result = {k: v for k, v in index.items() if k not in ('modules', 'sources', 'categories')}
    result.update(schema_version=2, categories=categories, sources=sources,
                  modules=[{k: v for k, v in item.items() if k not in ('essential', 'example')} for item in modules])
    return result


def load_markdown_catalog(path):
    path = Path(path)
    raw = path.read_bytes()
    index = json.loads(raw)
    if index.get('schema_version') != 2 or index.get('library_side') not in ('buyer', 'seller'):
        raise ValueError('Unsupported Markdown clause index')
    if any('essential' in item or 'example' in item for item in index.get('modules', [])):
        raise ValueError('Index must not duplicate Markdown clause text')
    modules, files = read_library(index, path)
    regenerated = index_from_markdown(index, path, modules)
    if regenerated != index:
        raise ValueError('Clause metadata changed; run reindex before using this library')
    # Include text bytes, not just index metadata, in every intake/library receipt.
    identity = {'index_sha256': hashlib.sha256(raw).hexdigest(), 'files': files}
    digest = hashlib.sha256(json.dumps(identity, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
    return {**index, 'schema_version': 1, 'storage_schema_version': 2,
            'modules': modules, 'source_files': files}, digest
