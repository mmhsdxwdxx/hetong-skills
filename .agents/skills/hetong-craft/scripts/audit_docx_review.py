#!/usr/bin/env python3
"""Read-only DOCX review inventory. No legal, semantic or rendering verdicts."""
import argparse
import hashlib
import json
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, cast
from zipfile import BadZipFile, ZipFile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# Windows 控制台默认编码可能不是 UTF-8；命令行输出统一为 UTF-8。
for _stream in (sys.stdout, sys.stderr):
    try:
        cast(Any, _stream).reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass
Q = lambda name: '{' + W + '}' + name
REV = {'ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange',
       'tblPrChange', 'trPrChange', 'tcPrChange', 'sectPrChange',
       'tblGridChange', 'numberingChange', 'cellIns', 'cellDel', 'cellMerge'}
STORY = re.compile(r'word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$')
MARKERS = {Q(n) for n in ('commentRangeStart', 'commentRangeEnd', 'commentReference')}
MAX_ENTRIES = 50_000
MAX_UNCOMPRESSED = 512 * 1024 * 1024


def plain(node, visible=True):
    if visible and node.tag in {Q('del'), Q('moveFrom')}:
        return ''
    if node.tag in {Q('t'), Q('delText')}:
        return node.text or ''
    if node.tag == Q('tab'):
        return '\t'
    if node.tag in {Q('br'), Q('cr')}:
        return '\n'
    return ''.join(plain(c, visible) for c in node)


def canonical(node):
    """Prefix-insensitive representation, ignoring XML indentation."""
    return (node.tag, sorted(node.attrib.items()),
            node.text if node.text and (node.text.strip() or node.tag in {Q('t'), Q('delText')}) else '',
            [canonical(c) for c in node if c.tag not in MARKERS])


def history_signature(node):
    # Run splitting for a new anchor must not be mistaken for a historic edit.
    if node.tag in {Q('comment'), Q('ins'), Q('del'), Q('moveFrom'), Q('moveTo')}:
        atoms = []
        for run in node.iter(Q('r')):
            props = run.find(Q('rPr'))
            fmt = canonical(props) if props is not None else None
            atoms.extend((char, fmt) for char in plain(run, False))
        value = (node.tag, sorted(node.attrib.items()), atoms)
    else:
        value = canonical(node)
    return hashlib.sha256(json.dumps(value, ensure_ascii=False).encode()).hexdigest()


def read_package(path):
    with ZipFile(path) as z:
        infos = z.infolist()
        names = [info.filename for info in infos]
        if len(names) != len(set(names)):
            raise ValueError('ZIP 中存在重名部件')
        if len(infos) > MAX_ENTRIES:
            raise ValueError('ZIP 条目过多，拒绝解析')
        data = {}
        total = 0
        for info in infos:
            name = info.filename
            if name.startswith('/') or '\\' in name or '..' in Path(name).parts:
                raise ValueError(f'不安全的 ZIP 路径：{name}')
            if info.flag_bits & 0x1:
                raise ValueError('拒绝加密 ZIP 条目')
            total += info.file_size
            if total > MAX_UNCOMPRESSED:
                raise ValueError('ZIP 解压总量超过限制')
            blob = z.read(info)
            if name.endswith(('.xml', '.rels')) and (b'<!DOCTYPE' in blob.upper() or b'<!ENTITY' in blob.upper()):
                raise ValueError(f'{name}: 拒绝 DTD/实体声明')
            data[name] = blob
    roots = {}
    for name, blob in data.items():
        # Preserve opaque application backup/embedded parts; they are not review XML.
        if name == '[Content_Types].xml' or (name.startswith('word/') and
                ((name.count('/') == 1 and name.endswith('.xml')) or
                 (name.startswith('word/_rels/') and name.endswith('.rels')))):
            try:
                roots[name] = ET.fromstring(blob)
            except ET.ParseError as exc:
                raise ValueError(f'{name}: {exc}') from exc
    document = roots.get('word/document.xml')
    if document is None or document.tag != Q('document'):
        raise ValueError('需要常规 WordprocessingML DOCX；未识别 document.xml 命名空间')
    comments_parts = [name for name, root in roots.items() if root.tag == Q('comments')]
    if len(comments_parts) > 1:
        raise ValueError('多个 comments 部件需要专用工具，本检查器不合并其 ID 空间')
    comments = {}
    duplicate_comments = []
    for part in comments_parts:
        for c in roots[part].findall(Q('comment')):
            cid = c.get(Q('id'))
            if cid in comments:
                duplicate_comments.append(cid)
            comments[cid] = {'id': cid, 'part': part, 'author': c.get(Q('author')),
                             'text': '\n'.join(plain(p, False) for p in c.findall(Q('p'))),
                             'signature': history_signature(c)}
    return data, roots, comments, duplicate_comments


def inventory(roots):
    ranges, references, events = defaultdict(list), defaultdict(list), defaultdict(list)
    revisions, paragraphs, errors = [], {}, []
    for part, root in roots.items():
        if not STORY.fullmatch(part):
            continue
        ps = list(root.iter(Q('p')))
        pids = {id(p): i for i, p in enumerate(ps)}
        paragraphs[part] = [plain(p, False) for p in ps]
        parents = {id(child): parent for parent in root.iter() for child in parent}
        active = {}
        event_index = 0

        def walk(node, para=None, hidden=False, _pids=pids, _part=part, _active=active, _parents=parents):
            nonlocal event_index
            event_index += 1
            if node.tag == Q('p'):
                para = _pids[id(node)]
            hidden = hidden or node.tag in {Q('del'), Q('moveFrom')}
            if node.tag in MARKERS:
                cid = node.get(Q('id'))
                name = node.tag.rsplit('}', 1)[-1]
                events[cid].append({'kind': name, 'part': _part, 'order': event_index})
                if name == 'commentRangeStart':
                    if cid in _active:
                        errors.append(f'{_part}: 批注 {cid} 重复开始')
                    _active[cid] = {'part': _part, 'start_order': event_index, 'end_order': None,
                                    'paragraphs': [], 'text': '', 'markup_text': '', 'has_drawing': False}
                elif name == 'commentRangeEnd':
                    span = _active.pop(cid, None)
                    if span is None:
                        errors.append(f'{_part}: 批注 {cid} 结束前未开始')
                    else:
                        span['end_order'] = event_index
                        ranges[cid].append(span)
                else:
                    references[cid].append({'part': _part, 'paragraph': para, 'order': event_index})
            if node.tag.startswith('{' + W + '}') and node.tag.rsplit('}', 1)[-1] in REV:
                parent = _parents.get(id(node))
                is_property = parent is not None and parent.tag in {Q('rPr'), Q('pPr'), Q('trPr'), Q('tcPr')}
                revisions.append({'part': _part, 'id': node.get(Q('id')),
                                  'kind': node.tag.rsplit('}', 1)[-1], 'author': node.get(Q('author')),
                                  'paragraph': para, 'property_or_mark': is_property,
                                  'text': plain(node, False), 'signature': history_signature(node)})
            token = ''
            if node.tag in {Q('t'), Q('delText')}:
                token = node.text or ''
            elif node.tag == Q('tab'):
                token = '\t'
            elif node.tag in {Q('br'), Q('cr')}:
                token = '\n'
            for span in _active.values():
                if para is not None and para not in span['paragraphs']:
                    span['paragraphs'].append(para)
                if node.tag in {Q('drawing'), Q('pict')}:
                    span['has_drawing'] = True
                if token:
                    span['markup_text'] += token
                    if not hidden:
                        span['text'] += token
            for child in node:
                walk(child, para, hidden)

        walk(root)
        for cid in active:
            errors.append(f'{part}: 批注 {cid} 未闭合')
    return ranges, references, events, revisions, paragraphs, errors


def audit(final, before=None, author=None, comments_only=False):
    data, roots, comments, duplicates = read_package(final)
    ranges, refs, events, revisions, paragraphs, errors = inventory(roots)
    warnings = []
    old_comments, old_revisions = {}, []
    package_diff = None
    if before:
        old_data, old_roots, old_comments, _ = read_package(before)
        _, _, _, old_revisions, old_paragraphs, _ = inventory(old_roots)
        package_diff = {
            'changed': sorted(n for n in data.keys() & old_data.keys() if data[n] != old_data[n]),
            'added': sorted(data.keys() - old_data.keys()),
            'removed': sorted(old_data.keys() - data.keys())}
        for cid, old in old_comments.items():
            if cid not in comments or old['signature'] != comments[cid]['signature']:
                errors.append(f'历史批注 {cid} 丢失或内容/署名/文字格式改变，需核实授权')
        current = Counter((r['part'], r['id'], r['kind'], r['signature']) for r in revisions)
        for r in old_revisions:
            key = (r['part'], r['id'], r['kind'], r['signature'])
            if not current[key]:
                errors.append(f"历史修订 {r['part']}:{r['id']} 丢失或改变，需核实授权")
            else:
                current[key] -= 1
        if comments_only and old_paragraphs != paragraphs:
            errors.append('仅批注模式：正文故事部件的段落文字发生变化（包含已有删除文字）')
        for n in package_diff['removed']:
            errors.append(f'原有包部件被移除：{n}')
    elif comments_only:
        errors.append('--comments-only 需要 --before，不能没有基线就证明正文未改')
    if duplicates:
        errors.append(f'重复批注 ID：{duplicates}')
    if comments:
        comment_part = next(c['part'] for c in comments.values())
        ct = roots.get('[Content_Types].xml')
        if ct is None or not any(e.get('PartName') == '/' + comment_part and
                e.get('ContentType') == 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
                for e in ct):
            errors.append('comments Content Type 缺失或错误')
        for part in {event['part'] for values in events.values() for event in values}:
            relname = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
            relroot = roots.get(relname)
            valid = [] if relroot is None else [e for e in relroot
                if e.get('Type', '').endswith('/comments') and e.get('TargetMode') != 'External'
                and posixpath.normpath(posixpath.join(posixpath.dirname(part), e.get('Target', ''))).lstrip('/') == comment_part]
            if len(valid) != 1:
                errors.append(f'{part}: 缺少或重复 comments 关系')
    for cid in events:
        if cid not in comments:
            errors.append(f'批注锚点/引用 {cid} 无对应批注正文')
    new_comments = []
    for cid, comment in comments.items():
        new = cid not in old_comments
        comment['new'] = new if before else None
        comment['anchors'] = ranges[cid]
        comment['references'] = refs[cid]
        comment.pop('signature')
        if new:
            new_comments.append(comment)
            if cid is None or not cid.isdigit():
                errors.append(f'新增批注 ID 非非负整数：{cid}')
        counts = Counter(e['kind'] for e in events[cid])
        if new:
            expected = {'commentRangeStart': 1, 'commentRangeEnd': 1, 'commentReference': 1}
            if dict(counts) != expected:
                errors.append(f'批注 {cid} 需要唯一范围与引用；实得 {dict(counts)}')
            elif ranges[cid] and refs[cid]:
                a, ref = ranges[cid][0], refs[cid][0]
                if a['part'] != ref['part'] or ref['order'] <= a['end_order']:
                    errors.append(f'批注 {cid} 引用未置于同部件范围之后')
            if author and comment['author'] != author:
                errors.append(f'新增批注 {cid} 作者不符：{comment["author"]!r}')
            if not comment['text'].strip():
                errors.append(f'新增批注 {cid} 内容为空')
        elif not events[cid]:
            warnings.append(f'历史批注 {cid} 无直接锚点；可能是回复，需保留线程并用相应工具检查')
        for span in ranges[cid]:
            part_paragraphs = paragraphs[span['part']]
            indices = span['paragraphs']
            span['context'] = ([{'paragraph': i, 'text': part_paragraphs[i]}
                for i in range(max(0, min(indices) - 1), min(len(part_paragraphs), max(indices) + 2))]
                if indices else [])
            if not span['text'].strip():
                warnings.append(f'批注 {cid} 当前可见选区无文字：检查图片、空字段或删除范围')
            elif len(span['text']) <= 3 or len(span['text']) > 350 or len(span['paragraphs']) > 1:
                warnings.append(f'批注 {cid} 选区长度/跨段需语义复核：{len(span["text"])}字、{len(span["paragraphs"])}段')
    old_keys = {(r['part'], r['id']) for r in old_revisions}
    old_counts = Counter((r['part'], r['id']) for r in old_revisions)
    counts = Counter((r['part'], r['id']) for r in revisions)
    for key, count in counts.items():
        if count > max(1, old_counts[key]):
            errors.append(f'修订 ID 重复：{key[0]}:{key[1]}')
    new_revisions = []
    for r in revisions:
        is_new = (r['part'], r['id']) not in old_keys
        r['new'] = is_new if before else None
        r.pop('signature')
        if is_new:
            new_revisions.append(r)
            if r['id'] is None or not r['id'].isdigit():
                errors.append(f'{r["part"]}: 新增修订 ID 非非负整数：{r["id"]}')
            if author and r['author'] != author:
                errors.append(f'新增修订 {r["id"]} 作者不符：{r["author"]!r}')
    if comments_only and new_revisions:
        errors.append('仅批注模式新增了修订')
    if any('commentsExtended' in n or 'commentsIds' in n or 'commentsExtensible' in n for n in data):
        warnings.append('存在扩展/线程式批注；本工具不验证回复、解决状态及扩展 ID 关联')
    text_revisions = [r for r in new_revisions if not r['property_or_mark'] and r['text']]
    result = {
        'file': str(Path(final).resolve()), 'sha256': hashlib.sha256(Path(final).read_bytes()).hexdigest(),
        'before': str(Path(before).resolve()) if before else None,
        'summary': {'comments': len(comments), 'revisions': len(revisions),
                    'new_comments': len(new_comments) if before else None,
                    'new_revisions': len(new_revisions) if before else None,
                    'short_text_revision_candidates': sum(len(r['text']) <= 3 for r in text_revisions),
                    'long_text_revision_candidates': sum(len(r['text']) > 180 for r in text_revisions)},
        'errors': errors, 'warnings': warnings, 'package_diff': package_diff,
        'comments': list(comments.values()), 'revisions': revisions,
        'limits': ['仅基础结构和文本检查；不验证法律结论、语义粒度、Word界面或排版',
                   '未模拟接受/拒绝修订；未证明格式、域、嵌入内容及全部OOXML不变量',
                   '历史同ID改写会报告错误；无before时无法识别本次新增内容',
                   '短/长候选是审阅线索，不是质量评分或自动缺陷判定']}
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx')
    parser.add_argument('--before')
    parser.add_argument('--author', help='Expected author for new review content only')
    parser.add_argument('--comments-only', action='store_true')
    parser.add_argument('--out', type=Path, help='Local JSON audit (contains document text)')
    args = parser.parse_args()
    if args.out and args.out.resolve() in {Path(p).resolve() for p in (args.docx, args.before) if p}:
        parser.error('--out 不能覆盖输入文件')
    try:
        result = audit(args.docx, args.before, args.author, args.comments_only)
    except (OSError, ValueError, ET.ParseError, KeyError, BadZipFile) as exc:
        print(json.dumps({'errors': [str(exc)]}, ensure_ascii=False))
        return 1
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: result[k] for k in ('summary', 'errors', 'warnings')}, ensure_ascii=False, indent=2))
    return 1 if result['errors'] else 0


if __name__ == '__main__':
    sys.exit(main())
