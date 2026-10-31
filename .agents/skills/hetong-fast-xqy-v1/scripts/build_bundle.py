#!/usr/bin/env python3
"""Publish a linked report and comments-only contract copy in one new directory."""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import shutil
import sys
import tempfile
from zipfile import ZipFile
import xml.etree.ElementTree as ET

HERE = Path(__file__).resolve().parent


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    sys.modules[name] = result
    spec.loader.exec_module(result)
    return result


report = module('xqy_linked_report', HERE / 'build_report.py')
engine = module('xqy_comment_engine', HERE.parents[1] / 'hetong-craft/scripts/hetong_redline.py')
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def sha(blob):
    return hashlib.sha256(blob).hexdigest()


def filename(value, default):
    name = value or default
    if not isinstance(name, str) or any(ord(c) < 32 for c in name) or Path(name).name != name or '/' in name or '\\' in name or not name.endswith('.docx'):
        raise report.ReportError('交付文件名须为无目录的.docx文件名')
    return name


def resolve_source(source_json, data):
    cfg = data.get('bundle')
    if not isinstance(cfg, dict):
        raise report.ReportError('双交付需要bundle源文件及SHA256')
    value = report.required_text(cfg, 'source', 'bundle')
    path = Path(value).expanduser()
    path = (path if path.is_absolute() else source_json.parent / path).resolve()
    expected = report.required_text(cfg, 'source_sha256', 'bundle')
    if not re.fullmatch(r'[0-9a-f]{64}', expected):
        raise report.ReportError('bundle.source_sha256须为实际源文件指纹')
    if path.suffix.lower() != '.docx':
        raise report.ReportError('批注基线须为已核验.docx；旧格式先另存转换并核对')
    if sha(path.read_bytes()) != expected:
        raise report.ReportError('源文件指纹变化，须重新读取并定位，不自动沿用旧锚点')
    return path


def comment_text(issue, report_name):
    section = '框架外问题' if issue['section'] == 'extra' else f"第{issue['section']}项"
    kind = report.ISSUE_TYPES[issue['type']]
    if issue['type'] == 'attention':
        kind += '（正文保留）'
    summary = issue['comment_summary'].rstrip('。；')
    return f"{issue['id']}｜{kind}：{summary}。详见《{report_name}》{section} {issue['id']}。"


def comment_plan(ctx, data, report_name):
    grouped = {}
    for issue in data['issues']:
        if not issue.get('anchors'):
            raise report.ReportError(f"{issue['id']}缺少可核验的批注锚点，不能静默漏批")
        for anchor in issue['anchors']:
            para = ctx.by_id.get(anchor['para'])
            if para is None or not para.editable:
                raise report.ReportError(f"{issue['id']}锚点不存在或含不支持结构：{anchor['para']}")
            text = anchor['text']
            if para.text.count(text) != 1:
                raise report.ReportError(f"{issue['id']}锚文本须在指定段落中唯一命中：{anchor['para']}")
            start = para.text.index(text)
            group = grouped.setdefault(para.pid, {'para': para, 'start': start, 'end': start + len(text), 'issues': {}})
            group['start'] = min(group['start'], start)
            group['end'] = max(group['end'], start + len(text))
            group['issues'][issue['id']] = issue
    edits, mappings = [], []
    for para_id, group in grouped.items():
        ids = list(group['issues'])
        edits.append({'id': '+'.join(ids) + '@' + para_id, 'para': para_id,
                      'old': group['para'].text[group['start']:group['end']],
                      'comment': '\n'.join(comment_text(i, report_name) for i in group['issues'].values())})
        mappings.append({'para': para_id, 'issue_ids': ids, 'anchor_text': edits[-1]['old']})
    plan = engine.build_plan(ctx, edits) if edits else None
    if plan and any(op['kind'] != 'comment' for op in plan['ops']):
        raise report.ReportError('双交付只允许批注，禁止正文替换或插入删除')
    return plan, mappings


def report_text(blob):
    import io
    with ZipFile(io.BytesIO(blob)) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    return ''.join(root.itertext()).replace('\n', '').replace('\r', '')


def check_report_issues(blob, data):
    import io
    text = report_text(blob)
    with ZipFile(io.BytesIO(blob)) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    headings = []
    for paragraph in root.findall('.//w:p', NS):
        style = paragraph.find('w:pPr/w:pStyle', NS)
        if style is not None and style.get('{' + NS['w'] + '}val') == 'Heading2':
            headings.append(''.join(paragraph.itertext()))
    for issue in data['issues']:
        if not any(h.startswith(issue['id'] + '｜') for h in headings):
            raise report.ReportError(f"详细报告缺少{issue['id']}独立事项标题")
        for value in [issue['id'], issue['detail'], issue['suggestion']]:
            if value.replace('\n', '').replace('\r', '') not in text:
                raise report.ReportError(f"详细报告缺少{issue['id']}正文或文字建议")


def prepare(source_json, out_dir):
    data = report.validate(json.loads(source_json.read_text(encoding='utf-8-sig')))
    if 'issues' not in data:
        raise report.ReportError('双交付必须显式列出issues；无事项使用空数组')
    source = resolve_source(source_json, data)
    cfg = data['bundle']
    report_name = filename(cfg.get('report_name'), '合同_v1_XQY审核报告.docx')
    comments_name = filename(cfg.get('comments_name'), '合同_v1_XQY批注稿.docx')
    if report_name == comments_name:
        raise report.ReportError('审核报告与批注稿必须使用不同文件名')
    if out_dir.exists() or out_dir.is_symlink() or '原件' in out_dir.parts or '原件' in out_dir.resolve().parts:
        raise report.ReportError('交付目录必须是新的非原件目录，不覆盖旧交付')
    if source == (out_dir / report_name).resolve() or source == (out_dir / comments_name).resolve():
        raise report.ReportError('交付不得覆盖源文件')
    author = engine.validate_author(cfg.get('author', '合同审阅'))
    ctx = engine.load_context(str(source), for_write=True)
    plan, mappings = comment_plan(ctx, data, report_name)
    return data, source, report_name, comments_name, author, ctx, plan, mappings


def build(source_json, out_dir, dry_run=False):
    source_json, out_dir = source_json.resolve(), out_dir.absolute()
    data, source, rname, cname, author, ctx, plan, mappings = prepare(source_json, out_dir)
    info = {'schema': 'hetong-fast-xqy-v1/bundle-1', 'dry_run': dry_run,
            'report': rname, 'comments': cname, 'issue_ids': [i['id'] for i in data['issues']],
            'comment_count': len(mappings), 'source_sha256': ctx.sha256,
            'body_changes': 0, 'visual_verified': False}
    if dry_run:
        return info
    report_blob = report.word(data)
    check_report_issues(report_blob, data)
    if plan:
        comments_blob, stats = engine.build_result(ctx, plan, author)
        checked = engine.verify_core(ctx.pkg, engine.package_from_bytes(comments_blob, cname), plan)
        if not checked['ok'] or stats['new_revision_ids']:
            raise report.ReportError('批注稿结构验证失败或出现正文修订，未发布')
        for mapping, cid in zip(mappings, stats['new_comment_ids']):
            mapping['word_comment_id'] = cid
    else:
        comments_blob = source.read_bytes()
        checked = {'ok': True, 'notes': ['无新增事项，批注副本与基线字节一致']}
    if sha(source.read_bytes()) != ctx.sha256:
        raise report.ReportError('生成过程中源文件变化，未发布')
    manifest = {**info, 'source': str(source), 'author': author,
                'input_sha256': sha(source_json.read_bytes()), 'mappings': mappings,
                'files': {rname: sha(report_blob), cname: sha(comments_blob)},
                'structural_verified': True, 'engine_verification': checked}
    out_dir.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix='.xqy-pair-', dir=out_dir.parent))
    try:
        (staging / rname).write_bytes(report_blob)
        (staging / cname).write_bytes(comments_blob)
        audit = staging / '核验'; audit.mkdir()
        (audit / 'report.md').write_text(report.markdown(data), encoding='utf-8')
        (audit / 'review-data.json').write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        (audit / 'comment-plan.json').write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding='utf-8')
        (audit / 'pair-manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
        # Only the complete verified directory is exposed as a deliverable.
        if out_dir.exists():
            raise report.ReportError('交付目录已被创建，未覆盖')
        staging.rename(out_dir)
    finally:
        if staging.exists(): shutil.rmtree(staging)
    return {**manifest, 'directory': str(out_dir)}


def verify_bundle(directory):
    audit = directory / '核验'
    manifest = json.loads((audit / 'pair-manifest.json').read_text())
    source = Path(manifest['source'])
    if sha(source.read_bytes()) != manifest['source_sha256']:
        raise report.ReportError('双交付对应的源文件已变化')
    for name, expected in manifest['files'].items():
        if sha((directory / name).read_bytes()) != expected:
            raise report.ReportError(f'成对交付文件已变化：{name}')
    data = json.loads((audit / 'review-data.json').read_text())
    check_report_issues((directory / manifest['report']).read_bytes(), data)
    plan = json.loads((audit / 'comment-plan.json').read_text())
    if plan:
        checked = engine.verify_core(engine.load_package(str(source)), engine.load_package(str(directory / manifest['comments'])), plan)
        if not checked['ok']:
            raise report.ReportError('原生批注或正文保持检查失败')
        with ZipFile(directory / manifest['comments']) as z:
            root = ET.fromstring(z.read('word/comments.xml'))
        comments = {c.get('{' + NS['w'] + '}id'): ''.join(c.itertext()).replace('\n', '') for c in root}
        issues = {issue['id']: issue for issue in data['issues']}
        linked = set()
        for mapping in manifest['mappings']:
            body = comments.get(str(mapping['word_comment_id']), '')
            for code in mapping['issue_ids']:
                if code not in issues or comment_text(issues[code], manifest['report']).replace('\n', '') not in body:
                    raise report.ReportError('报告与原生批注的R编号或说明未对应')
                linked.add(code)
        if linked != set(issues):
            raise report.ReportError('存在未对应批注的报告事项')
    elif (directory / manifest['comments']).read_bytes() != source.read_bytes():
        raise report.ReportError('空事项批注副本与源文件不一致')
    return {'ok': True, 'issue_ids': manifest['issue_ids'], 'comment_count': manifest['comment_count'], 'body_changes': 0}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path)
    parser.add_argument('--out-dir', type=Path)
    parser.add_argument('--dry-run', action='store_true')
    parser.add_argument('--verify', type=Path, help='验证已生成的成对交付目录')
    args = parser.parse_args()
    try:
        if args.verify:
            result = verify_bundle(args.verify)
        elif args.input and args.out_dir:
            result = build(args.input, args.out_dir, args.dry_run)
        else:
            parser.error('提供--input和--out-dir，或--verify')
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (report.ReportError, engine.RedlineError, OSError, ValueError) as exc:
        print(f'错误：{exc}', file=sys.stderr)
        return 2


if __name__ == '__main__':
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8')
        except (AttributeError, OSError):
            pass
    raise SystemExit(main())
