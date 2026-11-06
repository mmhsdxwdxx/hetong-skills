#!/usr/bin/env python3
"""Verify the actual two-page limit against a rendered, fingerprinted report."""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify(source: Path, manifest_path: Path) -> dict:
    from pypdf import PdfReader
    source = source.resolve(strict=True)
    manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
    entries = [item for item in manifest.get('entries', [])
               if Path(item.get('input', '')).resolve() == source]
    if len(entries) != 1:
        raise ValueError('渲染清单须唯一对应本次简要报告')
    entry = entries[0]
    if entry.get('input_sha256') != sha(source):
        raise ValueError('简报已变化，须重新渲染核对页数')
    pdf = Path(entry['pdf']).resolve(strict=True)
    if entry.get('pdf_sha256') != sha(pdf):
        raise ValueError('PDF与渲染清单不一致，须重新验证')
    reader = PdfReader(str(pdf))
    if reader.is_encrypted:
        raise ValueError('无法核对加密PDF的实际页数')
    pages = len(reader.pages)
    if not 1 <= pages <= 2:
        raise ValueError(f'简要报告实际为{pages}页，须压缩内容后重新排版，最多2页')
    return {'ok': True, 'source': str(source), 'source_sha256': sha(source),
            'pdf': str(pdf), 'pdf_sha256': sha(pdf), 'pages': pages, 'max_pages': 2}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--manifest', type=Path, required=True)
    args = parser.parse_args()
    try:
        result = verify(args.source, args.manifest)
    except (ValueError, OSError, ImportError, KeyError) as exc:
        result = {'ok': False, 'error': str(exc)}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result['ok'] else 2


if __name__ == '__main__':
    raise SystemExit(main())
