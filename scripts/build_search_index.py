#!/usr/bin/env python3
"""Erzeugt einen lokalen JSON-Suchindex aus HTML- und Markdown-Dateien."""
from __future__ import annotations

import argparse
import json
import re
from datetime import datetime
from html import unescape
from pathlib import Path

# Fix: der Indexer hat bisher ALLE .html/.md-Dateien im Repo erfasst und ihren
# vollen Text in die öffentliche search-index.json geschrieben - auch Entwürfe
# oder interne Notizen, sofern sie irgendwo im Repo lagen. Diese Ordnernamen
# werden jetzt standardmäßig übersprungen (bei Bedarf ergänzen/anpassen):
SKIP_DIRS = {
    '.git', 'node_modules', 'vendor', 'assets', '.github',
    'drafts', '_drafts', 'draft', 'entwuerfe', 'entwürfe',
    'private', 'internal', 'intern', 'notizen', 'notes',
}

NOINDEX_PATTERN = re.compile(
    r'<meta[^>]+name=["\']robots["\'][^>]+content=["\'][^"\']*noindex',
    re.I,
)


def strip_html(text: str) -> str:
    text = re.sub(r'<script\b[^>]*>.*?</script>', ' ', text, flags=re.I | re.S)
    text = re.sub(r'<style\b[^>]*>.*?</style>', ' ', text, flags=re.I | re.S)
    text = re.sub(r'<[^>]+>', ' ', text)
    return re.sub(r'\s+', ' ', unescape(text)).strip()


def frontmatter(raw: str) -> tuple[dict[str, str], str]:
    if not raw.startswith('---'):
        return {}, raw
    parts = raw.split('---', 2)
    if len(parts) < 3:
        return {}, raw
    meta: dict[str, str] = {}
    for line in parts[1].splitlines():
        if ':' in line:
            key, value = line.split(':', 1)
            meta[key.strip()] = value.strip().strip('"\'')
    return meta, parts[2]


def title_from_html(raw: str, fallback: str) -> str:
    for pattern in (r'<h1[^>]*>(.*?)</h1>', r'<title[^>]*>(.*?)</title>'):
        match = re.search(pattern, raw, flags=re.I | re.S)
        if match:
            return strip_html(match.group(1))
    return fallback


def excerpt(text: str, length: int = 220) -> str:
    return text if len(text) <= length else text[:length].rsplit(' ', 1)[0] + ' …'


def url_for(path: Path, root: Path) -> str:
    rel = path.relative_to(root).as_posix()
    if rel.endswith('index.html'):
        rel = rel[:-10]
    elif rel.endswith('.html'):
        rel = rel[:-5]
    elif rel.endswith('.md'):
        rel = rel[:-3]
    return '/' + rel.lstrip('/')


def should_skip(path: Path, output: Path) -> bool:
    if path == output:
        return True
    return any(part in SKIP_DIRS for part in path.parts)


def is_excluded_from_search(raw: str, meta: dict[str, str]) -> bool:
    """Seiten, die explizit als nicht-öffentlich/nicht-suchbar markiert sind, überspringen.

    Wird respektiert:
    - HTML: <meta name="robots" content="noindex ..."> (üblich für nicht verlinkte Seiten)
    - Markdown-Frontmatter: searchable: false / noindex: true / draft: true / published: false
    """
    if NOINDEX_PATTERN.search(raw):
        return True
    if meta.get('searchable', '').strip().lower() in {'false', 'no', '0'}:
        return True
    if meta.get('noindex', '').strip().lower() in {'true', 'yes', '1'}:
        return True
    if meta.get('draft', '').strip().lower() in {'true', 'yes', '1'}:
        return True
    if meta.get('published', '').strip().lower() in {'false', 'no', '0'}:
        return True
    return False


def build(root: Path, output: Path) -> None:
    entries = []
    skipped_private = 0
    for path in sorted(root.rglob('*')):
        if not path.is_file() or path.suffix.lower() not in {'.html', '.md'} or should_skip(path, output):
            continue
        if path.name.lower() == 'search.html':
            continue

        raw = path.read_text(encoding='utf-8', errors='ignore')
        meta, body = frontmatter(raw) if path.suffix.lower() == '.md' else ({}, raw)

        if is_excluded_from_search(raw, meta):
            skipped_private += 1
            continue

        text = strip_html(body)
        if len(text) < 40:
            continue

        fallback = path.stem.replace('-', ' ').replace('_', ' ').title()
        title = meta.get('title') or (title_from_html(raw, fallback) if path.suffix.lower() == '.html' else fallback)
        date = meta.get('date', '')
        tags_raw = meta.get('tags', '')
        tags = [tag.strip() for tag in re.split(r'[,;]', tags_raw.strip('[]')) if tag.strip()]

        entries.append({
            'title': title,
            'url': meta.get('permalink') or url_for(path, root),
            'excerpt': meta.get('description') or excerpt(text),
            'content': text,
            'date': date,
            'tags': tags,
        })

    output.write_text(json.dumps(entries, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print(f'{len(entries)} Seiten indexiert → {output}')
    if skipped_private:
        print(f'{skipped_private} Seite(n) wegen noindex/draft/searchable:false übersprungen '
              f'(nicht im öffentlichen Suchindex enthalten).')


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', default='.', help='Wurzelordner der Website')
    parser.add_argument('--output', default='search-index.json', help='Zieldatei')
    args = parser.parse_args()
    root = Path(args.root).resolve()
    output = Path(args.output)
    if not output.is_absolute():
        output = root / output
    build(root, output)


if __name__ == '__main__':
    main()
