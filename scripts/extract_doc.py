#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract plain text from legacy Word .doc (OLE compound) files.

antiword often fails on WPS-exported .doc on Windows. This script scans
WordDocument stream for UTF-16LE Chinese/ASCII runs.

Usage:
    PYTHONIOENCODING=utf-8 python scripts/extract_doc.py "报价.doc"
    PYTHONIOENCODING=utf-8 python scripts/extract_doc.py "报价.doc" -o quote_extract.txt

Requires: pip install olefile
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


def extract_with_antiword(path: Path) -> str | None:
    import shutil
    import subprocess

    antiword = shutil.which('antiword')
    if not antiword:
        return None
    try:
        r = subprocess.run(
            [antiword, '-m', 'UTF-8', str(path)],
            capture_output=True,
            timeout=60,
        )
        text = r.stdout.decode('utf-8', errors='replace').strip()
        if text and len(text) > 50:
            return text
    except Exception:
        pass
    return None


def _is_cjk(ch: str) -> bool:
    o = ord(ch)
    return (
        0x4E00 <= o <= 0x9FFF
        or 0x3400 <= o <= 0x4DBF
        or 0xF900 <= o <= 0xFAFF
        or 0x3000 <= o <= 0x303F
        or 0xFF00 <= o <= 0xFFEF
    )


def extract_ole_utf16(path: Path) -> str:
    try:
        import olefile
    except ImportError as e:
        raise SystemExit(
            'olefile is required: pip install olefile\n' + str(e)
        )

    if not olefile.isOleFile(str(path)):
        raise SystemExit(f'Not an OLE .doc file: {path}')

    ole = olefile.OleFileIO(str(path))
    # Prefer WordDocument; fall back to all streams
    streams = []
    for entry in ole.listdir():
        name = '/'.join(entry)
        if name.upper().endswith('WORDDOCUMENT') or name == 'WordDocument':
            streams.insert(0, entry)
        else:
            streams.append(entry)

    chunks: list[str] = []
    seen = set()

    for entry in streams:
        try:
            data = ole.openstream(entry).read()
        except Exception:
            continue
        # Scan even-aligned UTF-16LE runs
        i = 0
        n = len(data)
        while i + 3 < n:
            # try decode a run starting at i if even
            if i % 2 == 1:
                i += 1
                continue
            chars = []
            j = i
            while j + 1 < n:
                code = data[j] | (data[j + 1] << 8)
                if code == 0:
                    break
                # printable ASCII, common punctuation, CJK
                ch = chr(code)
                if (
                    0x20 <= code <= 0x7E
                    or code in (0x0A, 0x0D, 0x09)
                    or _is_cjk(ch)
                    or 0x2010 <= code <= 0x2027
                    or 0x2030 <= code <= 0x205E
                ):
                    chars.append(ch)
                    j += 2
                else:
                    break
            if len(chars) >= 4:
                s = ''.join(chars)
                # keep runs with Chinese or itinerary keywords
                if any(_is_cjk(c) for c in s) or re.search(
                    r'(Day|RMB|酒店|午餐|晚餐|早餐|行程|报价)', s
                ):
                    if s not in seen:
                        seen.add(s)
                        chunks.append(s)
                i = j
            else:
                i += 2

    ole.close()

    # Join and light cleanup
    text = '\n'.join(chunks)
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    # collapse huge blank runs
    text = re.sub(r'\n{3,}', '\n\n', text)
    # drop pure control junk lines
    lines = []
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        if len(line) == 1 and not _is_cjk(line):
            continue
        lines.append(line)
    return '\n'.join(lines)


def extract(path: Path) -> str:
    t = extract_with_antiword(path)
    if t:
        return t
    return extract_ole_utf16(path)


def main():
    ap = argparse.ArgumentParser(description='Extract text from .doc quote files')
    ap.add_argument('doc', type=Path, help='Path to .doc file')
    ap.add_argument('-o', '--output', type=Path, default=None,
                    help='Write to file (default: stdout)')
    args = ap.parse_args()

    if not args.doc.exists():
        raise SystemExit(f'File not found: {args.doc}')

    text = extract(args.doc)
    if not text.strip():
        raise SystemExit('No text extracted. Try antiword or open in Word and save as .docx/.txt.')

    if args.output:
        args.output.write_text(text, encoding='utf-8')
        print(f'Wrote {args.output} ({len(text)} chars)', file=sys.stderr)
    else:
        # Ensure stdout is utf-8 on Windows
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass
        print(text)


if __name__ == '__main__':
    main()
