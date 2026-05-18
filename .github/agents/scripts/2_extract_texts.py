"""Extract <a:t> text contents from unpacked PPTX slides and notes.

Usage:
    python 2_extract_texts.py <unpacked_dir> [--unique] [--output <path>]

Default output: one line per occurrence, formatted as
    <filename>: '<text>'

With --unique: emit one line per *distinct* trimmed text (sorted, no filename
prefix). This is the recommended mode for building a translation dictionary,
because it avoids feeding duplicate strings to the translator.

With --output <path>: write the result to <path> as UTF-8 (LF newlines) instead
of stdout. Prefer this on Windows: PowerShell's `>` redirection re-encodes the
child process stdout via [Console]::OutputEncoding (cp932 on Japanese Windows)
and writes UTF-16LE+BOM, which corrupts non-ASCII punctuation such as
EN DASH (U+2013) -> mojibake `窶?`. Using --output bypasses the shell.
"""
from __future__ import annotations

import glob
import os
import re
import sys

# Local import (script directory is on sys.path when run as a file).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib._text_normalize import normalize_text  # noqa: E402


PATTERNS = (
    'ppt/slides/slide*.xml',
    'ppt/notesSlides/notesSlide*.xml',
)

_TEXT_RE = re.compile(r'<a:t>(.*?)</a:t>', re.DOTALL)


def main() -> int:
    args = sys.argv[1:]
    unique = False
    if '--unique' in args:
        unique = True
        args.remove('--unique')
    output_path: str | None = None
    if '--output' in args:
        idx = args.index('--output')
        if idx + 1 >= len(args):
            print(__doc__, file=sys.stderr)
            return 2
        output_path = args[idx + 1]
        del args[idx:idx + 2]
    if len(args) != 1:
        print(__doc__, file=sys.stderr)
        return 2
    # Ensure UTF-8 output regardless of the active code page (e.g. cp932 on Windows).
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    root = args[0]

    # Open the output sink as UTF-8 with LF newlines so PowerShell's `>`
    # redirection (which re-encodes via cp932 on Japanese Windows) is not in
    # the path. When --output is omitted we still write to stdout for
    # backward compatibility, but callers should prefer --output on Windows.
    if output_path is not None:
        out = open(output_path, 'w', encoding='utf-8', newline='\n')
    else:
        out = sys.stdout

    try:
        if unique:
            seen: set[str] = set()
            for pattern in PATTERNS:
                for f in sorted(glob.glob(os.path.join(root, pattern))):
                    with open(f, 'r', encoding='utf-8') as fh:
                        content = fh.read()
                    for t in _TEXT_RE.findall(content):
                        t = normalize_text(t)
                        if t:
                            seen.add(t)
            for t in sorted(seen):
                print(t, file=out)
            print(f"[2_extract_texts] unique texts: {len(seen)}", file=sys.stderr)
            return 0

        for pattern in PATTERNS:
            for f in sorted(glob.glob(os.path.join(root, pattern))):
                with open(f, 'r', encoding='utf-8') as fh:
                    content = fh.read()
                for t in _TEXT_RE.findall(content):
                    t = normalize_text(t)
                    if t:
                        print(f"{os.path.basename(f)}: {t!r}", file=out)
        return 0
    finally:
        if output_path is not None:
            out.close()


if __name__ == '__main__':
    raise SystemExit(main())
