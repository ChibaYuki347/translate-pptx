"""Extract <a:t> text contents from unpacked PPTX slides and notes.

Usage:
    python 2_extract_texts.py [options] <unpacked_dir>

Options:
    --unique
        Emit one line per *distinct* normalized text (sorted, no filename
        prefix). Recommended when building a translation dictionary, because
        it avoids feeding duplicate strings to the translator.

    --output <path>
        Write the result to <path> as UTF-8 (LF newlines) instead of stdout.
        Prefer this on Windows: PowerShell's ``>`` redirection re-encodes the
        child process stdout via ``[Console]::OutputEncoding`` (cp932 on
        Japanese Windows) and writes UTF-16LE+BOM, which corrupts non-ASCII
        punctuation such as EN DASH (U+2013) -> mojibake ``窶?``. Using
        ``--output`` bypasses the shell.

    --keep-as-is PATH
        Path to a "keep-as-is" list file. Texts whose full normalized form
        matches an entry are skipped (not extracted). Default:
        ``.github/agents/data/keep_as_is.txt`` (relative to repo root) if
        present.

    --no-keep-as-is
        Disable the keep-as-is filter (extract everything, including obvious
        product names and URLs).

Default output: one line per occurrence, formatted as
    <filename>: '<text>'
"""
from __future__ import annotations

import argparse
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

# Resolve default keep-as-is file path relative to repo root (two levels up
# from .github/agents/scripts/).
_DEFAULT_KEEP_AS_IS = os.path.normpath(
    os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        '..', 'data', 'keep_as_is.txt',
    )
)


def _load_keep_as_is(path: str) -> tuple[set[str], list[re.Pattern[str]]]:
    """Parse a keep-as-is file into a literal set and a list of regex patterns.

    Returns (literals, regexes). Literal entries are pre-normalized with
    normalize_text so callers can compare against normalized run text directly.
    """
    literals: set[str] = set()
    regexes: list[re.Pattern[str]] = []
    if not path or not os.path.isfile(path):
        return literals, regexes
    with open(path, 'r', encoding='utf-8') as fh:
        for raw in fh:
            line = raw.rstrip('\n').rstrip('\r')
            stripped = line.strip()
            if not stripped or stripped.startswith('#'):
                continue
            if stripped.startswith('re:'):
                pattern = stripped[3:]
                try:
                    regexes.append(re.compile(pattern))
                except re.error as e:
                    print(
                        f"[2_extract_texts] skipping invalid regex {pattern!r}: {e}",
                        file=sys.stderr,
                    )
                continue
            literals.add(normalize_text(stripped))
    return literals, regexes


def _is_keep_as_is(
    text: str,
    literals: set[str],
    regexes: list[re.Pattern[str]],
) -> bool:
    """Return True iff *text* (already normalized) is a full-run keep-as-is match."""
    if text in literals:
        return True
    for pat in regexes:
        if pat.fullmatch(text):
            return True
    return False


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            'Extract <a:t> text contents from unpacked PPTX slides and notes. '
            'Supports an optional keep-as-is filter to skip strings that are '
            'intentionally English (product names, URLs, version strings, etc.) '
            'and should not be sent to the LLM for translation.'
        ),
    )
    parser.add_argument('unpacked_dir', help='Unpacked PPTX directory.')
    parser.add_argument(
        '--unique', action='store_true',
        help='Emit one line per distinct normalized text (sorted).',
    )
    parser.add_argument(
        '--output', dest='output', default=None,
        help=(
            'Write output to PATH as UTF-8 with LF newlines instead of stdout. '
            'Recommended on Windows to avoid PowerShell cp932 re-encoding.'
        ),
    )
    parser.add_argument(
        '--keep-as-is', dest='keep_as_is', default=None,
        help=(
            'Path to a keep-as-is list file. Default: '
            '.github/agents/data/keep_as_is.txt if present.'
        ),
    )
    parser.add_argument(
        '--no-keep-as-is', dest='no_keep_as_is', action='store_true',
        help='Disable the keep-as-is filter (extract everything).',
    )
    args = parser.parse_args()

    # Ensure UTF-8 output regardless of the active code page (e.g. cp932 on Windows).
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

    root = args.unpacked_dir

    # Resolve keep-as-is file path.
    literals: set[str] = set()
    regexes: list[re.Pattern[str]] = []
    keep_path: str | None = None
    if not args.no_keep_as_is:
        if args.keep_as_is:
            keep_path = args.keep_as_is
            if not os.path.isfile(keep_path):
                print(
                    f"[2_extract_texts] keep-as-is file not found: {keep_path}",
                    file=sys.stderr,
                )
                return 2
        elif os.path.isfile(_DEFAULT_KEEP_AS_IS):
            keep_path = _DEFAULT_KEEP_AS_IS
        if keep_path:
            literals, regexes = _load_keep_as_is(keep_path)

    # Open the output sink as UTF-8 with LF newlines so PowerShell's `>`
    # redirection (which re-encodes via cp932 on Japanese Windows) is not in
    # the path. When --output is omitted we still write to stdout for
    # backward compatibility, but callers should prefer --output on Windows.
    if args.output is not None:
        out = open(args.output, 'w', encoding='utf-8', newline='\n')
    else:
        out = sys.stdout

    skipped = 0

    def is_kept(text: str) -> bool:
        return _is_keep_as_is(text, literals, regexes)

    try:
        if args.unique:
            seen: set[str] = set()
            for pattern in PATTERNS:
                for f in sorted(glob.glob(os.path.join(root, pattern))):
                    with open(f, 'r', encoding='utf-8') as fh:
                        content = fh.read()
                    for t in _TEXT_RE.findall(content):
                        t = normalize_text(t)
                        if not t:
                            continue
                        if (literals or regexes) and is_kept(t):
                            skipped += 1
                            continue
                        seen.add(t)
            for t in sorted(seen):
                print(t, file=out)
            print(f"[2_extract_texts] unique texts: {len(seen)}", file=sys.stderr)
            if keep_path:
                print(
                    f"[2_extract_texts] keep-as-is filter skipped {skipped} occurrence(s) (from {keep_path})",
                    file=sys.stderr,
                )
            return 0

        for pattern in PATTERNS:
            for f in sorted(glob.glob(os.path.join(root, pattern))):
                with open(f, 'r', encoding='utf-8') as fh:
                    content = fh.read()
                for t in _TEXT_RE.findall(content):
                    t = normalize_text(t)
                    if not t:
                        continue
                    if (literals or regexes) and is_kept(t):
                        skipped += 1
                        continue
                    print(f"{os.path.basename(f)}: {t!r}", file=out)
        if keep_path:
            print(
                f"[2_extract_texts] keep-as-is filter skipped {skipped} occurrence(s) (from {keep_path})",
                file=sys.stderr,
            )
        return 0
    finally:
        if args.output is not None:
            out.close()


if __name__ == '__main__':
    raise SystemExit(main())
