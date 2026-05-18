"""Apply Japanese translations, lang attribute updates, and Yu Gothic UI font
replacement to an unpacked PPTX directory.

Fixed (pre-prepared) logic:
  - <a:t> text replacement using a caller-supplied JSON translation dictionary
    (applied to ppt/slides/*.xml and ppt/notesSlides/*.xml).
  - lang="en-*" -> lang="ja-JP" rewrite (slides + notesSlides).
  - East-Asian font (<a:ea>) rewrite to Yu Gothic UI across slides, notesSlides,
    slideLayouts, slideMasters, notesMasters, and theme files.
  - Insertion of <a:ea> into <a:rPr> / <a:endParaRPr> / <a:defRPr> elements
    that lack it, in OOXML-schema-correct position (after <a:latin>, before
    <a:hlinkClick>/<a:hlinkMouseOver>/<a:rtl>/<a:extLst> or closing tag).

The translation dictionary is the only per-job input and is supplied as a JSON
file mapping exact English source text to its Japanese translation.

Usage:
    python 3_apply_translations.py <unpacked_dir> <translations.json | translations_dir>

The translations argument may be either:
  * a single JSON file mapping English source -> Japanese translation, OR
  * a directory containing one or more ``*.json`` shards with the same shape.
    All shards are merged (later shards override earlier ones on key collision,
    in lexicographic filename order).
"""
from __future__ import annotations

import argparse
import glob
import html
import json
import os
import re
import sys
import time
from xml.sax.saxutils import escape as _xml_escape

# Local import (script directory is on sys.path when run as a file).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib._text_normalize import normalize_text  # noqa: E402


DEFAULT_EA_FONT = {
    "typeface": "Yu Gothic UI",
    "panose": "020B0400000000000000",
    "pitchFamily": "50",
    "charset": "-128",
}

FONT_PRESETS: dict[str, dict[str, str]] = {
    "yu-gothic-ui": DEFAULT_EA_FONT,
    "yu-gothic-ui-semibold": {
        "typeface": "Yu Gothic UI Semibold",
        "panose": "020B0500000000000000",
        "pitchFamily": "50",
        "charset": "-128",
    },
    "meiryo": {
        "typeface": "Meiryo",
        "panose": "020B0604030504040204",
        "pitchFamily": "50",
        "charset": "-128",
    },
    "biz-udpgothic": {
        "typeface": "BIZ UDPGothic",
        "panose": "020B0400000000000000",
        "pitchFamily": "50",
        "charset": "-128",
    },
    "noto-sans-jp": {
        "typeface": "Noto Sans JP",
        "panose": "020B0500000000000000",
        "pitchFamily": "50",
        "charset": "-128",
    },
}


def build_ea_tag(font: dict[str, str]) -> str:
    return (
        f'<a:ea typeface="{font["typeface"]}" panose="{font["panose"]}" '
        f'pitchFamily="{font["pitchFamily"]}" charset="{font["charset"]}"/>'
    )


# Kept as a module-level default so legacy callers (importing EA_TAG directly)
# still get a sensible value. Runtime overrides flow through the parameterized
# `_apply_fonts(content, ea_tag=...)` path.
EA_TAG = build_ea_tag(DEFAULT_EA_FONT)

TEXT_TARGETS = (
    'ppt/slides/slide*.xml',
    'ppt/notesSlides/notesSlide*.xml',
)

FONT_TARGETS = (
    'ppt/slides/slide*.xml',
    'ppt/notesSlides/notesSlide*.xml',
    'ppt/slideLayouts/slideLayout*.xml',
    'ppt/slideMasters/slideMaster*.xml',
    'ppt/notesMasters/notesMaster*.xml',
    'ppt/theme/theme*.xml',
    # Containers that also carry <a:rPr><a:ea/></a:rPr> styling fragments and
    # were previously missed by the FONT_TARGETS scan. Without these globs,
    # SmartArt nodes (diagrams/*.xml) and presentation-level defaults
    # (presentation.xml) keep their original East-Asian fonts, leaving stray
    # MS Mincho / Yu Mincho / Latin substitutes in the rendered slides.
    'ppt/presentation.xml',
    'ppt/diagrams/data*.xml',
    'ppt/diagrams/drawing*.xml',
    'ppt/diagrams/layout*.xml',
    'ppt/diagrams/colors*.xml',
    'ppt/diagrams/quickStyle*.xml',
    'ppt/charts/chart*.xml',
)

_LANG_RE = re.compile(r'lang="en-[A-Za-z]+"')
_TEXT_RE = re.compile(r'<a:t>(.*?)</a:t>', re.DOTALL)
_EA_RE = re.compile(r'<a:ea\b[^/>]*/>')
_BLOCK_RE = re.compile(
    r'<a:(rPr|endParaRPr|defRPr)\b[^>]*?/>'
    r'|<a:(rPr|endParaRPr|defRPr)\b[^>]*?>.*?</a:\2>',
    re.DOTALL,
)
_LATIN_RE = re.compile(r'(<a:latin\b[^/>]*/>)')
_SELFCLOSE_RE = re.compile(r'(<a:(?:rPr|endParaRPr|defRPr)\b[^>]*?)/>$')
_ANCHOR_RE = re.compile(r'(<a:(?:hlinkClick|hlinkMouseOver|rtl|extLst)\b)')


def _normalize_key(s: str) -> str:
    # Delegate to the shared normalizer so extract_texts.py and this script
    # produce identical dictionary keys (handles U+2011 non-breaking hyphen,
    # NBSP, HTML entities, internal whitespace runs, etc.).
    return normalize_text(s)


def _escape_value(s: str) -> str:
    # Accept translator values that may contain raw '&', '<', '>' or already-
    # escaped entities. Unescape first to a logical string, then re-escape so
    # the result is always well-formed XML character data.
    return _xml_escape(html.unescape(s))


def _replace_text(content: str, translations: dict[str, str]) -> str:
    # Pre-normalize dictionary keys once.
    norm_map = {_normalize_key(k): v for k, v in translations.items()}

    def repl(m: re.Match) -> str:
        original = m.group(1)
        key = _normalize_key(original)
        if key in norm_map:
            return f'<a:t>{_escape_value(norm_map[key])}</a:t>'
        return m.group(0)

    return _TEXT_RE.sub(repl, content)


def _fix_block(m: re.Match, ea_tag: str = EA_TAG) -> str:
    block = m.group(0)
    if '<a:ea' in block:
        return block
    tag_match = re.match(r'<a:(\w+)', block)
    if not tag_match:
        return block
    tag = tag_match.group(1)
    if _LATIN_RE.search(block):
        return _LATIN_RE.sub(r'\1' + ea_tag, block, count=1)
    self_close = _SELFCLOSE_RE.match(block)
    if self_close:
        return f'{self_close.group(1)}>{ea_tag}</a:{tag}>'
    if _ANCHOR_RE.search(block):
        return _ANCHOR_RE.sub(ea_tag + r'\1', block, count=1)
    return re.sub(r'(</a:' + tag + r'>)', ea_tag + r'\1', block, count=1)


def _apply_fonts(content: str, ea_tag: str = EA_TAG) -> str:
    content = _EA_RE.sub(ea_tag, content)
    content = _BLOCK_RE.sub(lambda m: _fix_block(m, ea_tag), content)
    return content


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Apply Japanese translations, lang attribute updates, and "
            "East-Asian font replacement to an unpacked PPTX directory."
        )
    )
    parser.add_argument("unpacked_dir", help="Unpacked PPTX directory")
    parser.add_argument(
        "translations",
        help="Translation JSON file or directory of *.json shards",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress per-shard/per-batch progress lines (final summary still emitted).",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Emit one log line per file processed (debug; very chatty).",
    )
    parser.add_argument(
        "--ea-preset",
        default="yu-gothic-ui",
        choices=sorted(FONT_PRESETS.keys()),
        help="East-Asian font preset (default: yu-gothic-ui)",
    )
    parser.add_argument(
        "--ea-font",
        default=None,
        help="Override East-Asian typeface name (e.g. 'Meiryo'). Takes precedence over --ea-preset.",
    )
    parser.add_argument("--ea-panose", default=None, help="Override panose attribute.")
    parser.add_argument(
        "--ea-pitch-family",
        default=None,
        help="Override pitchFamily attribute.",
    )
    parser.add_argument("--ea-charset", default=None, help="Override charset attribute.")
    args = parser.parse_args()

    root, translations_path = args.unpacked_dir, args.translations

    def log(msg: str) -> None:
        if not args.quiet:
            print(msg, file=sys.stderr, flush=True)

    def vlog(msg: str) -> None:
        if args.verbose:
            print(msg, file=sys.stderr, flush=True)

    t0 = time.monotonic()

    font = dict(FONT_PRESETS[args.ea_preset])
    if args.ea_font:
        font["typeface"] = args.ea_font
    if args.ea_panose:
        font["panose"] = args.ea_panose
    if args.ea_pitch_family:
        font["pitchFamily"] = args.ea_pitch_family
    if args.ea_charset:
        font["charset"] = args.ea_charset
    ea_tag = build_ea_tag(font)

    translations: dict[str, str] = {}
    if os.path.isdir(translations_path):
        shards = sorted(glob.glob(os.path.join(translations_path, '*.json')))
        if not shards:
            print(f"No *.json shards found in {translations_path}", file=sys.stderr)
            return 2
        for i, shard in enumerate(shards, 1):
            with open(shard, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
            if not isinstance(data, dict):
                print(f"Shard {shard} is not a JSON object; skipping", file=sys.stderr)
                continue
            translations.update(data)
            log(
                f"[3_apply_translations] Loaded shard {i}/{len(shards)}: "
                f"{os.path.basename(shard)} ({len(data)} entries, "
                f"cumulative {len(translations)})"
            )
    else:
        with open(translations_path, 'r', encoding='utf-8') as fh:
            translations = json.load(fh)
        log(f"[3_apply_translations] Loaded single dictionary: {len(translations)} entries")

    text_files = set()
    for pattern in TEXT_TARGETS:
        text_files.update(glob.glob(os.path.join(root, pattern)))

    font_files = set()
    for pattern in FONT_TARGETS:
        font_files.update(glob.glob(os.path.join(root, pattern)))

    text_files_sorted = sorted(text_files)
    log(
        f"[3_apply_translations] Applying translations to {len(text_files_sorted)} "
        f"text files (slides + notes)..."
    )

    t_text = time.monotonic()
    slides_updated = 0
    notes_updated = 0
    progress_every = max(1, len(text_files_sorted) // 10)
    for i, path in enumerate(text_files_sorted, 1):
        with open(path, 'r', encoding='utf-8') as fh:
            content = fh.read()
        original = content
        content = _replace_text(content, translations)
        content = _LANG_RE.sub('lang="ja-JP"', content)
        if content != original:
            if 'notesSlides' in path.replace('\\', '/'):
                notes_updated += 1
            else:
                slides_updated += 1
        with open(path, 'w', encoding='utf-8') as fh:
            fh.write(content)
        vlog(f"[3_apply_translations]   text {i}/{len(text_files_sorted)}: {os.path.basename(path)}")
        if i % progress_every == 0:
            log(
                f"[3_apply_translations] Text replace progress: "
                f"{i}/{len(text_files_sorted)} "
                f"(slides updated={slides_updated}, notes updated={notes_updated})"
            )

    translated_count = slides_updated + notes_updated
    t_text_done = time.monotonic()
    log(
        f"[3_apply_translations] Text/lang replace done in "
        f"{t_text_done - t_text:.1f}s"
    )

    font_files_sorted = sorted(font_files)
    log(
        f"[3_apply_translations] Applying Yu Gothic UI font to "
        f"{len(font_files_sorted)} font-target files..."
    )

    t_font = time.monotonic()
    font_changed = 0
    progress_every_f = max(1, len(font_files_sorted) // 10)
    for i, path in enumerate(font_files_sorted, 1):
        with open(path, 'r', encoding='utf-8') as fh:
            content = fh.read()
        new_content = _apply_fonts(content, ea_tag)
        if new_content != content:
            font_changed += 1
        with open(path, 'w', encoding='utf-8') as fh:
            fh.write(new_content)
        vlog(f"[3_apply_translations]   font {i}/{len(font_files_sorted)}: {os.path.basename(path)}")
        if i % progress_every_f == 0:
            log(
                f"[3_apply_translations] Font replace progress: "
                f"{i}/{len(font_files_sorted)} (changed={font_changed})"
            )

    t_font_done = time.monotonic()
    log(
        f"[3_apply_translations] Font replace done in "
        f"{t_font_done - t_font:.1f}s"
    )

    print(f"Translation entries: {len(translations)}")
    print(f"Text/lang updated files: {translated_count} (slides: {slides_updated}, notes: {notes_updated})")
    print(f"Font updated files: {font_changed} (typeface: {font['typeface']})")
    log(
        f"[3_apply_translations] Total elapsed: "
        f"{time.monotonic() - t0:.1f}s"
    )

    # Final check: surface <a:t> runs that still look like untranslated English
    # so the caller can iterate. Heuristic: contains 3+ consecutive ASCII
    # letters AND no CJK character (Hiragana/Katakana/CJK Unified Ideographs).
    _EN_RUN_RE = re.compile(r'[A-Za-z]{3,}')
    _CJK_RE = re.compile(r'[\u3040-\u30ff\u3400-\u9fff\uff66-\uff9f]')
    residual: list[tuple[str, str]] = []
    for path in sorted(text_files):
        with open(path, 'r', encoding='utf-8') as fh:
            content = fh.read()
        for t in _TEXT_RE.findall(content):
            stripped = t.strip()
            if not stripped:
                continue
            if _CJK_RE.search(stripped):
                continue
            if not _EN_RUN_RE.search(stripped):
                continue
            # Likely-English run with no Japanese characters.
            key = _normalize_key(stripped)
            if key in {_normalize_key(k) for k in translations}:
                # Was in dict but unchanged -> dictionary value equals source
                # (intentional pass-through, e.g. product names). Skip.
                continue
            residual.append((os.path.basename(path), stripped))
    if residual:
        print(
            f"[3_apply_translations] WARNING: {len(residual)} <a:t> run(s) still "
            "contain English-looking text with no CJK characters. Add them to "
            "the translation dictionary and re-run, or accept as intentional:",
            file=sys.stderr,
        )
        # Cap printed lines to avoid log explosion; full count is above.
        for name, text in residual[:50]:
            preview = text if len(text) <= 160 else text[:157] + '...'
            print(f"  {name}: {preview!r}", file=sys.stderr)
        if len(residual) > 50:
            print(f"  ... and {len(residual) - 50} more", file=sys.stderr)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
