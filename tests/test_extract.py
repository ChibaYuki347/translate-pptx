"""Tests for ``2_extract_texts.py``.

Covers default (filename-prefixed) output, ``--unique`` mode, normalization
(NBSP, U+2011, HTML entities, whitespace collapse), and skipping of empty
runs.
"""
from __future__ import annotations

from pathlib import Path


def test_unique_mode_emits_distinct_texts(unpacked_dir, write_slide_fn, run_script_fn):
    write_slide_fn(
        unpacked_dir,
        1,
        [
            ("en-US", "Authentication"),
            ("en-US", "Authorization"),
            ("en-US", "Authentication"),
        ],
    )
    rc, out, err = run_script_fn("2_extract_texts.py", "--unique", str(unpacked_dir))
    assert rc == 0, err
    lines = [ln for ln in out.splitlines() if ln.strip()]
    assert sorted(lines) == ["Authentication", "Authorization"]
    assert "[2_extract_texts] unique texts: 2" in err


def test_default_mode_prefixes_filename(unpacked_dir, write_slide_fn, run_script_fn):
    write_slide_fn(unpacked_dir, 1, [("en-US", "Hello")])
    rc, out, err = run_script_fn("2_extract_texts.py", str(unpacked_dir))
    assert rc == 0, err
    assert "slide1.xml: 'Hello'" in out


def test_notes_slides_included(unpacked_dir, write_slide_fn, run_script_fn):
    write_slide_fn(unpacked_dir, 1, [("en-US", "Slide body")])
    write_slide_fn(unpacked_dir, 1, [("en-US", "Notes content")], in_notes=True)
    rc, out, _ = run_script_fn("2_extract_texts.py", "--unique", str(unpacked_dir))
    assert rc == 0
    texts = set(ln for ln in out.splitlines() if ln.strip())
    assert "Slide body" in texts
    assert "Notes content" in texts


def test_normalization_applied(unpacked_dir, write_slide_fn, run_script_fn):
    # NBSP, non-breaking hyphen, HTML entity. After normalize_text these all
    # collapse to a single canonical form.
    write_slide_fn(
        unpacked_dir,
        1,
        [
            ("en-US", "on\u2011premises"),
            ("en-US", "on\u00a0premises"),
            ("en-US", "R&amp;D"),
        ],
    )
    rc, out, _ = run_script_fn("2_extract_texts.py", "--unique", str(unpacked_dir))
    assert rc == 0
    texts = set(ln for ln in out.splitlines() if ln.strip())
    assert "on-premises" in texts
    assert "on premises" in texts
    assert "R&D" in texts


def test_empty_runs_skipped(unpacked_dir, write_slide_fn, run_script_fn):
    write_slide_fn(unpacked_dir, 1, [("en-US", "  "), ("en-US", "real text")])
    rc, out, _ = run_script_fn("2_extract_texts.py", "--unique", str(unpacked_dir))
    assert rc == 0
    lines = [ln for ln in out.splitlines() if ln.strip()]
    assert lines == ["real text"]


def test_sorted_unique_output(unpacked_dir, write_slide_fn, run_script_fn):
    write_slide_fn(unpacked_dir, 1, [("en-US", "Zoo")])
    write_slide_fn(unpacked_dir, 2, [("en-US", "Apple")])
    write_slide_fn(unpacked_dir, 3, [("en-US", "Banana")])
    rc, out, _ = run_script_fn("2_extract_texts.py", "--unique", str(unpacked_dir))
    assert rc == 0
    lines = [ln for ln in out.splitlines() if ln.strip()]
    assert lines == sorted(lines)
    assert set(lines) == {"Apple", "Banana", "Zoo"}
