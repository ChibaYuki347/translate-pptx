"""Tests for ``3_apply_translations.py --direction ja2en`` (Phase 2).

Covers the Japanese -> English direction:
  - <a:t> text replacement via the Japanese -> English dictionary
  - lang="ja-*" -> lang="en-US" rewrite (slides + notesSlides)
  - <a:ea> tags stripped (East-Asian font removed)
  - <a:latin> rewritten to Segoe UI / Segoe UI Semibold so titles do not
    fall back to PowerPoint's default Calibri Light
  - Theme majorFont / minorFont <a:latin> rewritten
  - CJK-residual WARNING heuristic (printed to stderr)
  - en2ja regression (direction-neutral defaults preserved)
"""
from __future__ import annotations

import json


def _write_dict(tmp_path, entries: dict[str, str], name: str = "dict.json") -> str:
    p = tmp_path / name
    p.write_text(json.dumps(entries, ensure_ascii=False), encoding="utf-8")
    return str(p)


def test_ja2en_text_replaced(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(
        unpacked_dir,
        1,
        [("ja-JP", "認証"), ("ja-JP", "認可")],
    )
    dict_path = _write_dict(tmp_path, {"認証": "Authentication", "認可": "Authorization"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert "Authentication" in content
    assert "Authorization" in content
    assert "認証" not in content
    assert "認可" not in content


def test_ja2en_lang_attribute_rewritten_to_en_us(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(
        unpacked_dir,
        1,
        [("ja-JP", "こんにちは"), ("ja-JP", "世界")],
    )
    dict_path = _write_dict(tmp_path, {"こんにちは": "Hello", "世界": "World"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert 'lang="ja-JP"' not in content
    assert content.count('lang="en-US"') == 2


def test_ja2en_ea_stripped_and_latin_rewritten_to_segoe_ui(
    unpacked_dir, run_script_fn, tmp_path
):
    """ja2en path: <a:ea> tags must be removed AND any existing <a:latin>
    must be rewritten to Segoe UI (Segoe UI Semibold when the run is bold).

    Without forcing Latin to Segoe UI, JA->EN titles fall back to PowerPoint's
    default Calibri Light — see the colleague-reported rendering issue.
    """
    slide_path = unpacked_dir / "ppt" / "slides" / "slide1.xml"
    slide_path.write_text(
        '<?xml version="1.0"?><p:sld '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r>'
        '<a:rPr lang="ja-JP">'
        '<a:latin typeface="Segoe Sans Text"/>'
        '<a:ea typeface="Yu Gothic UI"/>'
        '</a:rPr><a:t>こんにちは</a:t>'
        "</a:r><a:r>"
        '<a:rPr lang="ja-JP" b="1">'
        '<a:latin typeface="Segoe Sans Text Semibold"/>'
        '<a:ea typeface="Yu Gothic UI"/>'
        '</a:rPr><a:t>タイトル</a:t>'
        "</a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>",
        encoding="utf-8",
    )
    dict_path = _write_dict(tmp_path, {"こんにちは": "Hello", "タイトル": "Title"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    content = slide_path.read_text("utf-8")
    # <a:ea> tags must be stripped.
    assert "<a:ea " not in content
    assert "Yu Gothic UI" not in content
    # Original Segoe Sans Text Latin must be replaced with Segoe UI variants.
    assert "Segoe Sans Text" not in content
    # Body run -> Segoe UI; bold run -> Segoe UI Semibold.
    assert 'typeface="Segoe UI"' in content
    assert 'typeface="Segoe UI Semibold"' in content


def test_ja2en_inserts_segoe_ui_semibold_for_bold_run_without_latin(
    unpacked_dir, run_script_fn, tmp_path
):
    """Title placeholders often have <a:rPr b="1"> with no explicit <a:latin>,
    relying on theme inheritance. After ja2en stripping, the inheritance
    chain can yield Calibri Light. Insert Segoe UI Semibold explicitly for
    such bold runs so the title renders in the brand font.
    """
    slide_path = unpacked_dir / "ppt" / "slides" / "slide1.xml"
    slide_path.write_text(
        '<?xml version="1.0"?><p:sld '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r>'
        '<a:rPr lang="ja-JP" b="1">'
        '<a:ea typeface="Yu Gothic UI"/>'
        '</a:rPr><a:t>見出し</a:t>'
        "</a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>",
        encoding="utf-8",
    )
    dict_path = _write_dict(tmp_path, {"見出し": "Heading"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    content = slide_path.read_text("utf-8")
    assert "<a:ea " not in content
    # A <a:latin> with Segoe UI Semibold must now be present in the bold run.
    assert 'typeface="Segoe UI Semibold"' in content


def test_ja2en_leaves_inheriting_runs_without_latin(
    unpacked_dir, run_script_fn, tmp_path
):
    """Non-bold runs that had no explicit <a:latin> should remain that way
    so they continue to inherit from layout / theme (which we also rewrite
    to Segoe UI). Inserting a slide-level latin here would override a
    valid bold inheritance from the parent layout.
    """
    slide_path = unpacked_dir / "ppt" / "slides" / "slide1.xml"
    slide_path.write_text(
        '<?xml version="1.0"?><p:sld '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r>'
        '<a:rPr lang="ja-JP">'
        '<a:ea typeface="Yu Gothic UI"/>'
        '</a:rPr><a:t>本文</a:t>'
        "</a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>",
        encoding="utf-8",
    )
    dict_path = _write_dict(tmp_path, {"本文": "Body"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    content = slide_path.read_text("utf-8")
    assert "<a:ea " not in content
    # No <a:latin> should have been inserted on this run (it had no b="1"
    # and no original latin); inheritance handles it.
    assert "<a:latin" not in content


def test_ja2en_theme_major_font_rewritten_to_segoe_ui_semibold(
    unpacked_dir, run_script_fn, tmp_path
):
    """Theme <a:majorFont> Latin must be rewritten to Segoe UI Semibold so
    title placeholders that inherit the heading font render in Semibold."""
    theme_dir = unpacked_dir / "ppt" / "theme"
    theme_dir.mkdir(parents=True, exist_ok=True)
    theme_path = theme_dir / "theme1.xml"
    theme_path.write_text(
        '<?xml version="1.0"?><a:theme '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:themeElements><a:fontScheme name="Office">'
        '<a:majorFont>'
        '<a:latin typeface="Segoe Sans Text Semibold"/>'
        '<a:ea typeface="Yu Gothic UI"/>'
        '<a:cs typeface=""/>'
        "</a:majorFont>"
        '<a:minorFont>'
        '<a:latin typeface="Segoe Sans Text"/>'
        '<a:ea typeface="Yu Gothic UI"/>'
        '<a:cs typeface=""/>'
        "</a:minorFont>"
        "</a:fontScheme></a:themeElements></a:theme>",
        encoding="utf-8",
    )
    # Apply an empty dict; we only care about font rewriting in theme files.
    dict_path = _write_dict(tmp_path, {})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    content = theme_path.read_text("utf-8")
    # <a:ea> stripped from theme blocks too.
    assert "Yu Gothic UI" not in content
    # majorFont -> Semibold; minorFont -> regular Segoe UI.
    import re
    major_block = re.search(
        r"<a:majorFont\b[^>]*>.*?</a:majorFont>", content, re.DOTALL
    )
    minor_block = re.search(
        r"<a:minorFont\b[^>]*>.*?</a:minorFont>", content, re.DOTALL
    )
    assert major_block, content
    assert minor_block, content
    assert 'typeface="Segoe UI Semibold"' in major_block.group(0)
    assert 'typeface="Segoe UI"' in minor_block.group(0)
    # And the minor block doesn't accidentally have Semibold.
    assert "Semibold" not in minor_block.group(0)


def test_ja2en_notes_slides_translated(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(unpacked_dir, 1, [("ja-JP", "本文")])
    write_slide_fn(unpacked_dir, 1, [("ja-JP", "ノート")], in_notes=True)
    dict_path = _write_dict(tmp_path, {"本文": "Body text", "ノート": "Notes text"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    slide = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    notes = (unpacked_dir / "ppt" / "notesSlides" / "notesSlide1.xml").read_text("utf-8")
    assert "Body text" in slide
    assert "Notes text" in notes
    assert 'lang="en-US"' in notes


def test_ja2en_residual_warning_for_untranslated_cjk(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(
        unpacked_dir,
        1,
        [
            ("ja-JP", "認証"),
            ("ja-JP", "未翻訳の日本語"),
        ],
    )
    dict_path = _write_dict(tmp_path, {"認証": "Authentication"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    assert "WARNING" in err
    assert "未翻訳の日本語" in err


def test_ja2en_no_residual_warning_when_all_translated(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(unpacked_dir, 1, [("ja-JP", "やあ")])
    dict_path = _write_dict(tmp_path, {"やあ": "Hi"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    assert "WARNING" not in err


def test_ja2en_does_not_rewrite_en_lang(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    """ja2en must only touch ja-* langs. en-US runs are left alone."""
    write_slide_fn(
        unpacked_dir,
        1,
        [("ja-JP", "認証"), ("en-US", "Microsoft Azure")],
    )
    dict_path = _write_dict(
        tmp_path,
        {"認証": "Authentication", "Microsoft Azure": "Microsoft Azure"},
    )
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "ja2en",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert "Microsoft Azure" in content
    assert content.count('lang="en-US"') == 2


def test_en2ja_default_still_works_without_direction_flag(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    """Regression: omitting --direction must keep the original en2ja behavior
    bit-identical (default value 'en2ja')."""
    write_slide_fn(unpacked_dir, 1, [("en-US", "Hello")])
    dict_path = _write_dict(tmp_path, {"Hello": "こんにちは"})
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), dict_path)
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert "こんにちは" in content
    assert 'lang="ja-JP"' in content
    assert "Yu Gothic UI" in content


def test_en2ja_explicit_direction_flag_works(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(unpacked_dir, 1, [("en-US", "Hello")])
    dict_path = _write_dict(tmp_path, {"Hello": "こんにちは"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "en2ja",
        str(unpacked_dir), dict_path,
    )
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert "こんにちは" in content
    assert "Yu Gothic UI" in content


def test_invalid_direction_rejected(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(unpacked_dir, 1, [("en-US", "Hi")])
    dict_path = _write_dict(tmp_path, {"Hi": "やあ"})
    rc, _, err = run_script_fn(
        "3_apply_translations.py", "--direction", "bogus",
        str(unpacked_dir), dict_path,
    )
    assert rc != 0
    assert "bogus" in err or "invalid choice" in err.lower()


def test_unpack_ja2en_allows_ja_input(
    make_pptx, tmp_path, run_script_fn
):
    """ja2en: _JA-suffixed input (the en2ja output) is the expected starting
    point and must be accepted."""
    src = make_pptx(["Sample"], filename="Deck_JA.pptx")
    out_dir = tmp_path / "unpacked"
    rc, out, err = run_script_fn(
        "1_unpack_pptx.py", "--direction", "ja2en", str(src), str(out_dir),
    )
    assert rc == 0, err or out
    assert (out_dir / "ppt" / "slides").is_dir()


def test_unpack_ja2en_rejects_en_input(
    make_pptx, tmp_path, run_script_fn
):
    """ja2en: _EN-suffixed input is already reverse-translated; refuse."""
    src = make_pptx(["Sample"], filename="Deck_EN.pptx")
    out_dir = tmp_path / "unpacked"
    rc, out, err = run_script_fn(
        "1_unpack_pptx.py", "--direction", "ja2en", str(src), str(out_dir),
    )
    assert rc != 0
    assert "_EN" in err or "already" in err.lower()


def test_unpack_en2ja_rejects_ja_input(
    make_pptx, tmp_path, run_script_fn
):
    """en2ja (default): _JA-suffixed input means double translation; refuse."""
    src = make_pptx(["Sample"], filename="Deck_JA.pptx")
    out_dir = tmp_path / "unpacked"
    rc, out, err = run_script_fn(
        "1_unpack_pptx.py", "--direction", "en2ja", str(src), str(out_dir),
    )
    assert rc != 0
    assert "_JA" in err or "already" in err.lower()
