"""Tests for ``3_apply_translations.py``.

Covers:
  - <a:t> text replacement via the translation dictionary
  - lang="en-*" -> lang="ja-JP" rewrite (slides + notesSlides)
  - <a:ea> Yu Gothic UI insertion / replacement
  - English-residual WARNING heuristic (printed to stderr)
  - Translation directory loading from multiple JSON shards
"""
from __future__ import annotations

import json


def _write_dict(tmp_path, entries: dict[str, str], name: str = "dict.json") -> str:
    p = tmp_path / name
    p.write_text(json.dumps(entries, ensure_ascii=False), encoding="utf-8")
    return str(p)


def test_text_replaced_via_single_json(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(
        unpacked_dir,
        1,
        [("en-US", "Authentication"), ("en-US", "Authorization")],
    )
    dict_path = _write_dict(tmp_path, {"Authentication": "認証", "Authorization": "認可"})
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), dict_path)
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert "認証" in content
    assert "認可" in content
    assert "Authentication" not in content
    assert "Authorization" not in content


def test_lang_attribute_rewritten_to_ja_jp(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(
        unpacked_dir,
        1,
        [("en-US", "Hello"), ("en-GB", "World")],
    )
    dict_path = _write_dict(tmp_path, {"Hello": "こんにちは", "World": "世界"})
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), dict_path)
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert 'lang="en-US"' not in content
    assert 'lang="en-GB"' not in content
    assert content.count('lang="ja-JP"') == 2


def test_ea_font_inserted_when_missing(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(unpacked_dir, 1, [("en-US", "Hello")])
    dict_path = _write_dict(tmp_path, {"Hello": "こんにちは"})
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), dict_path)
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert "Yu Gothic UI" in content
    assert "<a:ea " in content
    # OOXML schema requires <a:ea> AFTER <a:latin>.
    assert content.index("<a:latin") < content.index("<a:ea ")


def test_ea_font_replaced_when_present(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    slide_path = unpacked_dir / "ppt" / "slides" / "slide1.xml"
    slide_path.write_text(
        '<?xml version="1.0"?><p:sld '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r>'
        '<a:rPr lang="en-US">'
        '<a:latin typeface="Arial"/>'
        '<a:ea typeface="MS Gothic"/>'
        '</a:rPr><a:t>Hi</a:t>'
        "</a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>",
        encoding="utf-8",
    )
    dict_path = _write_dict(tmp_path, {"Hi": "やあ"})
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), dict_path)
    assert rc == 0, err
    content = slide_path.read_text("utf-8")
    assert "MS Gothic" not in content
    assert 'typeface="Yu Gothic UI"' in content


def test_notes_slides_also_translated(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(unpacked_dir, 1, [("en-US", "Body text")])
    write_slide_fn(unpacked_dir, 1, [("en-US", "Notes text")], in_notes=True)
    dict_path = _write_dict(
        tmp_path, {"Body text": "本文", "Notes text": "ノート"}
    )
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), dict_path)
    assert rc == 0, err
    slide = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    notes = (unpacked_dir / "ppt" / "notesSlides" / "notesSlide1.xml").read_text("utf-8")
    assert "本文" in slide
    assert "ノート" in notes
    assert 'lang="ja-JP"' in notes


def test_residual_warning_for_untranslated_english(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(
        unpacked_dir,
        1,
        [
            ("en-US", "Authentication"),
            ("en-US", "UntranslatedResidualText"),
        ],
    )
    dict_path = _write_dict(tmp_path, {"Authentication": "認証"})
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), dict_path)
    assert rc == 0, err
    assert "WARNING" in err
    assert "UntranslatedResidualText" in err


def test_no_residual_warning_when_all_translated(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(unpacked_dir, 1, [("en-US", "Hi")])
    dict_path = _write_dict(tmp_path, {"Hi": "やあ"})
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), dict_path)
    assert rc == 0, err
    assert "WARNING" not in err


def test_multi_shard_directory_merged(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    write_slide_fn(
        unpacked_dir,
        1,
        [("en-US", "Alpha"), ("en-US", "Beta"), ("en-US", "Gamma")],
    )
    shards = tmp_path / "shards"
    shards.mkdir()
    (shards / "chunk_001.json").write_text(
        json.dumps({"Alpha": "アルファ"}, ensure_ascii=False), "utf-8"
    )
    (shards / "chunk_002.json").write_text(
        json.dumps({"Beta": "ベータ", "Gamma": "ガンマ"}, ensure_ascii=False), "utf-8"
    )
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), str(shards))
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert "アルファ" in content and "ベータ" in content and "ガンマ" in content


def test_normalization_aware_lookup(
    unpacked_dir, write_slide_fn, run_script_fn, tmp_path
):
    # XML contains U+2011; dictionary key uses ASCII hyphen. Normalization
    # should reconcile the two.
    write_slide_fn(unpacked_dir, 1, [("en-US", "on\u2011premises")])
    dict_path = _write_dict(tmp_path, {"on-premises": "オンプレミス"})
    rc, _, err = run_script_fn("3_apply_translations.py", str(unpacked_dir), dict_path)
    assert rc == 0, err
    content = (unpacked_dir / "ppt" / "slides" / "slide1.xml").read_text("utf-8")
    assert "オンプレミス" in content
