"""Unit tests for ``lib._text_normalize.normalize_text``.

The extraction and apply phases must agree on canonical form for dictionary
keys; these tests guard the small set of rules documented in the module.
"""
from __future__ import annotations

from lib._text_normalize import normalize_text  # type: ignore[import-not-found]


class TestPunctuationMap:
    def test_nbsp_collapses_to_space(self):
        assert normalize_text("foo\u00a0bar") == "foo bar"

    def test_ideographic_space_collapses_to_space(self):
        assert normalize_text("a\u3000b") == "a b"

    def test_nonbreaking_hyphen_to_ascii(self):
        assert normalize_text("on\u2011premises") == "on-premises"

    def test_hyphen_u2010_to_ascii(self):
        assert normalize_text("on\u2010premises") == "on-premises"

    def test_minus_sign_u2212_to_ascii(self):
        assert normalize_text("3\u22122") == "3-2"


class TestHtmlEntities:
    def test_amp_unescaped(self):
        assert normalize_text("R&amp;D") == "R&D"

    def test_lt_gt_unescaped(self):
        assert normalize_text("&lt;tag&gt;") == "<tag>"

    def test_numeric_entity_unescaped(self):
        assert normalize_text("&#65;BC") == "ABC"


class TestWhitespace:
    def test_strip_leading_trailing(self):
        assert normalize_text("  hello  ") == "hello"

    def test_collapse_internal_runs(self):
        assert normalize_text("a   b\t\tc") == "a b c"

    def test_newline_normalized(self):
        assert normalize_text("line1\nline2") == "line1 line2"

    def test_empty_input(self):
        assert normalize_text("") == ""

    def test_whitespace_only_returns_empty(self):
        assert normalize_text("   \t\u00a0  ") == ""


class TestIdempotence:
    def test_already_normalized_unchanged(self):
        s = "Authentication is required."
        assert normalize_text(s) == s

    def test_double_normalize_idempotent(self):
        raw = "  on\u2011premises\u00a0 &amp;  more  "
        once = normalize_text(raw)
        assert normalize_text(once) == once
