"""Shared text normalization for translate-pptx extraction & replacement.

The extraction phase (`extract_texts.py`) and the replacement phase
(`apply_translations.py`) MUST agree on how source strings are normalized,
otherwise dictionary lookups silently miss runs containing characters such as
U+2011 NON-BREAKING HYPHEN, U+2010 HYPHEN, or NBSP that look identical to ASCII
hyphen/space but compare unequal.

Normalization rules (kept intentionally small and reversible-in-spirit):
  * HTML entity unescape (``&amp;`` -> ``&`` etc.)
  * Strip leading/trailing whitespace
  * Collapse internal whitespace runs (incl. NBSP / U+3000) to a single ASCII space
  * Map a small set of look-alike punctuation to ASCII equivalents:
        U+00A0 NBSP                -> ' '
        U+2010 HYPHEN              -> '-'
        U+2011 NON-BREAKING HYPHEN -> '-'
        U+2212 MINUS SIGN          -> '-'
"""
from __future__ import annotations

import html
import re

_WS_RE = re.compile(r"[\s\u3000]+")

_PUNCT_MAP = {
    "\u00a0": " ",
    "\u2010": "-",
    "\u2011": "-",
    "\u2212": "-",
}


def normalize_text(s: str) -> str:
    """Return the canonical key form used by both extraction and apply phases."""
    s = html.unescape(s)
    for src, dst in _PUNCT_MAP.items():
        if src in s:
            s = s.replace(src, dst)
    s = _WS_RE.sub(" ", s).strip()
    return s
