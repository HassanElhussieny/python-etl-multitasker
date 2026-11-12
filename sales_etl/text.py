"""Arabic-aware text helpers.

``canonical`` is used only to *match* labels (headers, document types, footer
labels). Original text is always preserved for display and output.
"""

from __future__ import annotations

import re
import unicodedata

_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹٫٬", "01234567890123456789.,")
# tatweel, zero-width chars, bidi marks/embeddings/isolates, BOM
_INVISIBLE = re.compile("[ـ​-‏‪-‮⁦-⁩﻿]")
_DIACRITICS = re.compile("[ً-ٰٟ]")
_LETTER_FOLD = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ى": "ي", "ة": "ه"})
_WHITESPACE = re.compile(r"\s+")


def clean_cell(value: str | None) -> str:
    """Strip surrounding whitespace and invisible formatting characters."""
    if value is None:
        return ""
    return _INVISIBLE.sub("", value).replace("\xa0", " ").strip()


def canonical(value: str | None) -> str:
    """Normalise a label for comparison (spelling variants such as ى/ي, أ/ا fold)."""
    if not value:
        return ""
    text = unicodedata.normalize("NFKC", value)
    text = _INVISIBLE.sub("", text)
    text = _DIACRITICS.sub("", text)
    text = text.translate(_LETTER_FOLD)
    return _WHITESPACE.sub(" ", text).strip()


def ascii_digits(value: str) -> str:
    """Convert Arabic-Indic / Persian digits and separators to ASCII."""
    return value.translate(_DIGITS)
