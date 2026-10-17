"""Unicode helpers that keep source text separate from comparison keys."""

import unicodedata


def normalize_nfc(value: str) -> str:
    """Return the NFC form of text without changing the caller's value."""
    if not isinstance(value, str):
        raise TypeError("value must be a string")
    return unicodedata.normalize("NFC", value)


def comparison_text(value: str, *, strip_outer_whitespace: bool = True) -> str:
    """Prepare text for comparison; never use this result to overwrite cell data."""
    normalized = normalize_nfc(value)
    if strip_outer_whitespace:
        normalized = normalized.strip()
    return normalized
