"""Vietnamese collation with an explicit, deterministic fallback."""

from __future__ import annotations

import importlib
import unicodedata
from collections.abc import Iterable
from typing import Literal

from unicode_utils import comparison_text

NameMode = Literal["full_name", "given_name"]
Backend = Literal["auto", "icu", "fallback"]

VIETNAMESE_ALPHABET = (
    "a", "ă", "â", "b", "c", "d", "đ", "e", "ê", "g", "h", "i",
    "k", "l", "m", "n", "o", "ô", "ơ", "p", "q", "r", "s", "t",
    "u", "ư", "v", "x", "y",
)
TONE_RANK = {
    "": 0,
    "\N{COMBINING GRAVE ACCENT}": 1,
    "\N{COMBINING HOOK ABOVE}": 2,
    "\N{COMBINING TILDE}": 3,
    "\N{COMBINING ACUTE ACCENT}": 4,
    "\N{COMBINING DOT BELOW}": 5,
}
_TONE_MARKS = frozenset(TONE_RANK) - {""}
_VARIANT_MARKS = {
    ("a", "\N{COMBINING BREVE}"): "ă",
    ("a", "\N{COMBINING CIRCUMFLEX ACCENT}"): "â",
    ("e", "\N{COMBINING CIRCUMFLEX ACCENT}"): "ê",
    ("o", "\N{COMBINING CIRCUMFLEX ACCENT}"): "ô",
    ("o", "\N{COMBINING HORN}"): "ơ",
    ("u", "\N{COMBINING HORN}"): "ư",
}
_LETTER_RANK = {letter: rank for rank, letter in enumerate(VIETNAMESE_ALPHABET)}


class ICUUnavailableError(RuntimeError):
    """Raised when ICU was requested but cannot provide Vietnamese collation."""


def _split_combining_sequences(text: str) -> list[str]:
    sequences: list[str] = []
    for character in unicodedata.normalize("NFD", text.casefold()):
        if unicodedata.combining(character) and sequences:
            sequences[-1] += character
        else:
            sequences.append(character)
    return sequences


def _resolve_letter(sequence: str) -> tuple[str, str] | None:
    if not sequence:
        return None
    base, *marks = sequence
    variant_marks = [mark for mark in marks if mark not in _TONE_MARKS]
    tone_marks = [mark for mark in marks if mark in _TONE_MARKS]
    if len(tone_marks) > 1:
        return None

    if variant_marks:
        if len(variant_marks) != 1:
            return None
        letter = _VARIANT_MARKS.get((base, variant_marks[0]))
        if letter is None:
            return None
    elif base == "đ":
        letter = "đ"
    else:
        letter = base

    if letter not in _LETTER_RANK:
        return None
    tone = tone_marks[0] if tone_marks else ""
    if tone and letter not in "aăâeêioôơuưy":
        return None
    return letter, tone


def fallback_sort_key(value: str) -> tuple[tuple[int, ...], tuple[int, ...], tuple[int, ...]]:
    """Build a documented fallback key: letters, tones, then other code points."""
    text = comparison_text(value).casefold()
    letters: list[int] = []
    tones: list[int] = []
    other_code_points: list[int] = []

    for sequence in _split_combining_sequences(text):
        resolved = _resolve_letter(sequence)
        if resolved is None:
            other_code_points.extend(ord(character) for character in sequence)
            continue
        letter, tone = resolved
        letters.append(_LETTER_RANK[letter])
        if letter in "aăâeêioôơuưy":
            tones.append(TONE_RANK[tone])

    return tuple(letters), tuple(tones), tuple(other_code_points)


def _given_name(value: str) -> str:
    words = comparison_text(value).split()
    return words[-1] if words else ""


class VietnameseSorter:
    """Sort text with ICU `vi_VN` when available, or the explicit fallback."""

    def __init__(self, backend: Backend = "auto") -> None:
        if backend not in ("auto", "icu", "fallback"):
            raise ValueError("backend must be 'auto', 'icu', or 'fallback'")
        self.backend: Literal["icu", "fallback"] = "fallback"
        self.backend_locale: str | None = None
        self.fallback_reason: str | None = None
        self._collator = None

        if backend == "fallback":
            self.fallback_reason = "Fallback backend selected explicitly."
            return

        try:
            icu = importlib.import_module("icu")
        except ModuleNotFoundError as error:
            if error.name != "icu":
                raise
            if backend == "icu":
                raise ICUUnavailableError("PyICU is not installed.") from error
            self.fallback_reason = "PyICU is not installed."
            return

        locale = icu.Locale("vi_VN")
        collator = icu.Collator.createInstance(locale)
        try:
            valid_locale = collator.getLocale(icu.ULocDataLocaleType.VALID_LOCALE)
        except AttributeError as error:
            raise ICUUnavailableError(
                "This PyICU build cannot report the collator's resolved locale."
            ) from error
        resolved_locale = valid_locale.getName().replace("-", "_").lower()
        if not resolved_locale.startswith("vi"):
            if backend == "icu":
                raise ICUUnavailableError(
                    f"ICU resolved vi_VN to unsupported locale {valid_locale.getName()!r}."
                )
            self.fallback_reason = (
                f"ICU resolved vi_VN to unsupported locale {valid_locale.getName()!r}."
            )
            return

        try:
            collator.setStrength(icu.Collator.SECONDARY)
            collator.setAttribute(
                icu.UCollAttribute.NORMALIZATION_MODE,
                icu.UCollAttributeValue.ON,
            )
        except AttributeError as error:
            raise ICUUnavailableError(
                "This PyICU build does not expose required collation settings."
            ) from error

        self.backend = "icu"
        self.backend_locale = valid_locale.getName()
        self._collator = collator

    def key(self, value: str, mode: NameMode = "full_name") -> tuple | bytes:
        if not isinstance(value, str):
            raise TypeError("value must be a string")
        if mode not in ("full_name", "given_name"):
            raise ValueError("mode must be 'full_name' or 'given_name'")

        normalized = comparison_text(value)
        if self.backend == "icu":
            assert self._collator is not None
            if mode == "given_name":
                return (
                    self._collator.getSortKey(_given_name(normalized)),
                    self._collator.getSortKey(normalized),
                )
            return self._collator.getSortKey(normalized)

        if mode == "given_name":
            return fallback_sort_key(_given_name(normalized)), fallback_sort_key(normalized)
        return fallback_sort_key(normalized)

    def sort(
        self,
        values: Iterable[str],
        *,
        descending: bool = False,
        blanks_last: bool = True,
        mode: NameMode = "full_name",
    ) -> list[str]:
        if mode not in ("full_name", "given_name"):
            raise ValueError("mode must be 'full_name' or 'given_name'")
        indexed = list(enumerate(values))
        for _, value in indexed:
            if not isinstance(value, str):
                raise TypeError("all values must be strings")

        blank_rows = [(index, value) for index, value in indexed if not value.strip()]
        nonblank_rows = [(index, value) for index, value in indexed if value.strip()]
        sorted_rows = sorted(
            nonblank_rows,
            key=lambda item: self.key(item[1], mode),
            reverse=descending,
        )
        blanks = [value for _, value in blank_rows]
        nonblanks = [value for _, value in sorted_rows]
        return nonblanks + blanks if blanks_last else blanks + nonblanks


def sort_vietnamese(
    values: Iterable[str],
    *,
    descending: bool = False,
    blanks_last: bool = True,
    mode: NameMode = "full_name",
    backend: Backend = "auto",
) -> list[str]:
    """Convenience function for sorting strings using the selected backend."""
    return VietnameseSorter(backend).sort(
        values,
        descending=descending,
        blanks_last=blanks_last,
        mode=mode,
    )
