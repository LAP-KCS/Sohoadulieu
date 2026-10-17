import unittest

from vietnamese_sort import (
    ICUUnavailableError,
    VIETNAMESE_ALPHABET,
    VietnameseSorter,
    fallback_sort_key,
)


class VietnameseSortTests(unittest.TestCase):
    def setUp(self) -> None:
        self.sorter = VietnameseSorter(backend="fallback")

    def test_base_alphabet_order(self) -> None:
        reversed_alphabet = list(reversed(VIETNAMESE_ALPHABET))
        actual = self.sorter.sort(reversed_alphabet)
        self.assertEqual(actual, list(VIETNAMESE_ALPHABET))

    def test_base_letter_distinctions(self) -> None:
        self.assertEqual(
            self.sorter.sort(["Â", "Ă", "A", "Đ", "D", "Ơ", "Ô", "O", "Ư", "U"]),
            ["A", "Ă", "Â", "D", "Đ", "O", "Ô", "Ơ", "U", "Ư"],
        )

    def test_tone_order(self) -> None:
        self.assertEqual(
            self.sorter.sort(["mạ", "má", "mã", "mả", "mà", "ma"]),
            ["ma", "mà", "mả", "mã", "má", "mạ"],
        )

    def test_tone_order_on_vowel_variants(self) -> None:
        self.assertEqual(
            self.sorter.sort(["mặ", "mắ", "mẵ", "mẳ", "mằ", "mă"]),
            ["mă", "mằ", "mẳ", "mẵ", "mắ", "mặ"],
        )

    def test_nfc_and_nfd_have_equal_keys(self) -> None:
        self.assertEqual(
            fallback_sort_key("Nguyễn"),
            fallback_sort_key("Nguye\u0302\u0303n"),
        )

    def test_case_is_ignored_and_equal_keys_are_stable(self) -> None:
        values = ["AN", "an", "An", "Án"]
        self.assertEqual(self.sorter.sort(values), ["AN", "an", "An", "Án"])

    def test_outer_whitespace_is_ignored_but_source_is_preserved(self) -> None:
        values = ["  An ", "Án", "An"]
        self.assertEqual(self.sorter.sort(values), ["  An ", "An", "Án"])
        self.assertEqual(values, ["  An ", "Án", "An"])

    def test_punctuation_uses_unicode_code_point_tie_break(self) -> None:
        self.assertEqual(
            self.sorter.sort(["An-An", "An An", "AnAn"]),
            ["AnAn", "An An", "An-An"],
        )

    def test_unknown_combining_marks_remain_distinguishable(self) -> None:
        self.assertNotEqual(fallback_sort_key("b"), fallback_sort_key("b\u0301"))

    def test_full_name_sample_expected_order(self) -> None:
        names = [
            "Nguyễn An",
            "Nguyễn Ánh",
            "Nguyễn Bình",
            "Nguyễn Đạt",
            "Nguyễn Đăng",
            "Nguyễn Đức",
            "Nguyễn Hùng",
            "Nguyễn Hương",
            "Nguyễn Minh",
            "Nguyễn Nam",
            "Nguyễn Ngọc",
            "Nguyễn Thị",
            "Nguyễn Văn",
        ]
        self.assertEqual(
            self.sorter.sort(names),
            [
                "Nguyễn An",
                "Nguyễn Ánh",
                "Nguyễn Bình",
                "Nguyễn Đạt",
                "Nguyễn Đăng",
                "Nguyễn Đức",
                "Nguyễn Hùng",
                "Nguyễn Hương",
                "Nguyễn Minh",
                "Nguyễn Nam",
                "Nguyễn Ngọc",
                "Nguyễn Thị",
                "Nguyễn Văn",
            ],
        )

    def test_given_name_mode_uses_final_word_then_full_name(self) -> None:
        self.assertEqual(
            self.sorter.sort(
                ["Nguyễn Bình", "Trần An", "Lê An", "Phạm Bình"],
                mode="given_name",
            ),
            ["Lê An", "Trần An", "Nguyễn Bình", "Phạm Bình"],
        )

    def test_blank_position_and_descending_are_independent(self) -> None:
        values = ["B", " ", "A", ""]
        self.assertEqual(
            self.sorter.sort(values, descending=True, blanks_last=True),
            ["B", "A", " ", ""],
        )
        self.assertEqual(
            self.sorter.sort(values, descending=True, blanks_last=False),
            [" ", "", "B", "A"],
        )

    def test_equal_keys_remain_stable_when_descending(self) -> None:
        values = ["an", "AN", "An"]
        self.assertEqual(self.sorter.sort(values, descending=True), values)

    def test_invalid_inputs_are_rejected(self) -> None:
        with self.assertRaises(TypeError):
            self.sorter.sort(["An", 5])  # type: ignore[list-item]
        with self.assertRaises(ValueError):
            self.sorter.sort(["An"], mode="surname")  # type: ignore[arg-type]
        with self.assertRaises(ValueError):
            self.sorter.sort([], mode="surname")  # type: ignore[arg-type]

    def test_explicit_icu_request_reports_missing_dependency(self) -> None:
        try:
            import icu  # noqa: F401
        except ModuleNotFoundError:
            with self.assertRaises(ICUUnavailableError):
                VietnameseSorter(backend="icu")

    def test_auto_backend_exposes_fallback_reason(self) -> None:
        sorter = VietnameseSorter(backend="auto")
        self.assertIn(sorter.backend, ("icu", "fallback"))
        if sorter.backend == "fallback":
            self.assertTrue(sorter.fallback_reason)


if __name__ == "__main__":
    unittest.main()
