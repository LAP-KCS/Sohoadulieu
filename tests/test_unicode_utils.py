import unittest

from unicode_utils import comparison_text, normalize_nfc


class UnicodeUtilsTests(unittest.TestCase):
    def test_nfc_normalizes_equivalent_compositions(self) -> None:
        composed = "Nguyễn"
        decomposed = "Nguye\u0302\u0303n"

        self.assertEqual(normalize_nfc(composed), composed)
        self.assertEqual(normalize_nfc(decomposed), composed)
        self.assertEqual(decomposed, "Nguye\u0302\u0303n")

    def test_comparison_text_strips_only_when_requested(self) -> None:
        self.assertEqual(comparison_text("  Nguyễn  "), "Nguyễn")
        self.assertEqual(
            comparison_text("  Nguyễn  ", strip_outer_whitespace=False),
            "  Nguyễn  ",
        )

    def test_non_string_input_is_rejected(self) -> None:
        with self.assertRaises(TypeError):
            normalize_nfc(42)  # type: ignore[arg-type]


if __name__ == "__main__":
    unittest.main()
