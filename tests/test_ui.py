from pathlib import Path

import pytest

from ui import VietnameseExcelSorterApp, _format_preview_value


def test_preview_formatting_keeps_text_and_formats_dates() -> None:
    from datetime import date

    assert _format_preview_value(None) == ""
    assert _format_preview_value("Nguyễn Ánh") == "Nguyễn Ánh"
    assert _format_preview_value(date(2026, 10, 17)) == "2026-10-17"


def test_copy_new_file_preserves_bytes_and_refuses_overwrite(tmp_path: Path) -> None:
    source = tmp_path / "preview.xlsx"
    output = tmp_path / "result.xlsx"
    source.write_bytes(b"test workbook bytes")

    VietnameseExcelSorterApp._copy_new_file(source, output)

    assert output.read_bytes() == source.read_bytes()
    with pytest.raises(FileExistsError):
        VietnameseExcelSorterApp._copy_new_file(source, output)
    assert output.read_bytes() == b"test workbook bytes"
