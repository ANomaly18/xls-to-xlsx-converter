from __future__ import annotations

import socket
from pathlib import Path

import pytest
import xlwt
from openpyxl import load_workbook

import xls_to_xlsx_converter.converter as converter_module
from xls_to_xlsx_converter.cli import main as cli_main
from xls_to_xlsx_converter.converter import (
    _infer_value,
    convert_file,
    detect_file_format,
)


def _write_generic_html_xls(path: Path) -> None:
    path.write_text(
        """<!doctype html>
<html xmlns:x="urn:schemas-microsoft-com:office:excel">
<head>
  <meta charset="utf-8">
  <x:Name>Sample/Report:*?[] With A Long Name</x:Name>
  <link rel="stylesheet" href="https://invalid.example/private.css">
  <style>.heading { background-color: #DDEEFF; font-weight: bold; }</style>
</head>
<body>
<table>
  <tr class="heading"><th>Reference ID</th><th>Label</th><th>Value</th><th>Literal</th><th>Link</th></tr>
  <tr><td>9999999999999999</td><td>Synthetic row A</td><td>123.5</td><td>=2+2</td><td><a href="javascript:alert(1)">Unsafe</a></td></tr>
  <tr><td>0000000000000000</td><td>Synthetic row B</td><td>7</td><td>plain</td><td><a href="https://example.com/help">Help</a></td></tr>
</table>
</body>
</html>
""",
        encoding="utf-8",
    )


def _write_binary_xls(path: Path) -> None:
    workbook = xlwt.Workbook(encoding="utf-8")
    sheet = workbook.add_sheet("Synthetic")
    heading = xlwt.easyxf("font: bold on; align: horiz center")
    sheet.write_merge(0, 0, 0, 2, "Synthetic workbook", heading)
    sheet.write(1, 0, "Reference ID", heading)
    sheet.write(1, 1, "Label", heading)
    sheet.write(1, 2, "Amount", heading)
    sheet.write(2, 0, "9999999999999999")
    sheet.write(2, 1, "Synthetic row")
    sheet.write(2, 2, 42.5)
    workbook.save(str(path))


def test_generic_html_conversion_is_offline_and_safe(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "generic.xls"
    output = tmp_path / "generic.xlsx"
    _write_generic_html_xls(source)

    def deny_network(*_args, **_kwargs):
        raise AssertionError("conversion attempted a network connection")

    monkeypatch.setattr(socket, "create_connection", deny_network)
    result = convert_file(source, output)

    assert result == output.resolve()
    workbook = load_workbook(output, data_only=False)
    sheet = workbook.active
    assert len(sheet.title) <= 31
    assert not any(character in sheet.title for character in r"[]:*?/\\")
    assert sheet["A2"].value == "9999999999999999"
    assert sheet["C2"].value == 123.5
    assert sheet["D2"].value == "'=2+2"
    assert sheet["D2"].data_type != "f"
    assert sheet["E2"].hyperlink is None
    assert sheet["E3"].hyperlink.target == "https://example.com/help"


def test_binary_conversion_preserves_core_values_and_formatting(tmp_path: Path) -> None:
    source = tmp_path / "binary.xls"
    output = tmp_path / "binary.xlsx"
    _write_binary_xls(source)

    assert detect_file_format(source) == "binary"
    convert_file(source, output)

    workbook = load_workbook(output, data_only=False)
    sheet = workbook["Synthetic"]
    assert "A1:C1" in {str(item) for item in sheet.merged_cells.ranges}
    assert sheet["A1"].font.bold is True
    assert sheet["A3"].value == "9999999999999999"
    assert sheet["C3"].value == 42.5


def test_existing_output_requires_explicit_overwrite(tmp_path: Path) -> None:
    source = tmp_path / "source.xls"
    output = tmp_path / "output.xlsx"
    _write_generic_html_xls(source)
    convert_file(source, output)

    with pytest.raises(FileExistsError):
        convert_file(source, output)

    convert_file(source, output, overwrite=True)
    assert output.is_file()


def test_failed_conversion_leaves_no_partial_output(tmp_path: Path) -> None:
    source = tmp_path / "invalid.xls"
    output = tmp_path / "invalid.xlsx"
    source.write_text("<html><body>no table</body></html>", encoding="utf-8")

    with pytest.raises(ValueError, match="tidak memiliki elemen <table>"):
        convert_file(source, output)

    assert not output.exists()
    assert list(tmp_path.glob(".invalid.*.xlsx")) == []


def test_file_and_expansion_limits_fail_closed(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    oversized = tmp_path / "oversized.xls"
    oversized.write_bytes(b"12345")
    monkeypatch.setattr(converter_module, "MAX_FILE_SIZE", 4)
    with pytest.raises(ValueError, match="File terlalu besar"):
        detect_file_format(oversized)

    expanded = tmp_path / "expanded.xls"
    expanded.write_text(
        "<html><body><table><tr><td colspan='5'>x</td></tr></table></body></html>",
        encoding="utf-8",
    )
    monkeypatch.setattr(converter_module, "MAX_FILE_SIZE", 50 * 1024 * 1024)
    monkeypatch.setattr(converter_module, "MAX_HTML_EXPANDED_CELLS", 4)
    with pytest.raises(ValueError, match="expanded-cell safety limit"):
        convert_file(expanded, tmp_path / "expanded.xlsx")


def test_type_inference_preserves_identifiers_and_neutralizes_formulas() -> None:
    assert _infer_value("9999999999999999") == "9999999999999999"
    assert _infer_value("000001") == "000001"
    assert _infer_value("=SUM(A1:A2)") == "'=SUM(A1:A2)"
    assert _infer_value("2026-08-24").date().isoformat() == "2026-08-24"


def test_cli_converts_synthetic_input(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    source = tmp_path / "cli.xls"
    _write_generic_html_xls(source)

    assert cli_main([str(source)]) == 0
    output = tmp_path / "cli_converted.xlsx"
    assert output.is_file()
    assert "Converted:" in capsys.readouterr().out
