from __future__ import annotations

import datetime
from pathlib import Path

import pytest
from bs4 import BeautifulSoup

from xls_to_xlsx_converter.converter import (
    _build_dom_grid,
    _collect_embedded_css,
    _determine_sheet_name,
    _get_cell_style,
    _get_color_rgb,
    _infer_value,
    _is_cell_bold,
    _map_border_style,
    _map_horiz,
    _map_underline,
    _map_vert,
    _parse_border,
    _parse_color,
    _parse_css_rules,
    _parse_font_size,
    _sanitize_hyperlink,
    _sanitize_sheet_name,
    convert_file,
    decode_with_fallback,
    detect_file_format,
    identify_and_convert,
    identify_file_type,
)


def test_decode_and_hyperlink_safety_helpers() -> None:
    assert decode_with_fallback("Halo".encode()) == "Halo"
    assert decode_with_fallback(b"caf\xe9") == "café"

    assert _sanitize_hyperlink(" https://example.com/help ") == "https://example.com/help"
    assert _sanitize_hyperlink("http://example.com") == "http://example.com"
    for unsafe in (
        None,
        "",
        "javascript:alert(1)",
        "/relative/path",
        "https://user:secret@example.com",
        "http://[invalid",
        "https://example.com/" + ("a" * 2048),
    ):
        assert _sanitize_hyperlink(unsafe) is None


def test_detection_validates_paths_extensions_and_content(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError):
        detect_file_format(tmp_path / "missing.xls")

    wrong_extension = tmp_path / "sample.txt"
    wrong_extension.write_text("<html><table><tr><td>x</td></tr></table></html>")
    with pytest.raises(ValueError, match="Unsupported file extension"):
        detect_file_format(wrong_extension)

    html_file = tmp_path / "sample.xls"
    html_file.write_text("<html><table><tr><td>x</td></tr></table></html>")
    assert detect_file_format(html_file) == "html"

    unknown_file = tmp_path / "unknown.xls"
    unknown_file.write_bytes(b"not an HTML or BIFF workbook")
    assert detect_file_format(unknown_file) == "binary"


def test_binary_style_mapping_helpers() -> None:
    assert [_map_horiz(value) for value in range(8)] == [
        "general",
        "left",
        "center",
        "right",
        "fill",
        "justify",
        "centerContinuous",
        "distributed",
    ]
    assert _map_horiz(99) == "general"
    assert [_map_vert(value) for value in range(5)] == [
        "top",
        "center",
        "bottom",
        "justify",
        "distributed",
    ]
    assert _map_vert(99) == "bottom"
    assert [_map_underline(value) for value in (0, 1, 2, 33, 34, 99)] == [
        None,
        "single",
        "double",
        "singleAccounting",
        "doubleAccounting",
        "single",
    ]
    assert _get_color_rgb({12: (1, 2, 255)}, 12) == "0102FF"
    assert _get_color_rgb({}, 8) is None
    assert _get_color_rgb({}, 77) is None
    assert _map_border_style(1) == "thin"
    assert _map_border_style(13) == "slantDashDot"
    assert _map_border_style(99) is None


def test_type_inference_and_sheet_names() -> None:
    assert _infer_value("") == ""
    assert _infer_value("=SUM(A1:A2)") == "'=SUM(A1:A2)"
    assert _infer_value("1.") == "1."
    assert _infer_value("00123") == "00123"
    assert _infer_value("9999999999999999") == "9999999999999999"
    assert _infer_value("42") == 42
    assert _infer_value("4.25") == 4.25
    assert _infer_value("2026-08-24") == datetime.datetime(2026, 8, 24)
    assert _infer_value("2026-99-99") == "2026-99-99"
    assert _infer_value("24 Agustus 2026") == "24 Agustus 2026"

    assert _sanitize_sheet_name("  'Bad/Name:*?[]'  ") == "Bad_Name_____"
    assert _sanitize_sheet_name("") == "Sheet1"
    assert len(_sanitize_sheet_name("x" * 100)) == 31
    assert _determine_sheet_name("Daftar Ibu") == "Daftar_Ibu"
    assert _determine_sheet_name(None) == "Sheet1"
    assert _determine_sheet_name(None, html="<x:Name>Custom/Sheet</x:Name>") == "Custom_Sheet"


def test_css_and_dom_grid_helpers() -> None:
    assert _parse_color("#abc") == "AABBCC"
    assert _parse_color("#123456") == "123456"
    assert _parse_color("rgb(1, 2, 255)") == "0102FF"
    assert _parse_color("transparent") is None
    assert _parse_font_size("16px") == 12.0
    assert _parse_font_size("9pt") == 9.0
    assert _parse_font_size("20") == 15.0
    assert _parse_font_size("large") is None
    assert _parse_border("2px solid #102030") == {
        "width": 2.0,
        "style": "solid",
        "color": "102030",
    }

    css = """
    /* synthetic fixture */
    .report .heading { background-color: #abcdef; color: #112233;
      font-weight: 700; font-size: x-small; border-bottom: 2px solid; }
    .heading { border-color: #445566; }
    """
    rules = _parse_css_rules(css)
    assert rules["heading"]["bg_color"] == "ABCDEF"
    assert rules["heading"]["font_color"] == "112233"
    assert rules["heading"]["bold"] is True
    assert rules["heading"]["font_size"] == 10.0
    assert rules["heading"]["border_color"] == "445566"

    soup = BeautifulSoup(
        """<style>.heading { font-weight: bold; }</style><table>
        <tr class="heading"><th rowspan="2" style="font-size: 12pt">Title</th>
        <td colspan="2"><a href="https://example.com">Help</a></td></tr>
        <tr><td class="heading">A</td><td>B</td></tr></table>""",
        "lxml",
    )
    embedded = _collect_embedded_css(soup)
    rows = soup.find_all("tr")
    grid, merges, bold, hyperlinks, styles = _build_dom_grid(
        rows,
        include_links=True,
        css_rules=embedded,
    )
    assert grid[0] == ["Title", "Help https://example.com", ""]
    assert grid[1] == ["", "A", "B"]
    assert merges == [(0, 0, 1, 0), (0, 1, 0, 2)]
    assert (0, 0) in bold and (1, 1) in bold
    assert hyperlinks[(0, 1)] == "https://example.com"
    assert styles[(0, 0)]["font_size"] == 12.0
    assert _is_cell_bold(soup.find("th")) is True
    assert _get_cell_style(soup.find("th"), embedded)["font_size"] == 12.0


def test_known_layout_identification_and_wrapper(tmp_path: Path) -> None:
    headers = ["NIK", "NAMA", "TGL LAHIR", "TGL HPHT", "TGL HPL", "IMT", "LILA"]
    html = "<html><table><tr>" + "".join(f"<th>{value}</th>" for value in headers) + (
        "</tr><tr><td>9999999999999999</td><td>Synthetic</td>"
        "<td>2026-01-01</td><td>2026-02-01</td><td>2026-11-01</td>"
        "<td>20</td><td>25</td></tr></table></html>"
    )
    soup = BeautifulSoup(html, "lxml")
    assert identify_file_type(soup, soup.find("table")) == "Daftar Ibu Hamil"

    source = tmp_path / "known.xls"
    source.write_text(html, encoding="utf-8")
    output, file_type = identify_and_convert(source)
    assert output.exists()
    assert output.name == "known_converted.xlsx"
    assert file_type == "Daftar Ibu Hamil"


def test_public_api_rejects_invalid_output_locations(tmp_path: Path) -> None:
    source = tmp_path / "source.xls"
    source.write_text("<html><table><tr><td>x</td></tr></table></html>")

    with pytest.raises(ValueError, match="Output directory does not exist"):
        convert_file(source, tmp_path / "missing" / "output.xlsx")
    with pytest.raises(FileNotFoundError):
        convert_file(tmp_path / "missing.xls")
