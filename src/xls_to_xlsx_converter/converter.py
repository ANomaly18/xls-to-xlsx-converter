"""XLS to XLSX Converter - Core engine.

Preserves formatting, merged cells, column widths, row heights,
number formats, fonts, alignment, borders, and fills.
Handles both binary .xls (BIFF) and HTML-disguised .xls.
"""

from __future__ import annotations

import datetime
import logging
import math
import os
import re
import tempfile
from pathlib import Path
from typing import Callable

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    NamedStyle,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
import xlrd
from xlrd import XL_CELL_DATE, XL_CELL_EMPTY, XL_CELL_ERROR, XL_CELL_NUMBER
from xlrd import XL_CELL_TEXT, XL_CELL_BOOLEAN, XL_CELL_BLANK

logger = logging.getLogger(__name__)
if not logger.handlers:
    logger.addHandler(logging.NullHandler())


def decode_with_fallback(raw: bytes) -> str:
    """Decode bytes with encoding fallback chain: utf-8 → windows-1252 → latin-1."""
    for enc in ("utf-8", "windows-1252", "latin-1"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return raw.decode("latin-1", errors="replace")


def _sanitize_hyperlink(value: object) -> str | None:
    """Return a safe HTTP(S) hyperlink or ``None`` for unsafe targets."""
    if not isinstance(value, str):
        return None
    candidate = value.strip()
    if not candidate or len(candidate) > 2048:
        return None

    from urllib.parse import urlsplit

    try:
        parsed = urlsplit(candidate)
    except ValueError:
        return None
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        return None
    if parsed.username or parsed.password:
        return None
    return candidate


# Maximum file size: 50 MB. Prevents DoS from massive files.
MAX_FILE_SIZE = 50 * 1024 * 1024
MAX_HTML_ROWS = 100_000
MAX_HTML_COLUMNS = 2_000
MAX_HTML_EXPANDED_CELLS = 2_000_000
MAX_CELL_SPAN = 10_000

# ---------------------------------------------------------------------------
# Detection helpers
# ---------------------------------------------------------------------------

_XLS_MAGIC_BYTES = {
    b"\xD0\xCF\x11\xE0",  # BIFF (OLE2) - xls, doc, ppt
    b"\x09\x08\x10\x00\x00\x06\x05\x00",  # BIFF3
    b"\x09\x04\x06\x00\x00\x00\x10\x00",  # BIFF4
}

HTML_SIGNATURES = [
    b"<html",
    b"<table",
    b"<tr",
    b"<td",
    b"office:excel",
    b"xmlns:x=",
    b"<head",
    b"<body",
]


def detect_file_format(path: str | Path) -> str:
    """Return 'binary' if real BIFF .xls, 'html' if HTML-disguised .xls."""
    path = Path(path)
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    if path.suffix.lower() not in (".xls",):
        raise ValueError(f"Unsupported file extension: {path.suffix}")

    # File size check — prevent DoS from massive files
    # Resolve symlinks to prevent bypass via symlink to large file
    real_path = os.path.realpath(path)
    file_size = os.path.getsize(real_path)
    if file_size > MAX_FILE_SIZE:
        raise ValueError(
            f"File terlalu besar ({file_size / 1024 / 1024:.1f} MB). "
            f"Maksimal {MAX_FILE_SIZE / 1024 / 1024:.0f} MB."
        )

    try:
        with open(path, "rb") as f:
            header = f.read(12)

        # Quick BIFF magic check
        for magic in _XLS_MAGIC_BYTES:
            if header.startswith(magic):
                return "binary"

        # Read more for HTML detection
        with open(path, "rb") as f:
            chunk = f.read(8192)

        # Decode as latin-1 (single-byte, never fails)
        text = chunk.decode("latin-1", errors="replace").lower()

        # Strong HTML signals
        html_score = 0
        for sig in HTML_SIGNATURES:
            if sig.decode("latin-1") in text:
                html_score += 1

        if html_score >= 2:
            return "html"

        # If we see table elements with data-like patterns, it's likely HTML
        if "<tr" in text and ("<td" in text or "<th" in text):
            return "html"

        # Otherwise default to binary and let xlrd handle it
        return "binary"

    except PermissionError:
        raise
    except Exception as exc:
        raise ValueError(f"Cannot read file: {exc}")


# ---------------------------------------------------------------------------
# File type identification (7 known Sigizi/Sihepi types)
# ---------------------------------------------------------------------------

# Each type has a unique signature: set of column headers that MUST be present.
# These are checked against the FIRST header row of the main table.
_FILE_TYPE_SIGNATURES = {
    "Daftar Ibu Hamil": {
        "required_columns": {"NIK", "NAMA", "TGL LAHIR", "TGL HPHT", "TGL HPL", "IMT", "LILA"},
        "source": "Sigizi",
    },
    "Daftar Ibu Nifas": {
        "required_columns": {"NIK", "NAMA", "TGL LAHIR", "TGL HPHT", "TGL ABORTUS"},
        "source": "Sigizi",
    },
    "Daftar Ibu": {
        "required_columns": {"NIK", "NAMA", "TGL LAHIR", "TGL HPHT", "TGL HPL", "NAMA SUAMI"},
        "excludes": {"IMT", "LILA"},  # Distinguish from Daftar Ibu Hamil
        "source": "Sigizi",
    },
    "Kohort Ibu": {
        "required_columns": {"NIK", "NAMA", "NAMA SUAMI", "PUKESMAS", "ALAMAT"},
        "source": "Sigizi",
    },
    "Pemeriksaan ANC": {
        "required_columns": {"NAMA", "NIK", "TANGGAL LAHIR", "UMUR"},
        "source": "Sigizi",
    },
    "Posyandu BNBA": {
        "required_columns": {"NIK", "NAMA", "JK", "TGL LAHIR", "BB LAHIR", "TB LAHIR"},
        "source": "Sigizi",
    },
    "Rekap 3e": {
        "required_columns": {"TANGGAL SCREENING", "NIK", "NAMA"},
        "multi_row_check": True,  # Check first 4 rows for signature
        "source": "Sihepi",
        "pre_table_keyword": "FORMULIR REGISTRASI",
    },
}

KNOWN_FILE_TYPES = list(_FILE_TYPE_SIGNATURES.keys())


def _clean_text(text: str) -> str:
    """Clean extracted text: normalize whitespace but preserve \\xa0 for expected match."""
    # Replace multiple spaces/newlines/tabs with single space
    text = re.sub(r'[\s]+', ' ', text)
    return text.strip()


def _extract_text_with_spacing(element) -> str:
    """Extract text from an element preserving spacing between children."""
    # Use separator=' ' to preserve spaces between elements
    parts = []
    for child in element.descendants:
        if hasattr(child, 'string') and child.string is not None:
            t = str(child.string)
            # Replace \xa0 with space
            t = t.replace("\xa0", " ")
            parts.append(t)
    if not parts:
        # Fallback: get_text with separator
        text = element.get_text(separator=" ")
        text = text.replace("\xa0", " ")
        return text.strip()
    return " ".join(parts).strip()


def _get_cell_text(cell_element) -> str:
    """Extract text from a td/th element preserving spacing and \\xa0."""
    # Use separator=' ' to preserve spaces between elements
    text = cell_element.get_text(separator=" ")
    # Normalize whitespace: collapse regular spaces/tabs/newlines but preserve \xa0
    text = re.sub(r'[ \t\n\r]+', ' ', text)
    # Strip leading/trailing regular spaces but NOT \xa0
    text = text.strip(' \t\n\r')
    return text


def identify_file_type(soup, main_table) -> str | None:
    """Identify which of the 7 known file types this HTML represents.

    Checks column headers against unique signatures for each type.
    Returns the type name or None if unrecognized.
    """
    # Collect header text from first 4 rows (some files have multi-row headers)
    header_text = set()
    tr = main_table.find("tr")
    rows_checked = 0
    while tr and rows_checked < 4:
        for cell in tr.find_all(["td", "th"]):
            t = _get_cell_text(cell).upper()
            if t:
                header_text.add(t)
        tr = tr.find_next_sibling("tr")
        rows_checked += 1

    # Try to match against each signature
    for type_name, sig in _FILE_TYPE_SIGNATURES.items():
        required = sig["required_columns"]
        excludes = sig.get("excludes", set())

        # Check if all required columns are present
        if not required.issubset(header_text):
            continue

        # Check if any excluded columns are present
        if excludes and excludes.intersection(header_text):
            continue

        # Additional check for Rekap 3e: pre-table keyword
        if sig.get("pre_table_keyword"):
            page_text = soup.get_text(" ", strip=True)[:2000].upper()
            if sig["pre_table_keyword"].upper() not in page_text:
                continue

        return type_name

    return None


# ---------------------------------------------------------------------------
# Binary .xls conversion (xlrd + openpyxl)
# ---------------------------------------------------------------------------


def _map_horiz(xlrd_align: int) -> str:
    """Map xlrd horizontal alignment to openpyxl."""
    mapping = {
        0: "general",
        1: "left",
        2: "center",
        3: "right",
        4: "fill",
        5: "justify",
        6: "centerContinuous",
        7: "distributed",
    }
    return mapping.get(xlrd_align, "general")


def _map_vert(xlrd_align: int) -> str:
    """Map xlrd vertical alignment to openpyxl."""
    mapping = {
        0: "top",
        1: "center",
        2: "bottom",
        3: "justify",
        4: "distributed",
    }
    return mapping.get(xlrd_align, "bottom")


def _map_underline(xlrd_underline: int) -> str | None:
    """Map xlrd underline type to openpyxl."""
    if xlrd_underline == 0:
        return None
    if xlrd_underline == 1:
        return "single"
    if xlrd_underline == 2:
        return "double"
    if xlrd_underline == 33:
        return "singleAccounting"
    if xlrd_underline == 34:
        return "doubleAccounting"
    return "single"


def _get_color_rgb(colour_map: dict, colour_idx: int) -> str | None:
    """Convert xlrd colour index to 'RRGGBB' string or None."""
    if colour_idx is None or colour_idx == 0:
        return None
    # Special indices
    if colour_idx in (8, 9, 10, 32767):
        return None  # default colours
    rgb = colour_map.get(colour_idx)
    if rgb and len(rgb) == 3:
        return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    return None


def _map_border_style(xlrd_style: int) -> str | None:
    """Map xlrd border style to openpyxl border style string."""
    mapping = {
        0: None,
        1: "thin",
        2: "medium",
        3: "dashed",
        4: "dotted",
        5: "thick",
        6: "double",
        7: "hair",
        8: "mediumDashed",
        9: "dashDot",
        10: "mediumDashDot",
        11: "dashDotDot",
        12: "mediumDashDotDot",
        13: "slantDashDot",
    }
    return mapping.get(xlrd_style, None)


def _convert_binary_xls(input_path: Path, output_path: Path, progress: Callable) -> None:
    """Convert binary (BIFF) .xls to .xlsx using xlrd + openpyxl."""
    progress("Membaca file .xls (binary)...")
    xl_book = xlrd.open_workbook(str(input_path), formatting_info=True)
    colormap = getattr(xl_book, "colour_map", {})
    total_sheets = xl_book.nsheets

    progress(f"Ditemukan {total_sheets} sheet. Menulis ke .xlsx...")

    wb = Workbook()
    # Remove default sheet - we'll add proper ones
    wb.remove(wb.active)

    for sheet_idx in range(total_sheets):
        xl_sheet = xl_book.sheet_by_index(sheet_idx)
        ws = wb.create_sheet(title=xl_sheet.name or f"Sheet{sheet_idx + 1}")

        nrows = xl_sheet.nrows
        ncols = xl_sheet.ncols
        progress(f"  Sheet {sheet_idx + 1}/{total_sheets}: '{xl_sheet.name}' ({nrows} baris x {ncols} kolom)")

        # --- Copy merged cells ---
        for merge in xl_sheet.merged_cells:
            # xlrd merged_cells returns (rlo, rhi, clo, chi)
            # rlo/rhi = row low/high (rhi exclusive), clo/chi = col low/high (chi exclusive)
            # openpyxl uses 1-based inclusive ranges
            rlo, rhi, clo, chi = merge
            ws.merge_cells(
                start_row=rlo + 1,
                start_column=clo + 1,
                end_row=rhi,  # xlrd exclusive = openpyxl inclusive (already one past)
                end_column=chi,  # xlrd exclusive = openpyxl inclusive (already one past)
            )

        # --- Copy column widths ---
        for col_idx in range(ncols):
            colinfo = xl_sheet.colinfo_map.get(col_idx)
            if colinfo and colinfo.width:
                # xlrd width is in 1/256 of character width
                w = colinfo.width / 256
                if w > 0 and w < 255:
                    ws.column_dimensions[get_column_letter(col_idx + 1)].width = w

        # --- Copy row heights ---
        for row_idx in range(nrows):
            rowinfo = xl_sheet.rowinfo_map.get(row_idx)
            if rowinfo and rowinfo.height:
                # xlrd height is in 1/20 of a point
                h = rowinfo.height / 20
                if h > 0:
                    ws.row_dimensions[row_idx + 1].height = h

        # --- Copy cells with data and formatting ---
        for row_idx in range(nrows):
            for col_idx in range(ncols):
                cell = xl_sheet.cell(row_idx, col_idx)

                # Skip completely empty cells (but not blank-looking ones with formatting)
                if cell.ctype == 0:  # XL_CELL_EMPTY
                    continue

                # Map value
                value = _map_xlrd_cell_value(cell, xl_book)
                if value is None and cell.ctype != 6:  # skip if error or blank
                    continue

                openpyxl_cell = ws.cell(row=row_idx + 1, column=col_idx + 1, value=value)

                # --- Apply formatting from XF record ---
                try:
                    xf_index = cell.xf_index
                    if xf_index is not None and xf_index < len(xl_book.xf_list):
                        xf = xl_book.xf_list[xf_index]
                        _apply_xf_format(openpyxl_cell, xf, xl_book, colormap)
                except Exception:
                    pass  # best-effort formatting

    progress("Menyimpan file .xlsx...")
    wb.save(str(output_path))
    progress("Konversi binary .xls selesai.")


def _get_format_str(xl_book, fmt_key) -> str:
    """Get format string from xlrd format_map, handling Format objects."""
    if fmt_key is None:
        return ""
    fmt = xl_book.format_map.get(fmt_key)
    if fmt is None:
        return ""
    # xlrd 1.2.0 format_map values are Format objects with .format_str
    if hasattr(fmt, "format_str"):
        return fmt.format_str or ""
    if isinstance(fmt, str):
        return fmt
    return str(fmt)


def _map_xlrd_cell_value(cell, xl_book):
    """Map xlrd cell to Python value."""
    if cell.ctype == XL_CELL_EMPTY:
        return None
    if cell.ctype == XL_CELL_BLANK:
        return None
    if cell.ctype == XL_CELL_ERROR:
        return None
    if cell.ctype == XL_CELL_TEXT:
        return cell.value
    if cell.ctype == XL_CELL_NUMBER:
        # Check if it's actually a date stored as number
        if cell.xf_index is not None and cell.xf_index < len(xl_book.xf_list):
            fmt_key = xl_book.xf_list[cell.xf_index].format_key
            fmt_str = _get_format_str(xl_book, fmt_key)
            if _is_date_format(fmt_str):
                try:
                    dt = xlrd.xldate_as_datetime(cell.value, xl_book.datemode)
                    return dt
                except Exception:
                    pass
        return cell.value
    if cell.ctype == XL_CELL_DATE:
        try:
            return xlrd.xldate_as_datetime(cell.value, xl_book.datemode)
        except Exception:
            return cell.value
    if cell.ctype == XL_CELL_BOOLEAN:
        return bool(cell.value)
    return cell.value


_DATE_FORMAT_PATTERNS = re.compile(
    r"(?:^|[^a-z])"
    r"(?:d{1,4}|m{1,5}|y{2,4}|yy|hh|mm|ss|AM/PM|"
    r"mmmm|mmm|dddd|ddd)"
    r"(?:$|[^a-z])",
    re.IGNORECASE,
)


def _is_date_format(fmt_str: str) -> bool:
    """Heuristic check if a number format string is a date format."""
    if not fmt_str:
        return False
    # General/number formats
    if fmt_str.strip() in ("General", "0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%"):
        return False
    # Standard date format codes
    if _DATE_FORMAT_PATTERNS.search(fmt_str):
        return True
    return False


def _apply_xf_format(cell, xf, xl_book, colormap) -> None:
    """Apply xlrd XF formatting to an openpyxl cell (best-effort)."""

    # --- Number format ---
    fmt_key = xf.format_key
    fmt_str = _get_format_str(xl_book, fmt_key)
    if fmt_str and fmt_str not in ("General", ""):
        cell.number_format = fmt_str
        # Special handling for '@' (text format): force value to clean string
        if fmt_str == "@" and cell.value is not None:
            try:
                raw = cell.value
                if isinstance(raw, float):
                    if not (math.isnan(raw) or math.isinf(raw)):
                        s = str(int(raw)) if raw == int(raw) else str(raw)
                        cell.value = s
                elif isinstance(raw, int):
                    cell.value = str(raw)
                else:
                    cell.value = str(raw)
            except Exception:
                pass

    # --- Font ---
    try:
        font_obj = xl_book.font_list[xf.font_index]
        font_kw = {}

        if font_obj.name:
            font_kw["name"] = font_obj.name
        # xlrd height is in twentieths of a point; convert to points
        if font_obj.height:
            font_kw["size"] = font_obj.height / 20
        # xlrd bold flag (1=True) OR weight >= 700
        if font_obj.bold or getattr(font_obj, "weight", 400) >= 700:
            font_kw["bold"] = True
        if font_obj.italic:
            font_kw["italic"] = True
        if font_obj.struck_out:
            font_kw["strike"] = True

        ul = _map_underline(font_obj.underline_type)
        if ul:
            font_kw["underline"] = ul

        color_rgb = _get_color_rgb(colormap, font_obj.colour_index)
        if color_rgb:
            font_kw["color"] = color_rgb

        cell.font = Font(**font_kw)
    except Exception:
        pass

    # --- Alignment ---
    try:
        align = xf.alignment
        align_kw = {}
        if align is not None:
            hor = _map_horiz(align.hor_align)
            if hor != "general":
                align_kw["horizontal"] = hor
            ver = _map_vert(align.vert_align)
            if ver != "bottom":
                align_kw["vertical"] = ver
            if align.text_wrapped:
                align_kw["wrap_text"] = True
            if align_kw:
                cell.alignment = Alignment(**align_kw)
    except Exception:
        pass

    # --- Border ---
    try:
        border = xf.border
        if border is not None:
            sides = {}
            # xlrd border uses *_line_style and *_colour_index attributes
            border_map = [
                ("left", "left_line_style", "left_colour_index"),
                ("right", "right_line_style", "right_colour_index"),
                ("top", "top_line_style", "top_colour_index"),
                ("bottom", "bottom_line_style", "bottom_colour_index"),
            ]
            for side_name, style_attr, color_attr in border_map:
                style_val = getattr(border, style_attr, 0)
                mapped_style = _map_border_style(style_val)
                if mapped_style:
                    side_kw = {"style": mapped_style}
                    color_idx = getattr(border, color_attr, None)
                    side_rgb = _get_color_rgb(colormap, color_idx)
                    if side_rgb:
                        side_kw["color"] = side_rgb
                    sides[side_name] = Side(**side_kw)

            if sides:
                cell.border = Border(**sides)
    except Exception:
        pass

    # --- Fill (background) ---
    try:
        bg = xf.background
        if bg is not None:
            fill_pattern = getattr(bg, "fill_pattern", 0)
            if fill_pattern:  # non-zero = some pattern
                fill_rgb = _get_color_rgb(colormap, bg.pattern_colour_index)
                if fill_rgb:
                    cell.fill = PatternFill(
                        start_color=fill_rgb, end_color=fill_rgb, fill_type="solid"
                    )
    except Exception:
        pass


# ---------------------------------------------------------------------------
# HTML-disguised .xls conversion (BeautifulSoup + openpyxl)
# ---------------------------------------------------------------------------

# Sentinel for rowspan-occupied cells
_OCCUPIED = object()

# Keywords used to identify the main data table in HTML .xls files
_TABLE_SCORE_KEYWORDS = {
    "NIK": 20,
    "NAMA": 15,
    "TANGGAL LAHIR": 20,
    "TGL LAHIR": 18,
    "UMUR": 10,
    "PUSKESMAS": 10,
    "POSYANDU": 10,
    "HPHT": 10,
    "PEMERIKSAAN": 8,
    "ALAMAT": 8,
    "DESA": 5,
    "KECAMATAN": 5,
    "IBU": 5,
    "ANAK": 5,
    "NO.": 4,
    "NO": 3,
}


def _score_table(table) -> int:
    """Score a BeautifulSoup <table> by content relevance."""
    text = table.get_text(separator=" ", strip=True).upper()
    score = 0
    for keyword, points in _TABLE_SCORE_KEYWORDS.items():
        if keyword in text:
            score += points
    # Bonus for row count
    rows = table.find_all("tr")
    score += min(len(rows) * 2, 40)
    return score


def _infer_value(text: str) -> str | int | float | datetime.datetime:
    """Try to infer the proper Python type from a string value.

    Returns datetime for ISO date strings (YYYY-MM-DD), int/float for numbers,
    or the original string. Preserves strings that should stay as strings:
    - "1.", "2." (numbering with trailing dot)
    - "0804234253" (leading zeros)
    - 16-digit identifier-like numbers such as "9999999999999999"
    - Indonesian date format "DD MMMM YYYY" (e.g., "02 Januari 2026")
    """
    if not text:
        return ""

    # openpyxl interprets strings beginning with '=' as formulas. Prefix an
    # apostrophe so HTML-sourced content remains inert spreadsheet text.
    if text.startswith("="):
        return "'" + text

    # Don't convert if it ends with "." (like "1.", "2." in the "No" column)
    if text.endswith("."):
        return text

    # Do not convert digit strings with leading zeros (identifier-safe).
    if text.startswith("0") and text.replace(" ", "").isdigit():
        return text

    # Keep 16-digit identifier-like numbers as strings.
    if text.isdigit() and len(text) == 16:
        return text

    # Try ISO date: YYYY-MM-DD only (NOT zero dates like 0000-00-00)
    date_match = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", text)
    if date_match:
        year, month, day = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
        if year > 0 and month > 0 and day > 0:
            try:
                return datetime.datetime(year, month, day)
            except (ValueError, OverflowError):
                pass

    # NOTE: Indonesian date format "DD MMMM YYYY" is NOT converted to datetime.
    # These should remain as strings to match the expected output format.

    # Try integer (only if no leading zeros preserved above)
    try:
        val = int(text)
        # Don't convert if original had leading zeros
        if text != str(val):
            return text
        return val
    except (ValueError, OverflowError):
        pass

    # Try float
    try:
        return float(text)
    except (ValueError, OverflowError):
        pass

    return text


def _parse_color(color_str):
    """Parse CSS color to RRGGBB format for openpyxl."""
    if not color_str:
        return None
    color_str = color_str.strip().lower()
    # Hex: #rgb or #rrggbb
    if color_str.startswith("#"):
        hex_val = color_str[1:]
        if len(hex_val) == 3:
            return f"{hex_val[0]*2}{hex_val[1]*2}{hex_val[2]*2}".upper()
        elif len(hex_val) == 6:
            return hex_val.upper()
    # rgb(r, g, b)
    import re
    rgb_match = re.match(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", color_str)
    if rgb_match:
        r, g, b = int(rgb_match.group(1)), int(rgb_match.group(2)), int(rgb_match.group(3))
        return f"{r:02X}{g:02X}{b:02X}"
    return None


def _parse_font_size(size_str):
    """Parse CSS font-size to points for openpyxl."""
    if not size_str:
        return None
    size_str = size_str.strip().lower()
    import re
    # px value
    px_match = re.match(r"([\d.]+)\s*px", size_str)
    if px_match:
        return float(px_match.group(1)) * 0.75  # px to pt approx
    # pt value
    pt_match = re.match(r"([\d.]+)\s*pt", size_str)
    if pt_match:
        return float(pt_match.group(1))
    # numeric (assume px)
    num_match = re.match(r"([\d.]+)$", size_str)
    if num_match:
        return float(num_match.group(1)) * 0.75
    return None


def _parse_border(style_str):
    """Parse CSS border properties from inline style."""
    if not style_str:
        return None
    style_lower = style_str.lower()
    import re
    # border: 1px solid #color (shorthand)
    border_match = re.match(r"([\d.]+)\s*(px|pt)?\s*(solid|dashed|dotted|double)\s*(#\w+|rgb\([^)]+\))", style_lower)
    if border_match:
        width = float(border_match.group(1))
        style = border_match.group(3)
        color = _parse_color(border_match.group(4))
        return {"width": width, "style": style, "color": color}
    return None


def _get_cell_style(cell_element, css_rules: dict | None = None):
    """Extract inline CSS style from HTML cell, plus external CSS class rules."""
    style_str = cell_element.get("style", "")
    if not style_str:
        # Also check parent tr for inherited styles
        parent_tr = cell_element.find_parent("tr")
        if parent_tr:
            style_str = parent_tr.get("style", "")
    if not style_str:
        style_str = ""

    style_lower = style_str.lower()
    result = {}

    import re
    # Background color (inline)
    bg_match = re.search(r"(?:background-color|background)\s*:\s*([^;]+)", style_lower)
    if bg_match:
        color = _parse_color(bg_match.group(1).strip())
        if color:
            result["bg_color"] = color

    # Font size (inline)
    fs_match = re.search(r"font-size\s*:\s*([^;]+)", style_lower)
    if fs_match:
        size = _parse_font_size(fs_match.group(1).strip())
        if size:
            result["font_size"] = size

    # Font color (inline)
    fc_match = re.search(r"(?:^|;)\s*color\s*:\s*([^;]+)", style_lower)
    if fc_match:
        fc_str = fc_match.group(1).strip()
        color = _parse_color(fc_str)
        if color:
            result["font_color"] = color

    # Border (inline)
    border_match = re.search(r"border\s*:\s*([^;]+)", style_lower)
    if border_match:
        border = _parse_border(border_match.group(1).strip())
        if border:
            result["border"] = border

    # Also check CSS class rules
    if css_rules:
        # Check cell's own class
        cell_classes = cell_element.get("class", [])
        if isinstance(cell_classes, str):
            cell_classes = cell_classes.split()
        for cls in cell_classes:
            if cls in css_rules:
                rule = css_rules[cls]
                for key, val in rule.items():
                    if key not in result:  # inline takes priority
                        result[key] = val

        # Check parent tr class
        parent_tr = cell_element.find_parent("tr")
        if parent_tr:
            tr_classes = parent_tr.get("class", [])
            if isinstance(tr_classes, str):
                tr_classes = tr_classes.split()
            for cls in tr_classes:
                if cls in css_rules:
                    rule = css_rules[cls]
                    for key, val in rule.items():
                        if key not in result:
                            result[key] = val

    return result if result else None


def _is_cell_bold(cell_element) -> bool:
    """Detect if an HTML cell should be bold based on tag, class, or inline style."""
    # <th> tags are bold by default
    if cell_element.name == "th":
        return True
    # Check CSS class for "head" (e.g., "head", "head-one", "head-two")
    css_class = cell_element.get("class", [])
    if isinstance(css_class, list):
        css_class = " ".join(css_class)
    if css_class and "head" in css_class.lower():
        return True
    # Check inline style for font-weight:bold or font-weight: 700+
    style = cell_element.get("style", "")
    if style:
        style_lower = style.lower()
        if "font-weight" in style_lower:
            if "bold" in style_lower or "700" in style_lower or "800" in style_lower or "900" in style_lower:
                return True
    # Check if parent <tr> has content_header class (Sihepi)
    parent_tr = cell_element.find_parent("tr")
    if parent_tr:
        tr_class = parent_tr.get("class", [])
        if isinstance(tr_class, list):
            tr_class = " ".join(tr_class)
        if tr_class and "content_header" in tr_class.lower():
            return True
    # Check if cell contains <b> or <strong> as direct child
    if cell_element.find(["b", "strong"]):
        return True
    return False


def _build_dom_grid(
    rows, include_links: bool = False, css_rules: dict | None = None
) -> tuple[list[list[str]], list[tuple[int, int, int, int]], set[tuple[int, int]], dict[tuple[int, int], str], dict[tuple[int, int], dict]]:
    """Build a 2D grid from HTML table rows, handling rowspan/colspan.

    Important: rowspan text is written ONLY in the source row (r=0).
    Spanned cells are marked as occupied to push later cells right,
    but do NOT duplicate the text.

    Uses explicit TR index (not len(grid)) so that rowspan-precreated
    grid rows are correctly reused by subsequent <tr> elements.

    Args:
        rows: BeautifulSoup <tr> elements
        include_links: If True, extract href URLs from <a> tags for cell text

    Returns:
        Tuple of (grid, merge_ranges, bold_cells, hyperlinks, cell_styles) where:
        - grid: 2D list of cell values
        - merge_ranges: list of (start_row, start_col, end_row, end_col)
        - bold_cells: set of (row, col) tuples that should be bold
        - hyperlinks: dict of (row, col) -> URL
        - cell_styles: dict of (row, col) -> {bg_color, font_size, font_color, border}
        All indices are 0-based.
    """
    grid: list[list[object]] = []
    merge_ranges: list[tuple[int, int, int, int]] = []
    bold_cells: set[tuple[int, int]] = set()
    hyperlinks: dict[tuple[int, int], str] = {}
    cell_styles: dict[tuple[int, int], dict] = {}
    expanded_cells = 0

    if len(rows) > MAX_HTML_ROWS:
        raise ValueError(
            f"HTML table exceeds the {MAX_HTML_ROWS:,}-row safety limit."
        )

    for tr_idx, tr in enumerate(rows):
        # Ensure grid has enough rows (may already exist from rowspan)
        while len(grid) <= tr_idx:
            grid.append([])
        row_idx = tr_idx

        cells = tr.find_all(["td", "th"])
        col_idx = 0
        for cell in cells:
            # Skip occupied columns
            while col_idx < len(grid[row_idx]) and grid[row_idx][col_idx] is not None:
                col_idx += 1

            # Extract text with spacing preserved (keep \xa0)
            text = _get_cell_text(cell)

            # Extract URL from <a> tag if present
            a_tag = cell.find("a")
            href = _sanitize_hyperlink(a_tag.get("href")) if a_tag else None
            if href:
                # For include_links mode (Pemeriksaan ANC), append URL to text
                if include_links and href and href not in text:
                    text = f"{text} {href}".strip()
                # Always store hyperlink separately for openpyxl
                # Store at source cell position
                hyperlinks[(row_idx, col_idx)] = href

            # Detect bold styling
            if _is_cell_bold(cell):
                bold_cells.add((row_idx, col_idx))

            # Extract inline styles
            style = _get_cell_style(cell, css_rules=css_rules)
            if style:
                cell_styles[(row_idx, col_idx)] = style

            try:
                rowspan = int(cell.get("rowspan", 1))
                colspan = int(cell.get("colspan", 1))
            except (TypeError, ValueError) as exc:
                raise ValueError("Invalid HTML rowspan or colspan value.") from exc
            if not 1 <= rowspan <= MAX_CELL_SPAN or not 1 <= colspan <= MAX_CELL_SPAN:
                raise ValueError("HTML rowspan or colspan exceeds the safety limit.")
            if col_idx + colspan > MAX_HTML_COLUMNS:
                raise ValueError(
                    f"HTML table exceeds the {MAX_HTML_COLUMNS:,}-column safety limit."
                )
            expanded_cells += rowspan * colspan
            if expanded_cells > MAX_HTML_EXPANDED_CELLS:
                raise ValueError(
                    "HTML table exceeds the expanded-cell safety limit."
                )

            # Track merge range if rowspan > 1 OR colspan > 1
            if rowspan > 1 or colspan > 1:
                merge_ranges.append((row_idx, col_idx, row_idx + rowspan - 1, col_idx + colspan - 1))

            for r in range(rowspan):
                while len(grid) <= row_idx + r:
                    grid.append([])
                for c in range(colspan):
                    while len(grid[row_idx + r]) <= col_idx + c:
                        grid[row_idx + r].append(None)
                    # Write text only in source row AND first column of span;
                    # mark all other spanned cells as occupied so they remain empty.
                    grid[row_idx + r][col_idx + c] = text if (r == 0 and c == 0) else _OCCUPIED
                    # Propagate bold, hyperlink, and styles to all merged cells
                    if (row_idx, col_idx) in bold_cells:
                        bold_cells.add((row_idx + r, col_idx + c))
                    if (row_idx, col_idx) in hyperlinks:
                        hyperlinks[(row_idx + r, col_idx + c)] = hyperlinks[(row_idx, col_idx)]
                    if (row_idx, col_idx) in cell_styles:
                        cell_styles[(row_idx + r, col_idx + c)] = cell_styles[(row_idx, col_idx)]

            col_idx += colspan

    # Replace None and _OCCUPIED with empty string for output
    clean_grid = [["" if v is None or v is _OCCUPIED else str(v) for v in row] for row in grid]
    return clean_grid, merge_ranges, bold_cells, hyperlinks, cell_styles


def _collect_pre_table_content(soup, main_table, file_type: str | None) -> list[list[str]]:
    """Collect content rows from elements before the main data table.

    Different file types have different pre-table content:
    - Rekap 3e: H1 title, BR tags → empty rows, Table 0 (Tahun/Bulan)
    - Pemeriksaan ANC, Posyandu BNBA: <caption> inside table
    - Others: div with column names (should be IGNORED - they duplicate the table header)
    """
    pre_rows: list[list[str]] = []

    if file_type == "Rekap 3e":
        # H1 title
        h1 = soup.find(["h1", "h2"])
        if h1 and main_table not in h1.parents and not h1.find_parent("table"):
            text = _clean_text(h1.get_text(" "))
            if text:
                pre_rows.append([text])
                # Add 1 empty row after H1
                pre_rows.append([""])

        # Table 0 (Tahun/Bulan) and BR tags between H1 and main table
        body = soup.find("body")
        if body:
            found_h1 = False
            br_between = 0

            for child in body.children:
                if not hasattr(child, 'name'):
                    continue

                if child.name in ('h1', 'h2'):
                    found_h1 = True
                    continue

                if child.name == 'br' and found_h1:
                    br_between += 1
                    continue

                if child.name == 'table' and found_h1 and child is not main_table:
                    # This is Table 0 (Tahun/Bulan)
                    for tr in child.find_all("tr"):
                        cells = [_clean_text(td.get_text(" ")) for td in tr.find_all(["td", "th"])]
                        if any(cells):
                            pre_rows.append(cells)
                    continue

                if child is main_table:
                    break

            # Add empty rows for BR tags between Table 0 and main table
            for _ in range(br_between):
                pre_rows.append([""])

        # Add 1 extra empty row before the main table header
        # (Expected output has 3 empty rows between Table 0 and main table)
        pre_rows.append([""])

    elif file_type in ("Pemeriksaan ANC", "Posyandu BNBA"):
        # Extract from <caption> inside the main table
        caption = main_table.find("caption")
        if caption:
            text = _clean_text(caption.get_text(" "))
            if text:
                pre_rows.append([text])

    # For other types: skip pre-table div content (it's just a duplicate of headers)

    return pre_rows


def _strip_trailing_empty(grid: list[list[str]]) -> list[list[str]]:
    """Remove trailing rows that are completely empty."""
    while grid and all(v == "" for v in grid[-1]):
        grid.pop()
    return grid


def _parse_css_rules(css_text: str) -> dict:
    """Parse CSS text into class rules dict.

    Returns: {class_name: {bg_color, font_color, font_size, bold, border}}
    """
    import re
    rules = {}

    # Remove comments
    css_text = re.sub(r'/\*.*?\*/', '', css_text, flags=re.DOTALL)

    # Parse each rule block — match selector (any chars before {)
    pattern = r'([.#\w][\w\s.#:>+-]*?)\s*\{([^}]+)\}'
    for match in re.finditer(pattern, css_text):
        selector = match.group(1).strip()
        body = match.group(2).strip()

        # Extract primary class name from selector
        # e.g. ".content_table .content_header td" → "content_header"
        # e.g. ".content_table td" → "content_table"
        parts = selector.split()
        # Find the last class (starts with .)
        class_name = None
        for part in reversed(parts):
            if part.startswith('.'):
                class_name = part.lstrip('.')
                break
        if not class_name and parts:
            class_name = parts[0].lstrip('.')
        if not class_name:
            continue

        props = {}

        # Background color
        bg_match = re.search(r'background-color\s*:\s*(#[0-9A-Fa-f]+)', body)
        if bg_match:
            props['bg_color'] = _parse_color(bg_match.group(1))

        # Font color
        fc_match = re.search(r'(?:^|;)\s*color\s*:\s*(#[0-9A-Fa-f]+)', body)
        if fc_match:
            fc_val = fc_match.group(1)
            # Skip if it's the same as background
            if 'bg_color' not in props or props.get('bg_color') != _parse_color(fc_val):
                props['font_color'] = _parse_color(fc_val)

        # Font weight
        if 'bold' in body.lower() or '700' in body:
            props['bold'] = True

        # Font size
        fs_match = re.search(r'font-size\s*:\s*([^;]+)', body)
        if fs_match:
            size_str = fs_match.group(1).strip()
            if size_str == 'xx-small':
                props['font_size'] = 8.0  # Excel default small
            elif size_str == 'x-small':
                props['font_size'] = 10.0
            elif size_str == 'small':
                props['font_size'] = 12.0
            else:
                props['font_size'] = _parse_font_size(size_str)

        # Border
        border_color_match = re.search(r'border-color\s*:\s*(#[0-9A-Fa-f]+)', body)
        if border_color_match:
            props['border_color'] = _parse_color(border_color_match.group(1))

        # Border width from individual border properties
        for border_prop in ['border-bottom', 'border-right', 'border-left', 'border-top']:
            bp_match = re.search(rf'{border_prop}\s*:\s*(\d+)\s*px\s*(\w+)', body)
            if bp_match:
                props['border_width'] = float(bp_match.group(1))
                props['border_style'] = bp_match.group(2)
                break

        if props:
            # Merge instead of overwrite (multiple selectors can target same class)
            if class_name in rules:
                rules[class_name].update(props)
            else:
                rules[class_name] = props

    return rules


def _collect_embedded_css(soup) -> dict:
    """Parse bounded inline ``<style>`` content without network access."""
    css_rules = {}
    remaining_chars = 512 * 1024
    for style_tag in soup.find_all("style"):
        if remaining_chars <= 0:
            break
        css_text = style_tag.get_text()[:remaining_chars]
        remaining_chars -= len(css_text)
        css_rules.update(_parse_css_rules(css_text))
    return css_rules


def _convert_html_xls(
    input_path: Path,
    output_path: Path,
    progress: Callable,
    soup=None,
    html: str | None = None,
) -> None:
    """Convert HTML-disguised .xls to .xlsx using BeautifulSoup + openpyxl.

    Args:
        soup: Pre-parsed BeautifulSoup instance (passed from convert_file to avoid double-parse).
        html: Raw HTML string (passed from convert_file to avoid re-decoding).
    """
    progress("Membaca file .xls (HTML)...")

    # Use pre-parsed soup if provided, otherwise parse from scratch
    if soup is None or html is None:
        with open(input_path, "rb") as f:
            raw = f.read()
        html = decode_with_fallback(raw)
        soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")

    if not tables:
        raise ValueError("File HTML .xls tidak memiliki elemen <table>.")

    # Pick best table by content scoring
    target_table = max(tables, key=_score_table)
    progress(f"Ditemukan table dengan {len(target_table.find_all('tr'))} baris.")

    # Identify file type
    file_type = identify_file_type(soup, target_table)
    if file_type:
        progress(f"Tipe file: {file_type}")
    else:
        progress("WARNING: Tipe file tidak dikenali, konversi tetap dilanjutkan.")

    # Parse embedded CSS only. External stylesheets are deliberately ignored
    # so conversion remains offline and untrusted files cannot trigger SSRF.
    css_rules = _collect_embedded_css(soup)
    if css_rules:
        progress(f"Loaded {len(css_rules)} embedded CSS rules")

    rows = target_table.find_all("tr")
    if not rows:
        raise ValueError("Table utama tidak memiliki baris data.")

    # Build grid from ALL rows
    # For Pemeriksaan ANC, include URLs from <a> tags in last column
    include_links = file_type == "Pemeriksaan ANC"
    grid, merge_ranges, bold_cells, cell_hyperlinks, cell_styles = _build_dom_grid(rows, include_links=include_links, css_rules=css_rules)

    # Collect pre-table content (type-specific)
    pre_rows = _collect_pre_table_content(soup, target_table, file_type)

    # Calculate row offset for merge ranges (pre-table content adds rows before main table)
    row_offset = len(pre_rows)

    # Combine: pre-table content + table grid
    all_rows = list(pre_rows)
    all_rows.extend(grid)

    # Strip trailing completely empty rows
    all_rows = _strip_trailing_empty(all_rows)

    nrows = len(all_rows)
    ncols = max(len(r) for r in all_rows) if all_rows else 0

    progress(f"Grid: {nrows} baris x {ncols} kolom. Menulis ke .xlsx...")

    # Determine sheet name (extract from HTML XML if available)
    sheet_name = _determine_sheet_name(file_type, html=html)
    progress(f"Menggunakan sheet name: '{sheet_name}'")

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Write cells with type inference, then apply styling
    # Check if table has content_table class (all cells get borders + default font)
    has_content_table = "content_table" in css_rules

    for r in range(nrows):
        grid_r = r - row_offset  # 0-based grid row index
        for c in range(len(all_rows[r])):
            val = all_rows[r][c]
            if val:
                # Infer proper type (datetime, int, float, or string)
                typed_val = _infer_value(val) if isinstance(val, str) else val
                openpyxl_cell = ws.cell(row=r + 1, column=c + 1, value=typed_val)

                # Apply number format for datetime values: use mm-dd-yy (matches expected output)
                if isinstance(typed_val, datetime.datetime):
                    openpyxl_cell.number_format = "mm-dd-yy"
            else:
                # Create empty cell for styling
                openpyxl_cell = ws.cell(row=r + 1, column=c + 1)

            # Apply content_table default styling (borders + font size) to ALL cells
            if has_content_table and grid_r >= 0:
                base_rule = css_rules.get("content_table", {})
                if base_rule:
                    # Apply default font size (xx-small = 8pt overrides Excel default 11pt)
                    if "font_size" in base_rule:
                        openpyxl_cell.font = Font(size=base_rule["font_size"])

                    # Apply default border to all cells
                    border_color = base_rule.get("border_color", "000000")
                    side = Side(style="thin", color=border_color)
                    if not openpyxl_cell.border.top:
                        openpyxl_cell.border = Border(left=side, right=side, top=side, bottom=side)

                # Apply inline styles (bg_color, font_size, font_color, border, bold)
                if (grid_r, c) in cell_styles:
                    style = cell_styles[(grid_r, c)]

                    # Build font with all properties — preserve base font_size from content_table
                    base_font_size = None
                    if has_content_table and "font_size" in css_rules.get("content_table", {}):
                        base_font_size = css_rules["content_table"]["font_size"]

                    font_kw = {}
                    # Bold from either HTML detection or CSS class
                    if openpyxl_cell.font and openpyxl_cell.font.bold:
                        font_kw["bold"] = True
                    elif style.get("bold"):
                        font_kw["bold"] = True
                    # Font size: from style or fall back to base
                    if "font_size" in style:
                        font_kw["size"] = style["font_size"]
                    elif base_font_size:
                        font_kw["size"] = base_font_size
                    if "font_color" in style:
                        font_kw["color"] = style["font_color"]
                    if font_kw:
                        openpyxl_cell.font = Font(**font_kw)

                    # Apply background fill
                    if "bg_color" in style:
                        openpyxl_cell.fill = PatternFill(
                            start_color=style["bg_color"],
                            end_color=style["bg_color"],
                            fill_type="solid"
                        )

                    # Apply inline border override if present
                    if "border" in style:
                        b = style["border"]
                        side_kw = {"style": "thin"}
                        if b.get("width", 1) >= 2:
                            side_kw["style"] = "medium"
                        if b.get("color"):
                            side_kw["color"] = b["color"]
                        side = Side(**side_kw)
                        openpyxl_cell.border = Border(
                            left=side, right=side, top=side, bottom=side
                        )
                    elif style.get("border_color") or style.get("border_width"):
                        # CSS class border (e.g., content_header)
                        bc = style.get("border_color", "000000")
                        bw = style.get("border_width", 1)
                        side_kw = {"style": "thin"}
                        if bw >= 2:
                            side_kw["style"] = "medium"
                        side_kw["color"] = bc
                        side = Side(**side_kw)
                        openpyxl_cell.border = Border(left=side, right=side, top=side, bottom=side)

            # Hyperlinks are independent of CSS classes. Apply them for generic
            # HTML tables as well as the known DCApps export layouts.
            if grid_r >= 0 and (grid_r, c) in cell_hyperlinks:
                openpyxl_cell.hyperlink = cell_hyperlinks[(grid_r, c)]

    # Apply merge cells from HTML rowspan/colspan
    if merge_ranges:
        progress(f"Menerapkan {len(merge_ranges)} merge cells...")
        for start_row, start_col, end_row, end_col in merge_ranges:
            # Adjust for pre-table row offset and convert to 1-based indexing
            adj_start_row = start_row + row_offset + 1
            adj_end_row = end_row + row_offset + 1
            adj_start_col = start_col + 1
            adj_end_col = end_col + 1

            # Validate merge range is within grid bounds
            if adj_end_row <= nrows and adj_end_col <= ncols:
                try:
                    ws.merge_cells(
                        start_row=adj_start_row,
                        start_column=adj_start_col,
                        end_row=adj_end_row,
                        end_column=adj_end_col,
                    )
                except Exception as e:
                    # Log but don't fail on merge errors (e.g., overlapping ranges)
                    progress(f"  Warning: Merge gagal di R{adj_start_row}C{adj_start_col}:R{adj_end_row}C{adj_end_col}: {e}")

    # Auto-fit column widths (sample all rows for accuracy)
    for c in range(ncols):
        max_len = 0
        for r in range(nrows):
            if c < len(all_rows[r]):
                cell_len = len(all_rows[r][c])
                max_len = max(max_len, cell_len)
        col_letter = get_column_letter(c + 1)
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 60)

    wb.save(str(output_path))
    progress("Konversi HTML .xls selesai.")


def _determine_sheet_name(file_type: str | None, soup=None, html: str | None = None) -> str:
    """Determine sheet name based on file type and HTML content.

    For Sigizi files: use 'Application List' (from Excel XML in HTML).
    For Sihepi Rekap 3e: extract year from content and construct name.
    """
    # Try to extract sheet name from Excel XML in HTML using regex
    # (BeautifulSoup can't handle x: namespace tags reliably)
    if html:
        xml_match = re.search(r'<x:Name>(.*?)</x:Name>', html)
        if xml_match and xml_match.group(1).strip():
            return _sanitize_sheet_name(xml_match.group(1))

    # Fallback: construct based on file type
    sheet_names = {
        "Rekap 3e": "Rekap_3e",
        "Daftar Ibu Hamil": "Daftar_Ibu_Hamil",
        "Daftar Ibu Nifas": "Daftar_Ibu_Nifas",
        "Daftar Ibu": "Daftar_Ibu",
        "Kohort Ibu": "Kohort_Ibu",
        "Pemeriksaan ANC": "Pemeriksaan_ANC",
        "Posyandu BNBA": "Posyandu_BNBA",
    }
    if file_type and file_type in sheet_names:
        return _sanitize_sheet_name(sheet_names[file_type])
    return "Sheet1"


def _sanitize_sheet_name(value: object) -> str:
    """Return a valid, bounded Excel worksheet title."""
    candidate = str(value or "").strip().strip("'")
    candidate = re.sub(r"[\\/*?:\[\]]", "_", candidate)
    return candidate[:31] or "Sheet1"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def convert_file(
    input_path: str | Path,
    output_path: str | Path | None = None,
    progress_callback: Callable[[str], None] | None = None,
    *,
    overwrite: bool = False,
) -> Path:
    """Convert an .xls file to .xlsx preserving content and formatting.

    Args:
        input_path: Path to .xls file.
        output_path: Desired output path (optional). Auto-generated if omitted.
        progress_callback: Function to receive status messages.
        overwrite: Replace an existing output only when explicitly enabled.

    Returns:
        Path to the generated .xlsx file.

    Raises:
        FileNotFoundError: If input file doesn't exist.
        ValueError: If file is not a valid .xls or conversion fails.
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"File tidak ditemukan: {input_path}")

    if output_path is None:
        stem = input_path.stem
        output_path = input_path.with_name(f"{stem}_converted.xlsx")
    else:
        output_path = Path(output_path)
        if output_path.suffix.lower() not in (".xlsx", ".xlsm"):
            output_path = output_path.with_suffix(".xlsx")

    real_output = output_path.resolve()
    real_input = input_path.resolve()
    if real_output == real_input:
        raise ValueError("Input and output paths must be different.")
    if not real_output.parent.is_dir():
        raise ValueError(f"Output directory does not exist: {real_output.parent}")
    if real_output.exists() and not overwrite:
        raise FileExistsError(
            f"Output already exists: {real_output}. Enable overwrite explicitly to replace it."
        )

    logger.info("Converting %s -> %s", real_input, real_output)

    log = progress_callback or (lambda msg: None)
    temporary_path: Path | None = None

    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{real_output.stem}.",
            suffix=".xlsx",
            dir=real_output.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        file_format = detect_file_format(input_path)
        log(f"Format file: {file_format.upper()}")

        if file_format == "binary":
            _convert_binary_xls(input_path, temporary_path, log)
        else:
            # For HTML files: parse ONCE, then validate type and convert
            with open(input_path, "rb") as f:
                raw = f.read()

            # Multi-encoding fallback (consistent with _convert_html_xls)
            html_content = decode_with_fallback(raw)

            soup = BeautifulSoup(html_content, "lxml")
            tables = soup.find_all("table")
            if tables:
                target_table = max(tables, key=_score_table)
                file_type = identify_file_type(soup, target_table)
                if file_type is None:
                    log("Tipe file tidak dikenali; menggunakan konversi tabel generik.")
                else:
                    log(f"Tipe file: {file_type}")

            # Pass pre-parsed soup and html to avoid double-parse
            _convert_html_xls(
                input_path,
                temporary_path,
                log,
                soup=soup,
                html=html_content,
            )

        os.replace(temporary_path, real_output)
        temporary_path = None
        size_mb = os.path.getsize(real_output) / (1024 * 1024)
        log(f"Output: {real_output.name} ({size_mb:.1f} MB)")
        return real_output

    except Exception as exc:
        logger.exception("Conversion failed")
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
        if isinstance(exc, (FileExistsError, FileNotFoundError)):
            raise
        raise ValueError(f"Konversi gagal: {exc}") from exc


def identify_and_convert(
    input_path: str | Path,
    output_path: str | Path | None = None,
    progress_callback: Callable[[str], None] | None = None,
    *,
    overwrite: bool = False,
) -> tuple[Path, str | None]:
    """Convert with file type identification.

    Thin wrapper around convert_file that returns the identified file type.
    Binary files return None for file_type (type detection only applies to HTML).

    Returns:
        Tuple of (output_path, file_type_name or None)

    """
    input_path = Path(input_path)
    file_format = detect_file_format(input_path)

    if file_format != "html":
        return convert_file(
            input_path,
            output_path,
            progress_callback,
            overwrite=overwrite,
        ), None

    # Parse HTML once to identify type, then let convert_file reuse the logic
    with open(input_path, "rb") as f:
        raw = f.read()

    html = decode_with_fallback(raw)
    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")

    if tables:
        target_table = max(tables, key=_score_table)
        file_type = identify_file_type(soup, target_table)
        return convert_file(
            input_path,
            output_path,
            progress_callback,
            overwrite=overwrite,
        ), file_type

    return convert_file(
        input_path,
        output_path,
        progress_callback,
        overwrite=overwrite,
    ), None
