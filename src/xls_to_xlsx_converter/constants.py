from __future__ import annotations

import sys
from pathlib import Path


APP_TITLE = "XLS to XLSX Converter v.1.0.0 by LHSCode"
APP_VERSION = "1.0.0"
FOOTER_TEXT = "XLS to XLSX Converter v.1.0.0 | Implemented on DCApps | LHSCode © 2026."

OUTPUT_SUFFIX = "_converted.xlsx"
ALLOWED_EXTENSIONS = {".xls"}


def app_root() -> Path:
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        return Path(bundle_dir)
    return Path(__file__).resolve().parents[2]


def resource_path(*parts: str) -> Path:
    return app_root().joinpath(*parts)
