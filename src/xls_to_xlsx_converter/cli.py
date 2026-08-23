from __future__ import annotations

import argparse
import sys
from pathlib import Path

from xls_to_xlsx_converter.converter import convert_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xls2xlsx",
        description="Convert binary or HTML-disguised .xls files to .xlsx offline.",
    )
    parser.add_argument("inputs", nargs="+", type=Path, help="Input .xls file(s).")
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Existing output directory. Defaults to each input directory.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing output files explicitly.",
    )
    parser.add_argument("--quiet", action="store_true", help="Print only errors.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    output_dir = args.output_dir
    if output_dir is not None and not output_dir.is_dir():
        print(f"Output directory does not exist: {output_dir}", file=sys.stderr)
        return 2

    failures = 0
    for input_path in args.inputs:
        destination = (output_dir or input_path.parent) / (
            input_path.stem + "_converted.xlsx"
        )
        try:
            result = convert_file(
                input_path,
                destination,
                overwrite=args.overwrite,
            )
            if not args.quiet:
                print(f"Converted: {input_path} -> {result}")
        except (FileNotFoundError, FileExistsError, ValueError) as exc:
            failures += 1
            print(f"Failed: {input_path}: {exc}", file=sys.stderr)

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
