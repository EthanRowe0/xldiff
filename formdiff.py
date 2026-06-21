#!/usr/bin/env python3
"""
FormDiff — Git-like line-by-line formula diffing for Excel workbooks.

Extracts raw formula definitions from two .xlsx files, serializes them into
a sorted linear array of  SheetName!Coordinate: =FORMULA  entries, and runs
a unified diff to surface additions, deletions, and modifications.
"""

import sys
import argparse
import difflib
import re
from pathlib import Path

from openpyxl import load_workbook
from colorama import init, Fore, Style

init(autoreset=True)

ARRAY_FORMULA_RE = re.compile(r"^\{(=.*)\}$")


def _normalize_formula(raw: str) -> str:
    """Strip array-formula braces and whitespace so comparisons are canonical."""
    stripped = raw.strip()
    m = ARRAY_FORMULA_RE.match(stripped)
    if m:
        return m.group(1)
    return stripped


def extract_formulas(file_path: str, ignore_sheets: set[str] | None = None) -> list[str]:
    """
    Open *file_path* in read-only mode and return a sorted list of formula
    strings in the format ``SheetName!A1: =SUM(B1:B10)``.

    Static values, blanks, dates, and plain numbers are skipped.
    ``read_only=True`` keeps memory usage low for large workbooks.
    Sheets whose names appear in *ignore_sheets* are skipped entirely.
    """
    try:
        wb = load_workbook(file_path, data_only=False, read_only=True)
    except Exception as e:
        print(f"Error opening workbook {file_path}: {e}", file=sys.stderr)
        sys.exit(1)

    skipped = ignore_sheets or set()
    records: list[str] = []
    for sheet_name in wb.sheetnames:
        if sheet_name in skipped:
            continue
        sheet = wb[sheet_name]
        for row in sheet.iter_rows(values_only=False):
            for cell in row:
                val = cell.value
                if isinstance(val, str) and val.startswith("="):
                    formula = _normalize_formula(val)
                    records.append(f"{sheet_name}!{cell.coordinate}: {formula}")

    wb.close()
    records.sort()
    return records


def run_diff(
    base_file: str,
    target_file: str,
    *,
    output_path: str | None = None,
    context_lines: int = 3,
    ignore_sheets: set[str] | None = None,
) -> bool:
    """
    Compare formulas between *base_file* and *target_file*.

    Returns ``True`` when at least one difference is found.
    If *output_path* is given the raw unified diff (no ANSI codes) is written
    to that file instead of being printed to stdout.
    Sheets listed in *ignore_sheets* are excluded from both files.
    """
    base_set = extract_formulas(base_file, ignore_sheets=ignore_sheets)
    target_set = extract_formulas(target_file, ignore_sheets=ignore_sheets)

    diff_lines = list(
        difflib.unified_diff(
            base_set,
            target_set,
            fromfile=base_file,
            tofile=target_file,
            lineterm="",
            n=context_lines,
        )
    )

    if not diff_lines:
        msg = "FormDiff complete: zero formula deviations detected."
        if output_path:
            Path(output_path).write_text(msg + "\n", encoding="utf-8")
        else:
            print(Fore.GREEN + msg)
        return False

    if output_path:
        Path(output_path).write_text("\n".join(diff_lines) + "\n", encoding="utf-8")
        print(f"Diff written to {output_path}")
    else:
        for line in diff_lines:
            if line.startswith("+++") or line.startswith("---"):
                print(Style.BRIGHT + line)
            elif line.startswith("+"):
                print(Fore.GREEN + line)
            elif line.startswith("-"):
                print(Fore.RED + line)
            elif line.startswith("@@"):
                print(Fore.CYAN + line)
            else:
                print(line)

    return True


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="formdiff",
        description="Git-like formula diff for Excel workbooks.",
    )
    parser.add_argument("base", help="Path to the base (old) .xlsx workbook")
    parser.add_argument("target", help="Path to the target (new) .xlsx workbook")
    parser.add_argument(
        "-o",
        "--output",
        metavar="FILE",
        default=None,
        help="Write the unified diff to FILE instead of stdout (ANSI codes stripped)",
    )
    parser.add_argument(
        "-C",
        "--context",
        type=int,
        default=3,
        metavar="N",
        help="Number of context lines around each change (default: 3)",
    )
    parser.add_argument(
        "-x",
        "--exclude-sheet",
        action="append",
        metavar="SHEET",
        dest="exclude_sheets",
        default=[],
        help="Sheet name to ignore (can be repeated: -x LTOP -x Summary)",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    for path in (args.base, args.target):
        if not Path(path).is_file():
            print(f"File not found: {path}", file=sys.stderr)
            sys.exit(1)

    ignore_sheets = set(args.exclude_sheets) if args.exclude_sheets else None
    if ignore_sheets:
        print(f"Ignoring sheets: {', '.join(sorted(ignore_sheets))}")

    print(f"FormDiff: {args.base} → {args.target}\n")
    changed = run_diff(
        args.base,
        args.target,
        output_path=args.output,
        context_lines=args.context,
        ignore_sheets=ignore_sheets,
    )
    sys.exit(1 if changed else 0)


if __name__ == "__main__":
    main()
