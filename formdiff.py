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
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook
from colorama import init, Fore, Style

init(autoreset=True)

ARRAY_FORMULA_RE = re.compile(r"^\{(=.*)\}$")
# Matches a cell coordinate like A1, BC204, $AF$3 — captures column letters and row number
COORD_RE = re.compile(r"\$?([A-Za-z]+)\$?(\d+)$")


def _normalize_formula(raw: str) -> str:
    """Strip array-formula braces and whitespace so comparisons are canonical."""
    stripped = raw.strip()
    m = ARRAY_FORMULA_RE.match(stripped)
    if m:
        return m.group(1)
    return stripped


def _coord_sort_key(entry: str) -> tuple:
    """
    Return a sort key that orders entries by sheet name (alphabetical) then
    by cell coordinate in natural spreadsheet order: column A-Z, then row 1-N
    numerically.  Falls back to the raw string if the coordinate can't be parsed.

    Example order: A1, A2, A10, B1, B2 ... (not A1, A10, A2, B1 ...)
    """
    # entry format: "SheetName!COL ROW: =FORMULA"
    bang = entry.find("!")
    colon = entry.find(":", bang)
    if bang == -1 or colon == -1:
        return (entry,)
    sheet = entry[:bang]
    coord = entry[bang + 1 : colon]
    m = COORD_RE.search(coord)
    if not m:
        return (sheet, coord)
    col_letters = m.group(1).upper()
    row_num = int(m.group(2))
    # Convert column letters to a numeric index so AA sorts after Z
    col_index = 0
    for ch in col_letters:
        col_index = col_index * 26 + (ord(ch) - ord("A") + 1)
    return (sheet, col_index, row_num)


def extract_formulas(file_path: str, ignore_sheets: set[str] | None = None) -> list[str]:
    """
    Open *file_path* in read-only mode and return a sorted list of formula
    strings in the format ``SheetName!A1: =SUM(B1:B10)``.

    Static values, blanks, dates, and plain numbers are skipped.
    ``read_only=True`` keeps memory usage low for large workbooks.
    Sheets whose names appear in *ignore_sheets* are skipped entirely.
    Entries are sorted in natural spreadsheet order (A1, A2...A10, B1...).
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
    records.sort(key=_coord_sort_key)
    return records


# ---------------------------------------------------------------------------
# Summary helpers
# ---------------------------------------------------------------------------

def _parse_diff_entry(line: str) -> tuple[str, str, str] | None:
    """
    Given a raw diff line (with leading +/-/ ), return (sign, sheet, formula).
    Returns None for hunk headers and file headers.
    """
    if not line or line.startswith("@@") or line.startswith("---") or line.startswith("+++"):
        return None
    sign = line[0]  # '+', '-', or ' '
    rest = line[1:]
    bang = rest.find("!")
    colon = rest.find(": ", bang)
    if bang == -1 or colon == -1:
        return None
    sheet = rest[:bang]
    formula = rest[colon + 2:]
    return (sign, sheet, formula)


def print_summary(diff_lines: list[str]) -> None:
    """
    Print a human-readable grouped summary of what changed after the raw diff.
    Groups identical formula-level changes (e.g. same substitution pattern
    repeated across many cells) and counts them per sheet.
    """
    # Collect removed and added formulas per sheet
    removed: dict[str, list[str]] = defaultdict(list)
    added: dict[str, list[str]] = defaultdict(list)

    for line in diff_lines:
        parsed = _parse_diff_entry(line)
        if parsed is None:
            continue
        sign, sheet, formula = parsed
        if sign == "-":
            removed[sheet].append(formula)
        elif sign == "+":
            added[sheet].append(formula)

    all_sheets = sorted(set(list(removed) + list(added)))
    if not all_sheets:
        return

    print(Fore.CYAN + Style.BRIGHT + "\n─── Change Summary ───────────────────────────────────")

    for sheet in all_sheets:
        r = removed[sheet]
        a = added[sheet]
        n_removed = len(r)
        n_added = len(a)

        print(Style.BRIGHT + f"\n  {sheet}")

        if n_removed == 0:
            print(Fore.GREEN + f"    {n_added} formula(s) added")
        elif n_added == 0:
            print(Fore.RED + f"    {n_removed} formula(s) removed")
        else:
            # Try to surface repeated substitution patterns.
            # Find pairs where only a substring changed (e.g. $C$52 → $C$62).
            pair_count = min(n_removed, n_added)
            pattern_counts: dict[tuple[str, str], int] = defaultdict(int)
            unpaired = 0

            for old_f, new_f in zip(r[:pair_count], a[:pair_count]):
                if old_f == new_f:
                    continue
                # Find the first and last differing character to extract the changed segment
                lo, hi_old, hi_new = 0, len(old_f), len(new_f)
                while lo < min(hi_old, hi_new) and old_f[lo] == new_f[lo]:
                    lo += 1
                while hi_old > lo and hi_new > lo and old_f[hi_old - 1] == new_f[hi_new - 1]:
                    hi_old -= 1
                    hi_new -= 1
                changed_from = old_f[lo:hi_old]
                changed_to = new_f[lo:hi_new]
                if changed_from and changed_to:
                    pattern_counts[(changed_from, changed_to)] += 1
                else:
                    unpaired += 1

            total_changed = n_removed + n_added - pair_count * 2 + pair_count
            print(f"    {max(n_removed, n_added)} formula(s) modified")

            for (frm, to), count in sorted(pattern_counts.items(), key=lambda x: -x[1]):
                print(
                    Fore.RED   + f"      {count}x  {frm}"
                    + Style.RESET_ALL + "  →  "
                    + Fore.GREEN + f"{to}"
                )
            if abs(n_removed - n_added) > 0:
                net = n_added - n_removed
                label = "added" if net > 0 else "removed"
                print(f"    {abs(net)} formula(s) net {label}")

    print(Fore.CYAN + "──────────────────────────────────────────────────────\n")


# ---------------------------------------------------------------------------
# Core diff runner
# ---------------------------------------------------------------------------

def run_diff(
    base_file: str,
    target_file: str,
    *,
    output_path: str | None = None,
    context_lines: int = 3,
    ignore_sheets: set[str] | None = None,
    summary: bool = False,
) -> bool:
    """
    Compare formulas between *base_file* and *target_file*.

    Returns ``True`` when at least one difference is found.
    If *output_path* is given the raw unified diff (no ANSI codes) is written
    to that file instead of being printed to stdout.
    Sheets listed in *ignore_sheets* are excluded from both files.
    If *summary* is True, a grouped change digest is printed after the diff.
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

    if summary:
        print_summary(diff_lines)

    return True


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

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
    parser.add_argument(
        "-s",
        "--summary",
        action="store_true",
        default=False,
        help="Print a grouped change summary after the diff",
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
        summary=args.summary,
    )
    sys.exit(1 if changed else 0)


if __name__ == "__main__":
    main()
