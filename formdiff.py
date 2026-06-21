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

ARRAY_FORMULA_RE = re.compile(r"^\{(=.*)\}$", re.DOTALL)
# Matches a cell coordinate like A1, BC204, $AF$3 — captures column letters and row number
COORD_RE = re.compile(r"\$?([A-Za-z]+)\$?(\d+)$")
# Matches the coord portion of a serialized entry: "Sheet!A1: =..."
ENTRY_COORD_RE = re.compile(r"^(.+)!(\$?[A-Za-z]+\$?\d+): ")


def _normalize_formula(raw: str) -> str:
    """
    Strip array-formula braces, collapse embedded newlines, and trim whitespace
    so each formula serializes as a single line and comparisons are canonical.
    """
    stripped = raw.strip()
    m = ARRAY_FORMULA_RE.match(stripped)
    if m:
        stripped = m.group(1)
    # Replace any embedded newlines (Excel line-break characters) with a space
    return re.sub(r"[\r\n]+", " ", stripped).strip()


def _coord_sort_key(entry: str) -> tuple:
    """
    Return a sort key that orders entries by sheet name (alphabetical) then
    by cell coordinate in natural spreadsheet order: column A-Z, then row 1-N
    numerically.  Falls back to the raw string if the coordinate can't be parsed.

    Example order: A1, A2, A10, B1, B2 ... (not A1, A10, A2, B1 ...)
    """
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
    col_index = 0
    for ch in col_letters:
        col_index = col_index * 26 + (ord(ch) - ord("A") + 1)
    return (sheet, col_index, row_num)


def extract_formulas(
    file_path: str,
    ignore_sheets: set[str] | None = None,
    only_sheets: set[str] | None = None,
) -> list[str]:
    """
    Open *file_path* in read-only mode and return a sorted list of formula
    strings in the format ``SheetName!A1: =SUM(B1:B10)``.

    Static values, blanks, dates, and plain numbers are skipped.
    ``read_only=True`` keeps memory usage low for large workbooks.
    Sheets in *ignore_sheets* are skipped entirely.
    If *only_sheets* is provided, every sheet NOT in that set is skipped
    (whitelist mode). *ignore_sheets* is still applied on top of *only_sheets*.
    Entries are sorted in natural spreadsheet order (A1, A2...A10, B1...).
    Embedded newlines inside formulas are collapsed to a single space.
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
        if only_sheets is not None and sheet_name not in only_sheets:
            continue
        sheet = wb[sheet_name]
        for row in sheet.iter_rows(values_only=False):
            for cell in row:
                val = cell.value
                if isinstance(val, str) and val.lstrip().startswith("="):
                    formula = _normalize_formula(val)
                    records.append(f"{sheet_name}!{cell.coordinate}: {formula}")

    wb.close()
    records.sort(key=_coord_sort_key)
    return records


# ---------------------------------------------------------------------------
# Compact diff helpers
# ---------------------------------------------------------------------------

def _parse_entry(line: str) -> tuple[str, str, str, str] | None:
    """
    Parse a raw diff line into (sign, sheet, coord, formula).
    Returns None for hunk/file headers or unparseable lines.
    """
    if not line or line.startswith("@@") or line.startswith("---") or line.startswith("+++"):
        return None
    sign = line[0]
    rest = line[1:]
    m = ENTRY_COORD_RE.match(rest)
    if not m:
        return None
    sheet = m.group(1)
    coord = m.group(2)
    formula = rest[m.end():]
    return (sign, sheet, coord, formula)


def _col_letters(coord: str) -> str:
    """Extract just the column letters from a coordinate like $AF$3 or D119."""
    m = COORD_RE.search(coord.replace("$", ""))
    return m.group(1).upper() if m else ""


def _row_num(coord: str) -> int:
    """Extract the row number from a coordinate."""
    m = COORD_RE.search(coord)
    return int(m.group(2)) if m else -1


def _col_to_index(col: str) -> int:
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def build_compact_diff(diff_lines: list[str]) -> list[str]:
    """
    Return a compacted version of *diff_lines* where runs of consecutive cells
    in the same column that all have the same formula change are collapsed into
    a single annotated line:

        -[×N] Sheet!ColFIRST:ColLAST  <old_formula_template>
        +[×N] Sheet!ColFIRST:ColLAST  <new_formula_template>

    Context lines and hunk headers are preserved as-is.
    A run is only collapsed when N >= 2.
    """
    # First pass: collect all changed entries keyed by (sign, sheet, col, formula)
    # so we can detect runs of consecutive rows.

    # We'll work entry by entry and group runs on the fly.
    # Strategy: scan through diff_lines; accumulate consecutive (+/-) entries
    # that share the same sheet + column + formula pattern; flush when broken.

    result: list[str] = []

    # A "run" is a list of (sign, sheet, coord, formula, original_line)
    pending_removed: list[tuple] = []
    pending_added: list[tuple] = []

    def _flush(removed: list[tuple], added: list[tuple]) -> list[str]:
        """Collapse or emit a matched block of removed/added lines."""
        out: list[str] = []
        n_rem = len(removed)
        n_add = len(added)

        # Only collapse when counts match, same col, same template, consecutive rows
        def is_collapsible(entries: list[tuple]) -> bool:
            if len(entries) < 2:
                return False
            sheet0, col0, formula0 = entries[0][1], _col_letters(entries[0][2]), entries[0][3]
            for i, (_, sh, coord, fml, _) in enumerate(entries):
                if sh != sheet0 or _col_letters(coord) != col0:
                    return False
                # rows must be consecutive
                if i > 0 and _row_num(entries[i][2]) != _row_num(entries[i-1][2]) + 1:
                    return False
            return True

        if n_rem == n_add and n_rem >= 2 and is_collapsible(removed) and is_collapsible(added):
            n = n_rem
            sheet = removed[0][1]
            first_coord = removed[0][2]
            last_coord  = removed[-1][2]
            col = _col_letters(first_coord)
            first_row = _row_num(first_coord)
            last_row  = _row_num(last_coord)
            range_label = f"{sheet}!{col}{first_row}:{col}{last_row}"
            old_formula = removed[0][3]
            new_formula = added[0][3]
            out.append(f"-[×{n}] {range_label}: {old_formula}")
            out.append(f"+[×{n}] {range_label}: {new_formula}")
        else:
            for _, _, _, _, orig in removed:
                out.append(orig)
            for _, _, _, _, orig in added:
                out.append(orig)
        return out

    for line in diff_lines:
        parsed = _parse_entry(line)

        if parsed is None:
            # Non-entry line — flush pending and emit as-is
            result.extend(_flush(pending_removed, pending_added))
            pending_removed.clear()
            pending_added.clear()
            result.append(line)
            continue

        sign, sheet, coord, formula = parsed

        if sign == " ":
            result.extend(_flush(pending_removed, pending_added))
            pending_removed.clear()
            pending_added.clear()
            result.append(line)
        elif sign == "-":
            # If this breaks the current removed run, flush first
            if pending_removed:
                prev = pending_removed[-1]
                same_col = (_col_letters(prev[2]) == _col_letters(coord) and prev[1] == sheet)
                consecutive = (_row_num(coord) == _row_num(prev[2]) + 1)
                same_formula = (prev[3] == formula)
                if not (same_col and consecutive and same_formula):
                    result.extend(_flush(pending_removed, pending_added))
                    pending_removed.clear()
                    pending_added.clear()
            pending_removed.append((sign, sheet, coord, formula, line))
        elif sign == "+":
            if pending_added:
                prev = pending_added[-1]
                same_col = (_col_letters(prev[2]) == _col_letters(coord) and prev[1] == sheet)
                consecutive = (_row_num(coord) == _row_num(prev[2]) + 1)
                same_formula = (prev[3] == formula)
                if not (same_col and consecutive and same_formula):
                    result.extend(_flush(pending_removed, pending_added))
                    pending_removed.clear()
                    pending_added.clear()
            pending_added.append((sign, sheet, coord, formula, line))

    result.extend(_flush(pending_removed, pending_added))
    return result


def _compact_output_path(output_path: str) -> str:
    """Insert '-compact' before the file extension: foo.diff → foo-compact.diff"""
    p = Path(output_path)
    return str(p.with_stem(p.stem + "-compact"))


# ---------------------------------------------------------------------------
# Summary helpers
# ---------------------------------------------------------------------------

def _parse_diff_entry(line: str) -> tuple[str, str, str, str] | None:
    """
    Given a raw diff line (with leading +/-/ ), return (sign, sheet, coord, formula).
    Returns None for hunk headers and file headers.
    """
    parsed = _parse_entry(line)
    if parsed is None:
        return None
    sign, sheet, coord, formula = parsed
    return (sign, sheet, coord, formula)


def print_summary(diff_lines: list[str]) -> None:
    """
    Print a human-readable grouped summary of what changed.

    Per sheet, shows:
    - Net additions / deletions
    - For paired changes: the substitution pattern with count and cell ranges
      e.g.  18x  $C$52 → $C$62   (V1119:V1196)
    """
    # Collect per-sheet lists of (coord, formula) for removed and added
    removed: dict[str, list[tuple[str, str]]] = defaultdict(list)
    added:   dict[str, list[tuple[str, str]]] = defaultdict(list)

    for line in diff_lines:
        parsed = _parse_diff_entry(line)
        if parsed is None:
            continue
        sign, sheet, coord, formula = parsed
        if sign == "-":
            removed[sheet].append((coord, formula))
        elif sign == "+":
            added[sheet].append((coord, formula))

    all_sheets = sorted(set(list(removed) + list(added)))
    if not all_sheets:
        return

    print(Fore.CYAN + Style.BRIGHT + "\n─── Change Summary ───────────────────────────────────")

    for sheet in all_sheets:
        r = removed[sheet]   # list of (coord, formula)
        a = added[sheet]
        n_removed = len(r)
        n_added = len(a)

        print(Style.BRIGHT + f"\n  {sheet}")

        if n_removed == 0:
            # Pure additions — list ranges
            ranges = _coords_to_ranges([coord for coord, _ in a])
            print(Fore.GREEN + f"    {n_added} formula(s) added  [{ranges}]")
        elif n_added == 0:
            ranges = _coords_to_ranges([coord for coord, _ in r])
            print(Fore.RED + f"    {n_removed} formula(s) removed  [{ranges}]")
        else:
            pair_count = min(n_removed, n_added)
            # Group by substitution pattern, tracking which coords had that pattern
            pattern_coords: dict[tuple[str, str], list[str]] = defaultdict(list)

            for (coord, old_f), (_, new_f) in zip(r[:pair_count], a[:pair_count]):
                if old_f == new_f:
                    continue
                lo, hi_old, hi_new = 0, len(old_f), len(new_f)
                while lo < min(hi_old, hi_new) and old_f[lo] == new_f[lo]:
                    lo += 1
                while hi_old > lo and hi_new > lo and old_f[hi_old - 1] == new_f[hi_new - 1]:
                    hi_old -= 1
                    hi_new -= 1
                changed_from = old_f[lo:hi_old]
                changed_to   = new_f[lo:hi_new]
                if changed_from and changed_to:
                    pattern_coords[(changed_from, changed_to)].append(coord)

            print(f"    {max(n_removed, n_added)} formula(s) modified")

            for (frm, to), coords in sorted(pattern_coords.items(), key=lambda x: -len(x[1])):
                count = len(coords)
                ranges = _coords_to_ranges(coords)
                print(
                    Fore.RED + f"      {count}x  {frm}"
                    + Style.RESET_ALL + "  →  "
                    + Fore.GREEN + f"{to}"
                    + Style.DIM + f"   ({ranges})"
                )

            net = n_added - n_removed
            if net != 0:
                label = "added" if net > 0 else "removed"
                print(f"    {abs(net)} formula(s) net {label}")

    print(Fore.CYAN + "──────────────────────────────────────────────────────\n")


def _coords_to_ranges(coords: list[str]) -> str:
    """
    Collapse a list of cell coordinates into compact range notation.

    Consecutive rows in the same column are merged:
    [D3, D4, D5, D10] → "D3:D5, D10"
    [D3, F5]          → "D3, F5"
    """
    if not coords:
        return ""

    # Parse and sort by (col_index, row)
    parsed = []
    for c in coords:
        m = COORD_RE.search(c.replace("$", ""))
        if m:
            col = m.group(1).upper()
            row = int(m.group(2))
            parsed.append((_col_to_index(col), col, row))

    parsed.sort()

    ranges: list[str] = []
    if not parsed:
        return ", ".join(coords)

    run_col_idx, run_col, run_start, run_end = parsed[0]

    for col_idx, col, row in parsed[1:]:
        if col_idx == run_col_idx and row == run_end + 1:
            run_end = row
        else:
            if run_end == run_start:
                ranges.append(f"{run_col}{run_start}")
            else:
                ranges.append(f"{run_col}{run_start}:{run_col}{run_end}")
            run_col_idx, run_col, run_start, run_end = col_idx, col, row, row

    if run_end == run_start:
        ranges.append(f"{run_col}{run_start}")
    else:
        ranges.append(f"{run_col}{run_start}:{run_col}{run_end}")

    return ", ".join(ranges)


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
    only_sheets: set[str] | None = None,
    summary: bool = False,
    compact: bool = False,
) -> bool:
    """
    Compare formulas between *base_file* and *target_file*.

    Returns ``True`` when at least one difference is found.
    If *output_path* is given the raw unified diff (no ANSI codes) is written
    to that file. When *compact* is also True, an additional '-compact' file
    is written alongside it with run-length collapsed entries.
    """
    base_set = extract_formulas(base_file, ignore_sheets=ignore_sheets, only_sheets=only_sheets)
    target_set = extract_formulas(target_file, ignore_sheets=ignore_sheets, only_sheets=only_sheets)

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
        if compact:
            compact_lines = build_compact_diff(diff_lines)
            compact_path = _compact_output_path(output_path)
            Path(compact_path).write_text("\n".join(compact_lines) + "\n", encoding="utf-8")
            print(f"Compact diff written to {compact_path}")
    else:
        display_lines = build_compact_diff(diff_lines) if compact else diff_lines
        for line in display_lines:
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
        "-i",
        "--include-sheet",
        action="append",
        metavar="SHEET",
        dest="include_sheets",
        default=[],
        help="Only diff this sheet — all others are ignored (can be repeated: -i BOM -i F2)",
    )
    parser.add_argument(
        "-s",
        "--summary",
        action="store_true",
        default=False,
        help="Print a grouped change summary after the diff",
    )
    parser.add_argument(
        "--compact",
        action="store_true",
        default=False,
        help=(
            "Collapse runs of identical consecutive-row changes into a single annotated line. "
            "When used with -o, also writes a separate FILE-compact.diff alongside the full diff."
        ),
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
    only_sheets = set(args.include_sheets) if args.include_sheets else None

    if only_sheets:
        print(f"Only diffing sheets: {', '.join(sorted(only_sheets))}")
    if ignore_sheets:
        print(f"Ignoring sheets: {', '.join(sorted(ignore_sheets))}")

    print(f"FormDiff: {args.base} → {args.target}\n")
    changed = run_diff(
        args.base,
        args.target,
        output_path=args.output,
        context_lines=args.context,
        ignore_sheets=ignore_sheets,
        only_sheets=only_sheets,
        summary=args.summary,
        compact=args.compact,
    )
    sys.exit(1 if changed else 0)


if __name__ == "__main__":
    main()
