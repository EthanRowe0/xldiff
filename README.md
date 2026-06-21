# xldiff (FormDiff)

Git-like line-by-line formula diffing for Excel workbooks.

Extracts raw formula definitions from two `.xlsx` files, serializes them into a sorted list, and produces a unified diff highlighting additions, deletions, and modifications — just like `git diff`, but for spreadsheet formulas.

## Features

- **Formula-only diffing** — ignores static values, formatting, dates, and blanks
- **Multi-sheet support** — processes every worksheet, prefixed by sheet name
- **Memory-efficient** — uses `read_only=True` for large workbooks (100k+ cells)
- **Natural coordinate sorting** — cells sort as A1, A2...A10, B1... (not A1, A10, A2...)
- **Whitespace normalization** — cosmetic spacing differences (indentation, spaces around commas/parens) are ignored; only meaningful formula changes are reported
- **Array formula handling** — strips `{=...}` wrapper braces for canonical comparisons
- **Colorized terminal output** — green for additions, red for deletions, cyan for hunk headers
- **HTML report** — auto-generated alongside every `.diff` file; side-by-side Before/After view with inline character-level highlighting
- **Compact mode** — collapses runs of identical consecutive-row changes into a single annotated line
- **Grouped summary** — per-sheet change digest with substitution patterns and cell ranges
- **Sheet filtering** — exclude or whitelist sheets by name per run

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python formdiff.py <old.xlsx> <new.xlsx>
```

### Options

| Flag | Description |
|------|-------------|
| `-o`, `--output FILE` | Write diff to `FILE` and auto-generate `FILE.html` report |
| `-x`, `--exclude-sheet SHEET` | Skip a sheet by name (repeatable: `-x LTOP -x Cover`) |
| `-i`, `--include-sheet SHEET` | Only diff this sheet — all others are skipped (repeatable: `-i BOM -i F2`) |
| `-s`, `--summary` | Print a grouped change summary after the diff |
| `--compact` | Collapse runs of identical consecutive-row changes into one line; with `-o`, also writes `FILE-compact.diff` and `FILE-compact.html` |

### Examples

```bash
# Print diff to terminal
python formdiff.py old.xlsx new.xlsx

# Save diff and auto-generate HTML report
python formdiff.py old.xlsx new.xlsx -o diffs/changes.diff

# Only look at the BOM sheet
python formdiff.py old.xlsx new.xlsx -i BOM -o diffs/bom.diff

# Exclude noisy sheets
python formdiff.py old.xlsx new.xlsx -x LTOP -x Summary

# Compact output + HTML + grouped summary
python formdiff.py old.xlsx new.xlsx -o diffs/changes.diff --compact -s
```

## Output Format

Each formula is serialized as:

```
SheetName!CellRef: =FORMULA
```

The `.diff` file follows standard unified diff format:

```diff
--- old.xlsx
+++ new.xlsx
@@ -10,3 +10,4 @@
-BOM!F3: =COUNTIF($C$3:$C$52,$E3)
+BOM!F3: =COUNTIF($C$3:$C$62,$E3)
```

## HTML Report

Every run with `-o` automatically produces a `.html` file alongside the `.diff`. Open it in any browser — no extra tools needed.

### HTML features

- **Side-by-side Before / After columns** with the cell reference on the left
- **Inline character-level highlighting** — exact changed characters are highlighted in darker red/green so you can instantly see what flipped
- **Ellipsis collapsing** — long unchanged formula segments are hidden by default and snap to natural argument boundaries (commas, parentheses). Click `…` to reveal, or use the per-row and global expand/collapse buttons
- **Collapsible sheet sections** — click a sheet header to collapse/expand it
- **Filter panel** (▾ Filter button) with three controls:
  - **Sheets** — clickable chips to toggle sheets on/off
  - **Column** — comma-separated column names (e.g. `D,V,AE`) to focus on specific columns
  - **Row range** — min/max row number inputs to narrow to a section of the sheet
- **Compact HTML** — when `--compact` is used, a separate `FILE-compact.html` is also generated where run-length-collapsed rows show a `×N` badge

## Compact Mode

When many consecutive cells in the same column share an identical change, `--compact` collapses them:

```
-[×78] L2!V1119:V1196: =IF($B{N}<>"",Lin+$T{N}-...-ABS(TURN1_R_COR),"")
+[×78] L2!V1119:V1196: =IF($B{N}<>"",(Lin+$T{N}-...-ABS(TURN1_R_COR))*IF(No_Turn<>0,1,0.5),"")
```

## Summary Mode (`-s`)

Appends a per-sheet digest after the diff:

```
─── Change Summary ───────────────────────────────────

  BOM
    19 formula(s) modified
      19x  $C$52  →  $C$62   (F3:F22)
──────────────────────────────────────────────────────
```
