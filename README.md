# xldiff (FormDiff)

Git-like line-by-line formula diffing for Excel workbooks.

Extracts raw formula definitions from two `.xlsx` files, serializes them into a sorted list, and produces a unified diff highlighting additions, deletions, and modifications — just like `git diff`, but for spreadsheet formulas.

## Features

- **Formula-only diffing** — ignores static values, formatting, dates, and blanks
- **Multi-sheet support** — processes every worksheet, prefixed with the sheet name
- **Memory-efficient** — uses `read_only=True` for workbooks with 100k+ cells
- **Colorized terminal output** — green for additions, red for deletions, cyan for hunk headers
- **File export** — optionally write a clean `.diff` patch file (no ANSI codes)
- **Array formula handling** — strips `{=...}` wrapper braces for canonical comparisons

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
| `-o`, `--output FILE` | Write the diff to a file instead of stdout |
| `-C`, `--context N` | Number of context lines around each change (default: 3) |

### Examples

```bash
# Print diff to terminal
python formdiff.py budget_q1.xlsx budget_q2.xlsx

# Save diff to a patch file
python formdiff.py old_model.xlsx new_model.xlsx -o changes.diff

# Show more context around changes
python formdiff.py base.xlsx updated.xlsx -C 5
```

## Output Format

Each formula is serialized as:

```
SheetName!CellRef: =FORMULA
```

The diff output follows standard unified diff format:

```diff
--- budget_q1.xlsx
+++ budget_q2.xlsx
@@ -10,3 +10,4 @@
 Sheet1!B2: =SUM(A1:A10)
-Sheet1!B3: =AVERAGE(A1:A10)
+Sheet1!B3: =AVERAGE(A1:A20)
+Sheet1!B4: =MAX(A1:A20)
```
