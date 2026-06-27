# Excel Exporter

A lightweight Node.js CLI that exports a single worksheet from an Excel workbook into its own `.xlsx` file.

The tool is designed for cases where you want to share one sheet without keeping live links to other sheets.

## Features

- Export exactly one worksheet by name.
- Keep workbook layout details (styles, merges, widths, validations, conditional formatting).
- Choose how formulas are handled:
  - `--break-links` (default): keep same-sheet formulas, replace cross-sheet formulas with last computed values.
  - `--values-only`: replace all formulas with last computed values.

## Requirements

- Node.js
- npm (or yarn)

## Installation

1. Clone this repository.
2. Install dependencies:

```bash
npm install
```

## Usage

```bash
node excel-exporter.js <input.xlsx> "<Sheet Name>" [outputDir] [--break-links | --values-only]
```

### Arguments

- `<input.xlsx>`: Path to the source workbook.
- `"<Sheet Name>"`: Worksheet name to export (wrap in quotes if it includes spaces).
- `[outputDir]`: Optional output folder. Defaults to the current directory (`.`).

### Modes

- `--break-links` (default)
  - Keeps formulas that only reference cells in the exported sheet.
  - Replaces formulas that reference other sheets with their cached result.
- `--values-only`
  - Replaces every formula with its cached result.

## Examples

Export a sheet to the current directory using the default mode:

```bash
node excel-exporter.js "./Project Plan.xlsx" "Wave-1"
```

Export a sheet to `./exports` and strip all formulas:

```bash
node excel-exporter.js "./Project Plan.xlsx" "Wave-1" "./exports" --values-only
```

## Output

- The output file is named after the worksheet: `<Sheet Name>.xlsx`.
- The file is written to `[outputDir]` (or `.` if omitted).

Example output log:

```text
Exported "Wave-1" -> ./exports/Wave-1.xlsx
```

## What Gets Preserved

- Worksheet properties and views
- Column widths and row heights
- Merged cell ranges
- Cell values and styles
- Data validations
- Conditional formatting

## Important Notes

- Formula results come from cached values already stored in the workbook.
- This tool does not recalculate formulas.
- If a formula has no cached result, the exported value may be empty (`null`).

## Error Handling

The script exits with an error if:

- Required arguments are missing.
- Unknown flags are provided.
- Both `--break-links` and `--values-only` are provided together.
- The requested sheet is not found in the workbook.

## Dependencies

- [exceljs](https://www.npmjs.com/package/exceljs)
