# Invoice Project Distribution

A small offline web app that splits one Excel workbook (with multiple sheets,
4 by default) into N Excel files. Every output file keeps the same sheet
structure as the source, and the rows of each sheet are distributed as evenly
as possible across the outputs.

## Usage

1. Open `index.html` in any modern browser. No server, no install.
2. Choose (or drag &amp; drop) your `.xlsx` file.
3. Enter the number of output files you want.
4. Tick "First row of each sheet is a header" if your sheets have headers
   (the header row is then copied into every output file).
5. Click **Split &amp; Download**. You get N files named
   `<original>_part_01.xlsx`, `<original>_part_02.xlsx`, ...

## How rows are distributed

For each sheet independently:

- Data rows are split into N contiguous chunks of nearly equal size.
- If the row count is not divisible by N, the first chunks get one extra
  row each (e.g. 10 rows into 3 files -> 4, 3, 3).

This means every output file gets approximately the same amount of work
across all 4 sheets.

## Offline

The app uses [SheetJS Community Edition](https://sheetjs.com/) for parsing
and writing `.xlsx`. The library is vendored under `vendor/xlsx.full.min.js`,
so the app runs entirely in the browser with no network access required.

## Files

- `index.html` - UI
- `styles.css` - styling
- `app.js` - parsing, splitting and download logic
- `vendor/xlsx.full.min.js` - SheetJS (offline)
