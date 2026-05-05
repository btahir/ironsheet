# Ironsheet

Move fast and break no spreadsheets.

Ironsheet is the lossless TypeScript engine for editing real XLSX and XLSM files without breaking formulas, styles, charts, pivots, macros, or layout.

## Current MVP Slice

This repository currently implements the first vertical slice:

- ZIP central-directory parsing and writing.
- CRC32.
- Raw compressed payload preservation for untouched entries.
- OPC relationship parsing and target resolution.
- Workbook sheet discovery.
- Defined-name inspection.
- Cell, batch-cell, and range read/write APIs.
- Append-row API for template-backed exports.
- Chunked row XML streaming primitive for large worksheet writer work.
- Basic table row replacement with table ref, autoFilter, style, totals-row preservation, formula recalculation, and worksheet dimension updates.
- Formula writes with namespace-aware `calcPr` recalculation metadata and stale calc-chain removal.
- Feature inspection for macros, shared strings, tables, drawings, charts, media, merges, hyperlinks, validation, conditional formatting, hidden sheets, comments, pivots, and defined names.
- Preservation fixtures for macros, merge cells, hyperlinks, validations, conditional formatting, drawings, charts, media, hidden sheets, and styled table rows.
- Namespace-prefix preservation for workbook recalculation edits, worksheet cell/row insertion, append rows, and table row replacement.
- Semantic validation for package relationship targets, duplicate relationship IDs, orphan relationship parts, orphan content type overrides, local worksheet/drawing relationship IDs, workbook sheet metadata, content types, worksheet dimensions, table refs, table part counts, table column structure, stale calc chains, and defined-name sheet references.
- Node adapter using `node:zlib`.
- CLI inspection, read, patch, range patch, table replacement, and package diff commands.
- Compatibility harness with ZIP integrity and optional app-level checks.

## Commands

Run the full quality gate:

```bash
npm run verify
```

Inspect a workbook:

```bash
npm run cli -- inspect path/to/workbook.xlsx
```

Validate package integrity and common workbook invariants:

```bash
npm run cli -- validate path/to/workbook.xlsx
```

Read a cell:

```bash
npm run cli -- read-cell path/to/workbook.xlsx Sheet1 A1
```

Read a range:

```bash
npm run cli -- read-range path/to/workbook.xlsx Sheet1 A1:C5
```

Patch a single cell:

```bash
npm run cli -- patch input.xlsx output.xlsx Sheet1 B2 "Hello from Ironsheet"
```

Patch a range:

```bash
npm run cli -- patch-range input.xlsx output.xlsx Sheet1 B2 '[["Name","Amount"],["ACME",42]]'
```

Append rows:

```bash
npm run cli -- append-rows input.xlsx output.xlsx Sheet1 '[["North",10],["South",20]]'
```

Replace basic table rows:

```bash
npm run cli -- replace-table input.xlsx output.xlsx RevenueTable '[["New",10],["Growth",20]]'
```

Diff two workbook packages:

```bash
npm run cli -- diff input.xlsx output.xlsx
```

Run compatibility checks:

```bash
npm run compat:check -- output.xlsx
```

Open in Numbers for an interactive smoke check on macOS:

```bash
IRONSHEET_RUN_NUMBERS=1 npm run compat:check -- output.xlsx
```

## Development

Use TypeScript for implementation, scripts, and tests.

Use the safe commit loop:

```bash
npm run commit:safe -- "feat: add workbook capability"
```

`IRONSHEET_SPEC.md` is local planning material and is intentionally ignored.
