# Ironsheet

Move fast and break no spreadsheets.

Ironsheet is the lossless TypeScript engine for editing real XLSX and XLSM files without breaking formulas, styles, charts, pivots, macros, or layout.

## Current MVP Slice

This repository currently implements the first vertical slice:

- ZIP central-directory parsing and writing.
- CRC32.
- Raw compressed payload preservation for untouched entries.
- Duplicate and unsafe ZIP entry path rejection on read and write.
- OPC relationship parsing and target resolution.
- Workbook sheet discovery.
- Defined-name inspection.
- Defined-name set/delete helpers for named ranges and scoped workbook names.
- Hyperlink inventory plus external hyperlink set/delete with worksheet relationship management.
- Merged-cell inventory plus guarded merge/unmerge helpers with overlap validation.
- Table metadata discovery with worksheet ownership, refs, totals-row count, and column metadata.
- Safe sheet renaming with formula, defined-name, chart, and pivot cache retargeting.
- Safe table and table-column renaming with structured-reference retargeting across formulas and defined names.
- Safe table column append and guarded rightmost-column removal.
- Style metadata inspection with cell format counts.
- Deduped cell style creation with custom number formats and style application.
- Cell, batch-cell, and range read/write APIs.
- Append-row API for template-backed exports.
- Chunked worksheet row XML reader for future large-sheet transforms.
- Chunked row XML streaming primitive for large worksheet writer work.
- Streamed worksheet row replacement for large-sheet transforms that should not materialize full worksheet XML.
- Basic table row replacement with table ref, autoFilter, style, totals-row preservation, formula recalculation, and worksheet dimension updates.
- Formula writes, formula removal, totals-row formula movement, and dependent value edits with namespace-aware `calcPr` recalculation metadata and stale calc-chain removal.
- Formula inventory plus sheet, cell/range, bounds, shared-formula metadata, and structured table-reference parsing for validation and future retargeting work.
- Feature inspection for macros, shared strings, formula cells, external relationships, tables, drawings, charts, media, merges, hyperlinks, validation, conditional formatting, hidden sheets, comments, pivots, and defined names.
- Exact chart formula retargeting and pivot cache worksheet-source retargeting.
- Preservation fixtures for macros, merge cells, hyperlinks, validations, conditional formatting, drawings, charts, media, hidden sheets, and styled table rows.
- Namespace-prefix preservation for workbook recalculation edits, worksheet cell/row insertion, append rows, and table row replacement.
- Streaming XML tokenizer, chunk transform primitives, and streamed element extraction powering local XML helpers and future large-worksheet transforms.
- Semantic validation for package relationship targets, duplicate relationship IDs, orphan relationship parts, orphan content type overrides, local worksheet/drawing relationship IDs, workbook sheet metadata, content types, worksheet dimensions, merge/validation/conditional-format/hyperlink refs, style indexes and cell format limits, shared string references/counts, worksheet/defined-name/chart formula references, formula bounds, formula table references, shared formula groups, table refs, table part counts, table metadata, table column structure, pivot table/cache source sanity, stale calc chains, and defined-name scope integrity.
- Node adapter using `node:zlib`.
- CLI inspection, read, patch, range patch, table replacement/rename, and package diff commands.
- Compatibility harness with ZIP integrity, Ironsheet semantic validation, and optional app-level checks.
- Runtime guard that automatically scans every core source file for Node-only imports and runtime dependencies.
- Workspace package manifests for core, node, compat, and CLI packages plus CI wiring.

## Commands

Run the fast quality gate:

```bash
npm run verify
```

Run the CI-equivalent local gate:

```bash
npm run ci
```

Inspect a workbook:

```bash
npm run cli -- inspect path/to/workbook.xlsx
```

List workbook tables:

```bash
npm run cli -- tables path/to/workbook.xlsx
```

List workbook formulas:

```bash
npm run cli -- formulas path/to/workbook.xlsx
```

List workbook hyperlinks:

```bash
npm run cli -- hyperlinks path/to/workbook.xlsx
```

List workbook merged cells:

```bash
npm run cli -- merged-cells path/to/workbook.xlsx
```

Inspect workbook styles:

```bash
npm run cli -- styles path/to/workbook.xlsx
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

Apply a cell style:

```bash
npm run cli -- style-cell input.xlsx output.xlsx Sheet1 B2 '{"numberFormat":"$#,##0.00"}'
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

Append or remove a table column:

```bash
npm run cli -- append-table-column input.xlsx output.xlsx RevenueTable Margin '[0.5,0.7]'
npm run cli -- remove-table-column input.xlsx output.xlsx RevenueTable Margin
```

Rename a sheet, table, or table column and retarget references:

```bash
npm run cli -- rename-sheet input.xlsx output.xlsx Sheet1 "Revenue 2026"
npm run cli -- rename-table input.xlsx output.xlsx RevenueTable SalesData
npm run cli -- rename-table-column input.xlsx output.xlsx RevenueTable Amount NetAmount
```

Set or delete a defined name:

```bash
npm run cli -- set-defined-name input.xlsx output.xlsx ReportRange 'Sheet1!$A$1:$B$10'
npm run cli -- delete-defined-name input.xlsx output.xlsx ReportRange
```

Set or delete an external hyperlink:

```bash
npm run cli -- set-hyperlink input.xlsx output.xlsx Sheet1 B2 https://example.com '{"display":"Example"}'
npm run cli -- delete-hyperlink input.xlsx output.xlsx Sheet1 B2
```

Merge or unmerge cells:

```bash
npm run cli -- merge-cells input.xlsx output.xlsx Sheet1 A1:B1
npm run cli -- unmerge-cells input.xlsx output.xlsx Sheet1 A1:B1
```

Retarget chart formulas or pivot cache sources:

```bash
npm run cli -- retarget-chart input.xlsx output.xlsx '[{"from":"Sheet1!$A$1:$B$2","to":"Sheet1!$A$1:$C$2"}]'
npm run cli -- retarget-pivot input.xlsx output.xlsx '[{"from":{"sheet":"Sheet1","ref":"A1:B2"},"to":{"sheet":"Sheet1","ref":"A1:C2"}}]'
```

Diff two workbook packages:

```bash
npm run cli -- diff input.xlsx output.xlsx
```

Run compatibility checks:

```bash
npm run compat:check -- output.xlsx
```

Run the fixture corpus compatibility matrix:

```bash
npm run compat:corpus
```

This builds ignored generated smoke workbooks, then validates the active corpus fixtures.

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

This runs the CI-equivalent local gate before committing.

`IRONSHEET_SPEC.md` is local planning material and is intentionally ignored.
