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
- One-cell worksheet patching.
- Formula writes with `calcPr` recalculation metadata.
- Node adapter using `node:zlib`.
- CLI inspection and patch commands.
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

Read a cell:

```bash
npm run cli -- read-cell path/to/workbook.xlsx Sheet1 A1
```

Patch a single cell:

```bash
npm run cli -- patch input.xlsx output.xlsx Sheet1 B2 "Hello from Ironsheet"
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
