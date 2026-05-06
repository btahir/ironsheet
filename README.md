# Ironsheet

Move fast and break no spreadsheets.

Ironsheet is a preservation-first TypeScript engine for editing real XLSX and XLSM workbooks without rewriting the parts you did not touch. It is built for report templates, finance models, dashboards, macro-enabled workbooks, and other Excel files where styles, formulas, charts, pivots, drawings, relationships, and layout matter.

Use Ironsheet when you need to modify an existing workbook and prove what changed.

## Why Ironsheet

Most spreadsheet libraries are optimized for creating workbook-shaped files from JavaScript objects. Ironsheet is optimized for guarded workbook mutation:

- **Lossless package editing**: untouched ZIP entries are preserved as raw compressed payloads.
- **Template anchors**: patch named ranges, tables, cells, ranges, and existing images in one operation.
- **Safety reports**: validate OOXML invariants, emit diagnostics, and classify package entries as changed, repacked, added, removed, or unchanged before writing.
- **XLSM-aware preservation**: macro parts are kept byte-for-byte unless explicitly touched.
- **Narrow failure modes**: unsupported workbook structures should produce targeted errors or warnings, not silent rewrites.

That makes Ironsheet a better fit for “fill this Excel-authored template” than broad workbook builders.

## Quickstart

```bash
npm install @ironsheet/node
```

```ts
import { renderWorkbookTemplateSafely } from "@ironsheet/node";

const report = await renderWorkbookTemplateSafely("template.xlsm", "output.xlsm", {
  names: [
    {
      name: "RevenueRange",
      values: [
        ["Region", "Amount"],
        ["North", 42000]
      ]
    }
  ],
  tables: [
    {
      tableName: "RevenueTable",
      rows: [
        ["North", 42000],
        ["South", 31500]
      ]
    }
  ],
  images: [
    {
      imagePartName: "xl/media/image1.png",
      data: await fetchLogoBytes()
    }
  ]
});

if (!report.wrote) {
  throw new Error(`Workbook failed validation: ${report.validation.summary.errors} error(s)`);
}

console.log(report.diff.summary);
```

## CLI

During repo development, run commands through `npm run cli -- ...`. Once installed as a package, the binary is `ironsheet`.

Inspect a workbook before editing it:

```bash
npm run cli -- template-manifest template.xlsx
```

Render a template with validation, diagnostics, and a package diff before writing:

```bash
npm run cli -- preflight-template template.xlsx @patch.json
npm run cli -- render-template-safe template.xlsx output.xlsx @patch.json
```

Example `patch.json`:

```json
{
  "names": [
    {
      "name": "RevenueRange",
      "values": [
        ["Region", "Amount"],
        ["North", 42000]
      ]
    }
  ],
  "tables": [
    {
      "tableName": "RevenueTable",
      "rows": [
        ["North", 42000],
        ["South", 31500]
      ]
    }
  ],
  "images": [{ "imagePartName": "xl/media/image1.png", "path": "logo.png" }]
}
```

Useful inspection commands:

```bash
npm run cli -- inspect workbook.xlsx
npm run cli -- validate workbook.xlsx
npm run cli -- diff before.xlsx after.xlsx
npm run cli -- named-ranges workbook.xlsx
npm run cli -- tables workbook.xlsx
npm run cli -- images workbook.xlsx
npm run cli -- formulas workbook.xlsx
npm run cli -- styles workbook.xlsx
```

Targeted mutation commands:

```bash
npm run cli -- patch input.xlsx output.xlsx Sheet1 B2 "Hello"
npm run cli -- patch-range input.xlsx output.xlsx Sheet1 B2 '[["Name","Amount"],["ACME",42]]'
npm run cli -- patch-named-range input.xlsx output.xlsx RevenueRange '[["Name","Amount"],["ACME",42]]'
npm run cli -- replace-table input.xlsx output.xlsx RevenueTable '[["North",10],["South",20]]'
npm run cli -- replace-image input.xlsx output.xlsx xl/media/image1.png logo.png
npm run cli -- insert-image input.xlsx output.xlsx Sheet1 logo.png
npm run cli -- rename-sheet input.xlsx output.xlsx Sheet1 "Revenue 2026"
npm run cli -- rename-table input.xlsx output.xlsx RevenueTable SalesData
npm run cli -- rename-table-column input.xlsx output.xlsx RevenueTable Amount NetAmount
```

Mutating CLI commands use safe writes by default. They print JSON reports and exit nonzero without writing the output file when validation errors are found.

## What Works Today

Core workbook capabilities:

- ZIP central-directory parse/write, ZIP64 metadata read for in-memory archives, CRC32, duplicate path rejection, unsafe path rejection, and raw compressed payload preservation.
- OPC relationship parsing, target resolution, relationship mutation, and content type validation.
- Sheet discovery, visibility edits, and safe sheet renaming with formula, defined-name, chart, and pivot-cache retargeting.
- Cell, batch-cell, range, named-range, append-row, and table-row writes.
- Table metadata, table row replacement, table rename, table-column rename, rightmost column removal, and append column.
- Defined names, worksheet auto filters, data validations, conditional formats, hyperlinks, merged cells, comments, images, styles, formulas, charts, pivots, and macros inspection.
- Existing image replacement plus new image insertion with one-cell and two-cell anchors.
- Safe-by-default CLI mutation reports with validation, diagnostics, and content-vs-repack package diffs.
- Style inspection and deduped cell format creation.
- Template manifest, public template preflight, and template render APIs.
- Semantic validation for relationships, dimensions, hyperlinks, merged cells, styles, shared strings, formulas, tables, pivots, charts, calc chains, defined names, and content types.
- Compatibility corpus with generated XLSX/XLSM/dashboard/pivot/large-sheet fixtures plus optional Numbers, LibreOffice, Open XML SDK, and Excel checks.

Runtime split:

- `@ironsheet/core`: runtime-neutral workbook engine and low-level OOXML primitives.
- `@ironsheet/node`: Node file IO, compression adapter, and safe write helpers.
- `@ironsheet/browser`: browser compression adapter and Blob/ArrayBuffer helpers over the core engine.
- `@ironsheet/compat`: compatibility report and fixture manifest types.
- `@ironsheet/cli`: command-line interface over the Node adapter.

## Guardrails

- Mutations preserve unknown XML by default.
- Formula edits and dependent data edits mark workbooks for recalculation and remove stale calc-chain parts.
- Named-range and table mutations conservatively force recalculation for defined-name and structured-reference formulas.
- Template rendering preflights all targets before applying mutations.
- Table expansion refuses occupied rows, and table column append refuses adjacent occupied cells or overlapping table ranges.
- Image replacement and insertion validate bytes against the target media type.
- Invalid worksheet dimensions and cell refs are reported as validation issues instead of crashing validation.
- Safe writes validate the final workbook and suppress output when validation errors are found.
- Core runtime code has a guard against Node-only imports and runtime dependencies.

## Known Gaps

These are intentional product boundaries, not hidden claims:

- ZIP64 writing is not implemented yet; ZIP64 reading is metadata-only and still memory-backed.
- Chart and pivot support focuses on preservation, validation, and targeted retargeting, not full chart/pivot authoring.
- Some real-world Excel structures still need corpus fixtures before they should be considered release-ready.

## Docs

- [API guide](docs/api.md)
- [Product research](docs/product-research.md)
- [Compatibility testing](docs/testing/compatibility.md)

## Development

Use TypeScript for implementation, scripts, fixtures, and tests.

Run the fast gate:

```bash
npm run verify
```

Run the CI-equivalent local gate:

```bash
npm run ci
```

Build local demo/regression workbook templates:

```bash
npm run templates:build
```

Build publishable JS and declaration output under each package `dist/` directory:

```bash
npm run build
```

Run release preflight, including CI, package dry-runs, workspace metadata checks, validator capability reporting, and npm public publish dry-run with provenance:

```bash
npm run release:check
```

Use the safe commit loop:

```bash
npm run commit:safe -- "feat: add workbook capability"
```

`IRONSHEET_SPEC.md` is local planning material and is intentionally ignored.
