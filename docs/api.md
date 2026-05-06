# Ironsheet API

Ironsheet is a preservation-first XLSX/XLSM engine. The stable path is: open an existing workbook, make a narrow mutation, validate the result, inspect the package diff, then write only if the workbook still passes validation.

## Packages

- `@ironsheet/core`: runtime-neutral ZIP, OPC, XML, OOXML workbook model, validators, streaming XML helpers, and lossless mutation primitives.
- `@ironsheet/node`: Node file IO, zlib compression adapter, and safe-write helpers.
- `@ironsheet/cli`: command-line inspection, validation, template rendering, and safe mutations.
- `@ironsheet/compat`: compatibility report and fixture manifest utilities.

`@ironsheet/core` should stay dependency-free and browser-compatible. Node-only APIs belong in `@ironsheet/node`.

## Recommended Node Flow

```ts
import { mutateWorkbookFile } from "@ironsheet/node";

const report = await mutateWorkbookFile("template.xlsm", "output.xlsm", async (workbook) => {
  await workbook.patchNamedRange("RevenueRange", [
    ["Region", "Amount"],
    ["North", 42000]
  ]);

  await workbook.replaceTableRows("RevenueTable", [
    ["North", 42000],
    ["South", 31500]
  ]);
});

if (!report.wrote) {
  throw new Error(`Workbook failed validation with ${report.validation.summary.errors} error(s)`);
}

console.log(report.diff.summary);
```

`mutateWorkbookFile` returns:

```ts
type WorkbookSafeWriteReport = {
  diagnostics: Diagnostic[];
  diff: PackageDiff;
  validation: ValidationReport;
  wrote: boolean;
};
```

Safe writes suppress output when validation errors are present unless `allowValidationErrors` is explicitly enabled.

## Template Rendering

Use template rendering when your workbook is authored in Excel and JavaScript only fills named anchors.

```ts
import { renderWorkbookTemplateSafely } from "@ironsheet/node";

const report = await renderWorkbookTemplateSafely("template.xlsx", "report.xlsx", {
  cells: [{ sheetName: "Summary", address: "B2", value: "Q1" }],
  ranges: [{ sheetName: "Summary", startAddress: "A5", values: [["Name", "Amount"]] }],
  names: [{ name: "RevenueRange", values: [["North", 42000]] }],
  tables: [{ tableName: "RevenueTable", rows: [["North", 42000]] }],
  images: [{ imagePartName: "xl/media/image1.png", data: logoPngBytes }]
});
```

Template rendering preflights every target before applying any mutation. If a later table, named range, cell, or image target is missing, earlier changes are not applied.

## Core Workbook Methods

Stable inspection:

- `workbook.inspect()`
- `workbook.validate()`
- `workbook.diagnostics()`
- `workbook.sheets()`
- `workbook.tables()`
- `workbook.definedNames()`
- `workbook.namedRanges(name?)`
- `workbook.formulas()`
- `workbook.images(sheetName?)`
- `workbook.styles()`

Stable read APIs:

- `workbook.readCell(sheetName, address)`
- `workbook.readRange(sheetName, ref)`
- `workbook.readNamedRange(name, options?)`

Stable mutation APIs:

- `workbook.patchCell(sheetName, address, value)`
- `workbook.patchCells(sheetName, patches)`
- `workbook.patchRange(sheetName, startAddress, values)`
- `workbook.patchNamedRange(name, values, options?)`
- `workbook.appendRows(sheetName, rows, options?)`
- `workbook.replaceTableRows(tableName, rows)`
- `workbook.appendTableColumn(tableName, columnName, values?)`
- `workbook.removeRightmostTableColumn(tableName, columnName)`
- `workbook.renameSheet(sheetName, nextName)`
- `workbook.renameTable(tableName, nextName)`
- `workbook.renameTableColumn(tableName, columnName, nextName)`
- `workbook.replaceImage(imagePartName, data)`
- `workbook.setDefinedName(name, text, options?)`
- `workbook.deleteDefinedName(name, options?)`

Advanced preservation APIs:

- `workbook.retargetChartFormulas(retargets)`
- `workbook.retargetPivotCacheSources(retargets)`
- `workbook.setAutoFilter(sheetName, autoFilter)`
- `workbook.setConditionalFormat(sheetName, conditionalFormat)`
- `workbook.setDataValidation(sheetName, dataValidation)`
- `workbook.setHyperlink(sheetName, ref, target, options?)`
- `workbook.mergeCells(sheetName, ref)`
- `workbook.styleCell(sheetName, address, style)`

These advanced APIs are intentionally narrow. They should preserve unknown XML, update only the targeted structure, and emit diagnostics when adjacent workbook features may require review.

## Cell Values

```ts
type CellInput =
  | string
  | number
  | boolean
  | Date
  | null
  | { formula: string; result?: string | number | boolean | Date | null };
```

Formula edits mark the workbook for recalculation and remove stale `xl/calcChain.xml`. Named-range and table mutations also force recalculation because formulas may reference defined names or structured references that are not direct cell references.

## Package Diffs

`diffZipPackages(before, after)` classifies entries as:

- `added`: entry exists only after.
- `removed`: entry exists only before.
- `changed`: uncompressed content changed.
- `repacked`: uncompressed content appears unchanged, but ZIP compression/container bytes changed.
- `unchanged`: content and container bytes match.

Safe reports include `contentChanged` and `containerChanged` booleans per entry. This lets callers distinguish a real workbook edit from compression noise.

## CLI Contract

Read-only commands print JSON:

```bash
npm run cli -- inspect workbook.xlsx
npm run cli -- validate workbook.xlsx
npm run cli -- template-manifest workbook.xlsx
npm run cli -- diff before.xlsx after.xlsx
```

Mutating commands use safe writes by default and print a `WorkbookSafeWriteReport`:

```bash
npm run cli -- patch input.xlsx output.xlsx Sheet1 B2 "Hello"
npm run cli -- patch-named-range input.xlsx output.xlsx RevenueRange '[["North",42000]]'
npm run cli -- replace-table input.xlsx output.xlsx RevenueTable '[["North",42000]]'
npm run cli -- replace-image input.xlsx output.xlsx xl/media/image1.png logo.png
```

If validation fails, the command exits nonzero and does not write the output file.

## Failure Rules

Ironsheet should fail clearly instead of silently damaging workbooks:

- Table expansion refuses to grow through occupied worksheet rows.
- Table column append refuses to overwrite occupied cells or adjacent table ranges.
- Existing image replacement validates bytes against the current part extension.
- Invalid worksheet dimensions and cell refs are reported as validation issues instead of throwing.
- XLSM macro parts are preserved byte-for-byte unless explicitly edited by a future macro API.

## Browser/Core Usage

Core can open packages with any compression adapter:

```ts
import { OoxmlPackage, Workbook, type CompressionAdapter } from "@ironsheet/core";

const compression: CompressionAdapter = {
  inflateRaw: inflateRawBrowser,
  deflateRaw: deflateRawBrowser
};

const pkg = OoxmlPackage.open(xlsxBytes, compression);
const workbook = await Workbook.fromPackage(pkg);
```

`inflateRawBrowser` and `deflateRawBrowser` should wrap browser compression primitives such as `CompressionStream` and `DecompressionStream` and return `Uint8Array`. The core package only requires that adapter boundary; it does not import Node `zlib`.

## Compatibility Gates

Run the fast local gate:

```bash
npm run verify
```

Run the full generated fixture corpus:

```bash
npm run ci
```

Real workbook fixtures are the next release gate. Add cleared XLSX/XLSM templates under `fixtures/corpus/workbooks/`, mark them active in `fixtures/corpus/manifest.json`, and require them to pass `ironsheet-validation` plus any available app validators.
