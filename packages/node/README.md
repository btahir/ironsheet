<p align="center">
  <a href="https://github.com/btahir/ironsheet">
    <img src="https://raw.githubusercontent.com/btahir/ironsheet/main/docs/assets/brand/ironsheet-opengraph.png" alt="Ironsheet — lossless XLSX and XLSM editing for TypeScript" width="640">
  </a>
</p>

<h1 align="center">@ironsheet/node</h1>

<p align="center"><strong>Move fast and break no spreadsheets.</strong></p>

<p align="center">
  <a href="https://www.npmjs.com/package/@ironsheet/node"><img alt="npm version" src="https://img.shields.io/npm/v/%40ironsheet%2Fnode?style=flat-square&label=npm"></a>
</p>

Node.js file IO, compression, and safe-write adapter for [Ironsheet](https://github.com/btahir/ironsheet) — the lossless TypeScript engine for editing real XLSX and XLSM files without breaking formulas, styles, charts, pivots, macros, or layout.

This is the package most Node users want. It bundles the runtime-neutral [`@ironsheet/core`](https://www.npmjs.com/package/@ironsheet/core) engine with filesystem IO, zlib compression, safe writes, and template render helpers.

## Install

```bash
npm install @ironsheet/node
```

## Usage

Guarded mutation of an existing workbook — Ironsheet validates before it writes and refuses to produce a corrupt file:

```ts
import { mutateWorkbookFile } from "@ironsheet/node";

const report = await mutateWorkbookFile("template.xlsx", "report.xlsx", async (workbook) => {
  await workbook.patchCell("Summary", "B2", "Q1");
  await workbook.patchNamedRange("RevenueRange", [["North", 42000]]);
  await workbook.replaceTableRows("RevenueTable", [["North", 42000]]);
});

if (!report.wrote) {
  throw new Error("Ironsheet refused to write an invalid workbook");
}

console.log(report.diff.summary);
```

Safe template rendering resolves anchors, validates the resize plan, then applies one transaction:

```ts
import { renderWorkbookTemplateSafely } from "@ironsheet/node";

const report = await renderWorkbookTemplateSafely("template.xlsm", "output.xlsm", {
  names: [{ name: "RevenueRange", values: [["Region", "Amount"], ["North", 42000]] }],
  tables: [{ tableName: "RevenueTable", rows: [["North", 42000], ["South", 31500]] }]
});
```

## Documentation

Full API guide, CLI, capabilities, and compatibility notes live in the [Ironsheet monorepo](https://github.com/btahir/ironsheet).

## License

[Apache-2.0](./LICENSE)
