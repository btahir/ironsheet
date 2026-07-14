<p align="center">
  <a href="https://github.com/btahir/ironsheet">
    <img src="https://raw.githubusercontent.com/btahir/ironsheet/main/docs/assets/brand/ironsheet-opengraph.png" alt="Ironsheet — lossless XLSX and XLSM editing for TypeScript" width="640">
  </a>
</p>

<h1 align="center">@ironsheet/browser</h1>

<p align="center"><strong>Move fast and break no spreadsheets.</strong></p>

<p align="center">
  <a href="https://www.npmjs.com/package/@ironsheet/browser"><img alt="npm version" src="https://img.shields.io/npm/v/%40ironsheet%2Fbrowser?style=flat-square&label=npm"></a>
</p>

Browser compression and file adapters for [Ironsheet](https://github.com/btahir/ironsheet) — the lossless TypeScript engine for editing real XLSX and XLSM files without breaking formulas, styles, charts, pivots, macros, or layout.

This package adapts the runtime-neutral [`@ironsheet/core`](https://www.npmjs.com/package/@ironsheet/core) engine to the browser using `Blob`, `File`, `ArrayBuffer`, `CompressionStream`, and `DecompressionStream`.

## Install

```bash
npm install @ironsheet/browser
```

## Usage

```ts
import { openWorkbookFromBlob, writeWorkbookToBlob } from "@ironsheet/browser";

const workbook = await openWorkbookFromBlob(file);
await workbook.patchCell("Sheet1", "B2", "Hello from the browser");

const output = await writeWorkbookToBlob(workbook);
```

The same workbook engine and lossless guarantees apply as in Node — only IO and compression differ.

## Documentation

Full docs — guides, recipes, and the generated API reference — are at [ironsheetdocs.vercel.app](https://ironsheetdocs.vercel.app). Source and issues live in the [Ironsheet monorepo](https://github.com/btahir/ironsheet).

## License

[Apache-2.0](./LICENSE)
