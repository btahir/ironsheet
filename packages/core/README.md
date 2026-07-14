<p align="center">
  <a href="https://github.com/btahir/ironsheet">
    <img src="https://raw.githubusercontent.com/btahir/ironsheet/main/docs/assets/brand/ironsheet-opengraph.png" alt="Ironsheet — lossless XLSX and XLSM editing for TypeScript" width="640">
  </a>
</p>

<h1 align="center">@ironsheet/core</h1>

<p align="center"><strong>Move fast and break no spreadsheets.</strong></p>

<p align="center">
  <a href="https://www.npmjs.com/package/@ironsheet/core"><img alt="npm version" src="https://img.shields.io/npm/v/%40ironsheet%2Fcore?style=flat-square&label=npm"></a>
</p>

The runtime-neutral workbook engine at the heart of [Ironsheet](https://github.com/btahir/ironsheet) — the lossless TypeScript engine for editing real XLSX and XLSM files without breaking formulas, styles, charts, pivots, macros, or layout.

This package is **dependency-free and browser-compatible**. It contains the ZIP/OPC/XML primitives, validators, and lossless mutation APIs. Runtime-specific file IO and compression live in the adapter packages:

- [`@ironsheet/node`](https://www.npmjs.com/package/@ironsheet/node) — Node filesystem IO, zlib, safe writes, template rendering.
- [`@ironsheet/browser`](https://www.npmjs.com/package/@ironsheet/browser) — `Blob`/`File`/`ArrayBuffer` and `CompressionStream` adapters.

Most applications should install an adapter rather than this package directly.

## Install

```bash
npm install @ironsheet/core
```

## What it does

- Parses and writes OOXML ZIP packages with raw compressed-payload preservation.
- Preserves unknown workbook XML, styles, drawings, charts, pivots, comments, merged cells, defined names, validations, conditional formats, images, and XLSM macro parts.
- Patches cells, ranges, named ranges, tables, images, styles, and more — with Excel-equivalent reference rewriting on row/sheet edits.
- Diffs two workbooks semantically and validates the OOXML package before a write.

## Documentation

Full docs — guides, recipes, and the generated API reference — are at [ironsheetdocs.vercel.app](https://ironsheetdocs.vercel.app). Source and issues live in the [Ironsheet monorepo](https://github.com/btahir/ironsheet).

## License

[Apache-2.0](./LICENSE)
