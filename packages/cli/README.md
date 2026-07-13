<p align="center">
  <a href="https://github.com/btahir/ironsheet">
    <img src="https://raw.githubusercontent.com/btahir/ironsheet/main/docs/assets/brand/ironsheet-opengraph.png" alt="Ironsheet — lossless XLSX and XLSM editing for TypeScript" width="640">
  </a>
</p>

<h1 align="center">@ironsheet/cli</h1>

<p align="center"><strong>Move fast and break no spreadsheets.</strong></p>

Safe command-line workbook inspection, validation, and XLSX/XLSM edits for [Ironsheet](https://github.com/btahir/ironsheet) — the lossless TypeScript engine for editing real Excel files without breaking formulas, styles, charts, pivots, macros, or layout.

Every mutating command uses safe writes by default: it prints a JSON report and exits nonzero **without** writing the output file when validation errors are found.

## Install

```bash
npm install -g @ironsheet/cli
# or run without installing:
npx @ironsheet/cli inspect workbook.xlsx
```

The installed binary is `ironsheet`.

## Commands

```bash
ironsheet inspect workbook.xlsx
ironsheet validate workbook.xlsx
ironsheet template-manifest template.xlsx
ironsheet preflight-template template.xlsx @patch.json
ironsheet render-template-safe template.xlsx output.xlsx @patch.json
ironsheet diff before.xlsx after.xlsx
```

All commands emit JSON, making them easy to wire into CI.

## Documentation

See the [Ironsheet monorepo](https://github.com/btahir/ironsheet) for CLI contracts and the full API guide.

## License

[Apache-2.0](./LICENSE)
