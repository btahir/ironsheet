# Ironsheet API Surface

Ironsheet is split into runtime-neutral core APIs and Node-specific IO APIs.

## Packages

- `@ironsheet/core`: ZIP, OPC, XML, workbook model, validation, streaming primitives, and lossless mutation helpers. It must not import Node-only modules or runtime dependencies.
- `@ironsheet/node`: Node file IO and zlib compression adapter.
- `@ironsheet/compat`: Compatibility report and fixture manifest types.
- `@ironsheet/cli`: Developer CLI over the Node adapter.

## Current Core Capabilities

- Inspect sheets, features, auto filters, comments, conditional formats, data validations, defined names, formulas, hyperlinks, images, merged cells, tables, and styles.
- Read and write cells, ranges, appended rows, worksheet auto filters, conditional-formatting blocks, data validations, defined names, external hyperlinks, existing image bytes, merged cells, table rows, sheet names, sheet visibility, table names, table columns, and cell styles.
- Render template patches across cells, ranges, tables, and existing images in one typed operation.
- Preserve untouched ZIP entries and OOXML parts.
- Validate relationships, dimensions, hyperlinks, merged cells, styles, shared strings, formula references, tables, pivots, charts, and calc chains.
- Retarget exact chart formulas and pivot cache worksheet sources.
- Stream row XML and replace selected worksheet rows from chunked XML input.

## Stability Rules

- Core remains browser-compatible and dependency-free.
- Mutations should either preserve unknown XML or fail with a narrow error rather than rewriting unsupported structures.
- APIs that can invalidate formulas or cached results must mark the workbook for recalculation or remove stale calc-chain parts.
- Any API that edits table, chart, pivot, style, formula, or macro-adjacent structures needs fixture coverage before being considered release-ready.

## Validation

Run the local gate:

```bash
npm run verify
```

Run the corpus gate:

```bash
npm run compat:corpus
```
