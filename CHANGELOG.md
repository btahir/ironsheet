# Changelog

## Unreleased

- No unreleased changes yet.

## 0.2.0 - 2026-07-29

- Browser archive preflight with compressed, uncompressed, entry-count, per-entry, worksheet-size, and compression-ratio limits before workbook inflation.
- Validation-gated browser writes through `writeWorkbookToBlobSafely`, matching Ironsheet's refuse-invalid-output principle without requiring filesystem access.
- Linear table-row removal instead of repeatedly rebuilding the full worksheet XML for every removed row.
- Large table dimension recalculation without spread-argument overflows; covered by a 140,000-cell regression fixture.
- Browser compression streams now drain readable output while writing, preventing large-entry backpressure deadlocks.
- A stateless browser demo for updating an Excel table from CSV, with Web Worker processing, local-only downloads, explicit column mapping, samples, and public implementation docs.

## 0.1.0 - 2026-07-12

- Style authoring: fonts, fills, borders, alignment, and number formats with record deduplication, `styleCell` merging into existing cell styles, and `styleRange` for range styling without style explosion.
- Structural row edits: `insertRows`/`deleteRows` with Excel-equivalent reference rewriting across formulas, defined names, merges, hyperlinks, validations, conditional formats, comment anchors, tables below the edit, and `#REF!` for dead references; refusals for table overlaps and shared-formula orphaning.
- Sheet lifecycle: `addSheet`, `copySheet`, and `deleteSheet` with cascade part cleanup, scoped defined-name reindexing, and formula breaking for deleted sheets.
- `clearCell`/`clearRange` with optional style preservation and formula recalc invalidation.
- Semantic workbook diff: `diffWorkbooks` reports cell-level adds/changes/removals plus sheet, defined-name, and table changes; `diff-cells` CLI command.
- Seeded fuzz test suite: randomized mutation sequences must keep workbooks valid across write/reopen round trips.
- Apache-2.0 license, publishable package metadata, Node 18/20/22 runtime verification, OpenXML SDK validator harness, contributor docs, issue templates, and runnable examples.
- New CLI commands: `style-range`, `clear-range`, `insert-rows`, `delete-rows`, `add-sheet`, `copy-sheet`, `delete-sheet`, `diff-cells`.
- Preservation-first XLSX/XLSM package editing with raw untouched ZIP entry preservation.
- Template manifest, template preflight, and safe template rendering APIs.
- Named-range, range, table, image, sheet, style, validation, hyperlink, merge, chart, and pivot-cache inspection/mutation helpers.
- Safe Node and CLI mutation reports with diagnostics, validation, and package diffs.
- Browser adapter with browser-targeted bundle smoke coverage.
- Generated compatibility fixtures, a cross-feature torture fixture, starter templates, and real workbook fixture intake workflow.
- Package mutation invariant tests and executable README/API example coverage.
- Worksheet child element order validation.
- Optional Numbers, LibreOffice, Open XML SDK, and Excel compatibility checks.
