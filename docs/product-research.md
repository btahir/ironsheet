# Product Research

Ironsheet should not compete as another object-to-XLSX builder. The valuable wedge is guarded workbook surgery: editing real Excel-authored XLSX/XLSM templates without breaking the workbook structures that users care about.

## Market Signals

- ExcelJS is a broad workbook manager, but pivot-table support has been experimental and limited. The project discussion for pivot tables notes no read support for existing pivot tables, one value field with sum, and one pivot table per document. Source: [ExcelJS pivot table discussion](https://github.com/exceljs/exceljs/discussions/2575).
- SheetJS Community Edition has public pain around preserving workbook features during round-trip edits. A 2024 style issue says styles were not preserved when reading, modifying, and re-exporting; the maintainer response points users to a Pro Edit workflow for modifying existing files. Source: [SheetJS style preservation issue](https://git.sheetjs.com/sheetjs/sheetjs/issues/3214).
- SheetJS macro support requires reading with `bookVBA` and carrying `vbaraw` into the output. That is workable, but it is an opt-in blob workflow rather than default XLSM package preservation. Source: [SheetJS VBA docs](https://docs.sheetjs.com/docs/csf/features/vba/).
- SheetJS chart round-tripping has long-standing complexity. In a chart round-trip issue, the maintainer points to a Pro Edit build that surgically edits underlying data blocks to preserve charts, and notes that keeping up with new chart types is difficult. Source: [SheetJS chart round-trip issue](https://git.sheetjs.com/sheetjs/sheetjs/issues/111).
- xlsx-populate positions itself around keeping existing features and styles intact, but open issues still show Excel repair/corruption cases, style loss, cell type drift, large-sheet export failures, and image gaps. Sources: [xlsx-populate README](https://github.com/dtjohnson/xlsx-populate), [xlsx-populate style corruption issue](https://github.com/dtjohnson/xlsx-populate/issues/157), [xlsx-populate open issues](https://github.com/dtjohnson/xlsx-populate/issues).
- Newer TypeScript projects are appearing. ExcelForge claims zero dependencies, XLSX/XLSM support, patch-only writes, and broad OOXML feature coverage; hucre is positioned around zero-dependency, ESM-native, edge-friendly spreadsheet processing. Sources: [ExcelForge README](https://github.com/node-projects/excelForge), [hucre listing](https://toolhunter.cc/tools/hucre).

## Positioning

Promise:

> Move fast and break no spreadsheets.
>
> The lossless TypeScript engine for editing real XLSX and XLSM files without breaking formulas, styles, charts, pivots, macros, or layout.

What that means in product terms:

- Preserve unknown parts by default.
- Make every mutation explicit and narrow.
- Validate before writing.
- Show a package diff that distinguishes content changes from ZIP repacking.
- Refuse ambiguous edits that would require row/column shifting until the engine can update formulas, drawings, merges, comments, tables, validations, and dependent ranges correctly.
- Treat XLSM and macro preservation as a first-class safety requirement.

## Differentiation

Against workbook builders:

- They are best when JavaScript owns the workbook shape.
- Ironsheet is best when Excel owns the workbook shape and JavaScript fills or patches it.

Against generic parsers:

- They normalize the workbook into a JS model, then reserialize.
- Ironsheet preserves untouched ZIP entries and OOXML parts as raw package payloads.

Against broad feature-complete claims:

- Ironsheet should make narrow claims, prove them with validators and fixtures, and expose unsupported structures as diagnostics or explicit errors.

Against commercial-only preservation workflows:

- The open-source wedge is a safe, typed, auditable subset: template manifests, named anchors, safe writes, package diffs, validation, XLSM preservation, and real fixture corpus testing.

## High-Value Product Requirements

Release-critical:

- Real-world XLSX/XLSM fixture corpus with charts, pivots, slicers, macros, images, named ranges, formulas, tables, styles, conditional formats, comments, hyperlinks, and protected sheets.
- Safe-by-default CLI and Node mutation paths.
- Template manifest and template render APIs that preflight all anchors before writing.
- Recalculation invalidation for formulas, direct references, named ranges, and structured table references.
- Table mutation guardrails that reject unsafe growth/adjacent overwrites.
- Validation that returns structured issues for malformed workbook XML.
- Audit diffs that classify `changed` vs `repacked`.
- Package mutation invariant tests for common safe-write paths.
- Generated cross-feature torture fixture in the automatic compatibility corpus.
- Executable README/API example coverage through public package import paths.

Next-tier:

- Browser compression adapter package with browser-targeted bundler smoke coverage.
- New image insertion with controlled drawing anchors, content-type updates, sizing helpers, and anchor metadata.
- Pivot-cache refresh diagnostics and explicit refresh metadata.
- Chart cache diagnostics and targeted chart-series retargeting.
- Style budget reporting near Excel's style limits.
- ZIP64 read/write for very large workbooks.
- Open XML SDK validation in CI through a .NET-compatible runner.

Non-goals for the first paid wedge:

- Full formula calculation engine.
- Full chart authoring UI.
- Full pivot table authoring.
- General row/column insertion that updates every dependent OOXML structure.
- Macro editing or VBA code generation.

## Developer UX Principles

- API names should say the workbook-risk level: `patchNamedRange`, `replaceTableRows`, `renderWorkbookTemplateSafely`, `mutateWorkbookFile`.
- CLI mutation output should be JSON reports, not only human strings.
- Errors should name the workbook structure that blocked the edit.
- Advanced APIs should be explicit about what they preserve, retarget, or invalidate.
- Docs should show real template workflows first, not generated toy workbooks.
- Every claim in README should map to a test, validator, fixture, or known roadmap item.

## Current Gaps To Close

- Add real fixtures and promote them from pending to active, then require `npm run compat:corpus:strict` before release.
- Keep expanding table collision tests for tables below the edited table, filters, and calculated columns.
- Expand chart, pivot, and image fixture coverage from Excel-authored workbooks.
- Add ZIP64 writing and streaming package IO before claiming true very-large-workbook support.
- Add release changelog/version automation on top of the current package metadata and dependency-ordered npm dry-run preflight.
