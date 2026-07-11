# Ironsheet Examples

Small, runnable examples for the core Ironsheet workflows. Each example is self-contained: it generates its own input workbook programmatically (there are no committed `.xlsx` fixtures in this repo), mutates it, and prints the safe-write report.

## Running

From the repository root, after `npm ci`:

```bash
npx tsx examples/01-safe-cell-patch.ts
npx tsx examples/02-render-template.ts
npx tsx examples/03-table-replace.ts
npx tsx examples/04-inspect-and-validate.ts
```

Generated workbooks are written to `examples/output/` (gitignored). Open them in Excel, Numbers, or LibreOffice to inspect the results.

## Examples

| Example | Shows |
| --- | --- |
| `01-safe-cell-patch.ts` | `mutateWorkbookFile`: patch cells, get validation plus a package diff, and only write output when validation passes. |
| `02-render-template.ts` | `preflightWorkbookTemplate` and `renderWorkbookTemplateSafely`: fill named ranges and tables in an Excel-authored template transactionally. |
| `03-table-replace.ts` | `replaceTableRows`: replace Excel table body rows while the table range, filter range, and worksheet cells are resized together. |
| `04-inspect-and-validate.ts` | `readWorkbook`, `inspect`, and `validateWorkbookFile`: read-only inspection and semantic validation. |

## Notes

- Examples import `@ironsheet/node` and `@ironsheet/core` through the npm workspaces; `tsx` resolves them via the repo `tsconfig.json` paths, so no build step is required.
- `examples/helpers/sample-workbook.ts` builds a tiny valid XLSX package in memory using `@ironsheet/core` ZIP primitives. In real usage, your input workbook comes from Excel — see `npm run templates:build` for richer generated starter templates under `templates/generated/`.
- The full API reference lives in `docs/api.md`.
