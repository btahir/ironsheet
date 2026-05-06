# Real Workbook Fixture Intake

Put cleared, non-confidential Excel-authored workbooks in this directory and activate the matching fixture in `../manifest.json`.

Required first fixtures:

- `styled-table-report.xlsx`
- `macro-enabled-model.xlsm`
- `chart-dashboard.xlsx`
- `pivot-cache-workbook.xlsx`
- `large-sheet-export.xlsx`

Intake checklist:

- Remove private customer data, credentials, links, tracked changes, and document properties.
- Open once in Excel and save from Excel, not only from another library.
- Keep the workbook representative: formulas, styles, charts, pivots, macros, hidden sheets, and layout should remain intact.
- Run `npm run compat:check -- fixtures/corpus/workbooks/<file>`.
- Change the fixture `status` from `pending` to `active` in `fixtures/corpus/manifest.json`.
- Run `npm run compat:corpus`.
- Before release, run `npm run compat:corpus:strict` so pending fixtures fail instead of being skipped.

Generated fixtures live under `generated/` and are ignored. They are useful smoke tests, but they do not replace real Excel-authored files.
