# Starter Templates

This directory defines generated starter workbooks for demos, examples, and regression workflows.

Build them locally:

```bash
npm run templates:build
```

The generated files are written to `templates/generated/` and are ignored by git:

- `styled-report-template.xlsx`
- `macro-model-template.xlsm`
- `dashboard-template.xlsx`
- `pivot-source-template.xlsx`
- `large-export-template.xlsx`

These templates are useful for development and public demos. They do not replace the real Excel-authored compatibility corpus in `fixtures/corpus/workbooks/`, because generated files cannot prove behavior against Excel's own OOXML serialization quirks.
