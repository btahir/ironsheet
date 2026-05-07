# Compatibility Corpus

This directory is for real-world XLSX/XLSM files that exercise Ironsheet's compatibility promise.

The manifest starts with pending fixture slots. To activate one:

1. Run `npm run compat:intake -- <fixture-id> /path/to/cleared-workbook.xlsx --activate`.
2. Add `--require=file,zip,ironsheet,openxml-sdk` or other validators when the local release environment supports them.
3. Run `npm run compat:corpus`.

Keep proprietary or customer files out of git unless they are explicitly cleared for repository use. For private workbooks, keep the manifest entry pending here and run a private manifest path locally:

```bash
npm run compat:corpus -- /path/to/private-manifest.json
```

The default manifest also includes generated smoke fixtures. They are written under
`fixtures/corpus/workbooks/generated/` by:

```bash
npm run compat:fixtures
```

Generated fixtures are intentionally ignored by git. They make the public corpus gate active
without committing binary workbook blobs, while the pending slots remain reserved for cleared
real-world workbooks.
