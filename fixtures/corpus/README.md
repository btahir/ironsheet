# Compatibility Corpus

This directory is for real-world XLSX/XLSM files that exercise Ironsheet's compatibility promise.

The manifest starts with pending fixture slots. To activate one:

1. Add the workbook under `fixtures/corpus/workbooks/`.
2. Change its `status` in `manifest.json` from `pending` to `active`.
3. Run `npm run compat:corpus`.

Keep proprietary or customer files out of git unless they are explicitly cleared for repository use. For private workbooks, keep the manifest entry pending here and run a private manifest path locally:

```bash
npm run compat:corpus -- /path/to/private-manifest.json
```
