---
name: Bug report
about: Report a workbook that Ironsheet mishandles or an API that misbehaves
title: ""
labels: bug
assignees: ""
---

## What happened

A clear description of the bug. If Ironsheet damaged or rejected a workbook, describe what changed (validation errors, Excel repair prompts, lost styles, broken formulas, etc.).

## Reproduction

Ironsheet bugs are almost always workbook-specific, so a reproduction workbook matters more than anything else.

- [ ] I attached the workbook that triggers the bug, or
- [ ] I attached a redacted copy (delete sensitive values; keep the structure that triggers the bug), or
- [ ] I included code that generates a reproduction workbook programmatically

```ts
// Minimal code that reproduces the issue
```

## Expected behavior

What you expected to happen instead.

## Environment

- Ironsheet package and version (e.g. `@ironsheet/node@0.1.0`):
- Runtime (Node version, browser, or CLI):
- OS:
- Excel version that authored the workbook (or other tool, e.g. LibreOffice, Google Sheets export):

## Diagnostics

If available, paste the output of:

```bash
npm run cli -- validate workbook.xlsx
```

or the `WorkbookSafeWriteReport` (`diagnostics`, `validation.summary`, `diff.summary`) from the failing call.
