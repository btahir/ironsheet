# Compatibility Testing

Unit tests are not enough for Ironsheet. The product promise depends on real workbook fidelity, so we need layered validation:

1. Unit tests for ZIP, OPC, XML, formulas, styles, and mutation planning.
2. Fixture round-trip tests for known workbook structures.
3. Package and semantic diffs for every generated workbook.
4. Optional real-application smoke checks with Numbers, LibreOffice, Excel, and Open XML SDK.

## Local Status

On this Mac, Numbers is installed. Microsoft Excel, LibreOffice, and the .NET SDK are not currently installed.

That means we can use Numbers for interactive import smoke checks now, and add stronger validators later.

## Commands

Run the fast repo gate:

```bash
npm run verify
```

Run the full local gate, including generated compatibility fixtures:

```bash
npm run ci
```

Run compatibility checks against a workbook:

```bash
npm run compat:check -- path/to/workbook.xlsx
```

This writes a JSON report under `compat-output/`.

Intake a cleared real workbook into the default corpus and activate its manifest entry:

```bash
npm run compat:intake -- styled-table-report path/to/styled-table-report.xlsx --activate
```

Use `--require=file,zip,ironsheet,openxml-sdk` when a release fixture must require additional validators. Intake copies the source workbook to the path declared in `fixtures/corpus/manifest.json`, runs compatibility checks, writes a report under `compat-output/`, and updates the manifest only after required validators pass.

Build generated smoke fixtures and run the fixture corpus:

```bash
npm run compat:corpus
```

Build demo/regression starter templates:

```bash
npm run templates:build
```

Run the corpus completeness gate:

```bash
npm run compat:corpus:strict
```

Strict corpus mode fails if any manifest fixture is still pending. It tracks the post-0.1 real-workbook coverage milestone; the `0.1.x` release preflight requires all active fixtures to pass without requiring every planned fixture to be present.

The default manifest is `fixtures/corpus/manifest.json`. It contains generated active smoke fixtures, a generated cross-feature torture workbook, and pending fixture slots for the real-world workbook shapes we still need to cover. `npm run compat:corpus` builds the generated fixtures first, then validates the manifest.

The generated torture fixture intentionally combines macros, tables, totals rows, styles, charts, images, comments, hyperlinks, pivots, hidden sheets, validations, conditional formats, and defined names. It is still generated code, not a substitute for Excel-authored files, but it keeps broad package interactions under automatic test coverage.

Generated fixtures are ignored under `fixtures/corpus/workbooks/generated/`. Add cleared real workbooks with `npm run compat:intake -- <fixture-id> <path> --activate`, then rerun the corpus command. Active fixtures fail if the workbook is missing, has failing compatibility checks, or does not pass its required validators.

Starter templates are ignored under `templates/generated/`. They are useful for demos and quick smoke checks. Real Excel-authored corpus files remain an important post-0.1 compatibility milestone.

See `fixtures/corpus/workbooks/README.md` for the intake checklist.

## Numbers Smoke Checks

By default, Numbers checks are reported as `manual` so CI and normal tests do not open desktop apps.

To open a workbook in Numbers:

```bash
IRONSHEET_RUN_NUMBERS=1 npm run compat:check -- path/to/workbook.xlsx
```

After the workbook opens, use manual inspection or Computer Use to check whether Numbers imported the workbook without an obvious error dialog.

Numbers is not an Excel substitute. Passing Numbers does not prove Excel compatibility, but it catches some structural failures and gives us a local visual smoke check.

## LibreOffice

When LibreOffice is installed, this command can run a headless import/export check:

```bash
IRONSHEET_RUN_LIBREOFFICE=1 npm run compat:check -- path/to/workbook.xlsx
```

LibreOffice is useful for automated compatibility checks, but Excel remains the final compatibility target.

## Excel

For serious release validation, use at least one environment with Microsoft Excel installed.

Future options:

- Local Mac with Excel installed.
- Dedicated Windows VM with Excel.
- Manual release checklist.
- Computer Use automation for visual smoke checks.
- Enterprise CI runner with Office installed.

## Open XML SDK

Open XML SDK validation can catch schema-level issues. The compatibility runner now looks for `tools/openxml-validator/OpenXmlValidator.csproj`; when both that project and the .NET SDK are available, set `IRONSHEET_RUN_OPENXML_SDK=1` to run it against a workbook.

When the .NET SDK is unavailable, the Open XML SDK check is skipped instead of blocking local CI.

## Release Preflight

Run the package release preflight before publishing:

```bash
npm login
npm run release:check:strict
```

The strict release preflight runs repository verification, rejects failing active corpus fixtures, checks publishable metadata and npm authentication, reports available external validators, and performs dependency-ordered workspace publish dry-runs. Planned pending fixtures do not block `0.1.x`. Publishing remains a local, manual operation and does not request npm provenance.
