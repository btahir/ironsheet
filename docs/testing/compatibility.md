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

Run compatibility checks against a workbook:

```bash
npm run compat:check -- path/to/workbook.xlsx
```

This writes a JSON report under `compat-output/`.

Run the fixture corpus:

```bash
npm run compat:corpus
```

The default manifest is `fixtures/corpus/manifest.json`. It intentionally starts with pending fixture slots for the workbook shapes we need to cover. Add cleared workbooks under `fixtures/corpus/workbooks/`, flip the matching manifest entry to `active`, and rerun the corpus command. Active fixtures fail if the workbook is missing, has failing compatibility checks, or does not pass its required validators.

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

Open XML SDK validation can catch schema-level issues. It requires the .NET SDK and a validator harness that we have not scaffolded yet.

This should become a future CI job once the workbook writer exists.
