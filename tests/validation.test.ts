import assert from "node:assert/strict";
import test from "node:test";
import { validateWorkbookPackage } from "../packages/core/src/index.ts";
import { openPackage, openWorkbook } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("validates a minimal workbook without issues", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  assert.deepEqual(await workbook.validate(), {
    issues: [],
    summary: {
      errors: 0,
      warnings: 0,
      infos: 0
    }
  });
});

test("validation reports missing relationship targets", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeDrawing: true }));
  pkg.deletePart("xl/charts/chart1.xml");

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(
    report.issues.some(
      (issue) =>
        issue.code === "RELATIONSHIP_TARGET_MISSING" && issue.target === "xl/charts/chart1.xml"
    ),
    true
  );
});

test("validation reports worksheet dimensions that exclude cells", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A1"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Original</t></is></c></row>
    <row r="2"><c r="B2"><v>42</v></c></row>
  </sheetData>
</worksheet>`
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "WORKSHEET_DIMENSION_EXCLUDES_CELL");
  assert.equal(report.issues[0]?.target, "B2");
});

test("validation reports table and autoFilter range mismatches", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeTable: true }));
  pkg.setText(
    "xl/tables/table1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="RevenueTable" displayName="RevenueTable" ref="A1:B2" totalsRowShown="0">
  <autoFilter ref="A1:B99"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Amount"/>
  </tableColumns>
</table>`
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "TABLE_AUTOFILTER_REF_MISMATCH");
});
