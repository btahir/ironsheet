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

test("validation reports orphan relationship parts", async () => {
  const pkg = await openPackage(
    await createMinimalWorkbook({ includeOrphanRelationshipPart: true })
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "RELATIONSHIP_PART_ORPHAN");
  assert.equal(report.issues[0]?.target, "xl/worksheets/missing.xml");
});

test("validation reports orphan content type overrides", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const xml = await pkg.readText("[Content_Types].xml");
  pkg.setText(
    "[Content_Types].xml",
    xml.replace(
      "</Types>",
      '<Override PartName="/xl/missing.xml" ContentType="application/xml"/></Types>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "CONTENT_TYPE_OVERRIDE_ORPHAN");
  assert.equal(report.issues[0]?.target, "xl/missing.xml");
});

test("validation reports workbook sheet relationship gaps", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const workbookXml = await pkg.readText("xl/workbook.xml");
  pkg.setText("xl/workbook.xml", workbookXml.replace('r:id="rId1"', 'r:id="rIdMissing"'));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "WORKBOOK_SHEET_RELATIONSHIP_MISSING");
  assert.equal(report.issues[0]?.target, "rIdMissing");
});

test("validation reports workbook sheets pointing to non-worksheet relationships", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const workbookXml = await pkg.readText("xl/workbook.xml");
  pkg.setText("xl/workbook.xml", workbookXml.replace('r:id="rId1"', 'r:id="rId2"'));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "WORKBOOK_SHEET_RELATIONSHIP_INVALID");
  assert.equal(report.issues[0]?.target, "rId2");
});

test("validation reports duplicate workbook sheet names", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeHiddenSheet: true }));
  const workbookXml = await pkg.readText("xl/workbook.xml");
  pkg.setText("xl/workbook.xml", workbookXml.replace('name="HiddenData"', 'name="Sheet1"'));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "WORKBOOK_SHEET_NAME_DUPLICATE");
  assert.equal(report.issues[0]?.target, "Sheet1");
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

test("validation reports worksheet drawing relationship id gaps", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeDrawing: true }));
  const xml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText("xl/worksheets/sheet1.xml", xml.replace('r:id="rIdDrawing1"', 'r:id="rIdMissing"'));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "DRAWING_RELATIONSHIP_MISSING");
  assert.equal(report.issues[0]?.target, "rIdMissing");
});

test("validation reports worksheet table part count and relationship id gaps", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeTable: true }));
  const xml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    xml
      .replace('<tableParts count="1">', '<tableParts count="2">')
      .replace('r:id="rIdTable1"', 'r:id="rIdMissing"')
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.summary.warnings, 1);
  assert.equal(
    report.issues.some((issue) => issue.code === "TABLE_PART_RELATIONSHIP_MISSING"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "TABLE_PART_COUNT_MISMATCH"),
    true
  );
});

test("validation reports drawing chart and image relationship id gaps", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeDrawing: true }));
  const xml = await pkg.readText("xl/drawings/drawing1.xml");
  pkg.setText(
    "xl/drawings/drawing1.xml",
    xml
      .replace('r:id="rIdChart1"', 'r:id="rIdMissingChart"')
      .replace('r:embed="rIdImage1"', 'r:embed="rIdMissingImage"')
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 2);
  assert.equal(
    report.issues.some((issue) => issue.code === "DRAWING_CHART_RELATIONSHIP_MISSING"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "DRAWING_IMAGE_RELATIONSHIP_MISSING"),
    true
  );
});

test("validation reports defined names that reference missing sheets", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeDefinedName: true }));
  const xml = await pkg.readText("xl/workbook.xml");
  pkg.setText("xl/workbook.xml", xml.replace("Sheet1!$A$1:$B$2", "'Missing Sheet'!$A$1:$B$2"));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "DEFINED_NAME_SHEET_MISSING");
  assert.equal(report.issues[0]?.target, "RevenueRange");
});

test("validation reports defined names with invalid local sheet ids", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeDefinedName: true }));
  const workbookXml = await pkg.readText("xl/workbook.xml");
  pkg.setText("xl/workbook.xml", workbookXml.replace('localSheetId="0"', 'localSheetId="9"'));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "DEFINED_NAME_LOCAL_SHEET_MISSING");
  assert.equal(report.issues[0]?.target, "_xlnm.Print_Titles");
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
