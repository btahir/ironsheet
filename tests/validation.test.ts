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

test("validation reports duplicate relationship ids", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const relsXml = await pkg.readText("xl/_rels/workbook.xml.rels");
  pkg.setText(
    "xl/_rels/workbook.xml.rels",
    relsXml.replace(
      "</Relationships>",
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "RELATIONSHIP_ID_DUPLICATE");
  assert.equal(report.issues[0]?.target, "rId1");
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

test("validation reports invalid worksheet range attributes", async () => {
  const pkg = await openPackage(
    await createMinimalWorkbook({
      includeConditionalFormatting: true,
      includeDataValidation: true,
      includeHyperlink: true,
      includeMerge: true
    })
  );
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml
      .replace('<mergeCells count="1">', '<mergeCells count="2">')
      .replace('<mergeCell ref="A1:B1"/>', '<mergeCell ref="XFE1:XFE2"/>')
      .replace('sqref="B2:B10"', 'sqref="BAD XFE1"')
      .replace('sqref="A1:A10"', 'sqref="A1048577"')
      .replace(
        '<hyperlink ref="A1" r:id="rIdHyperlink1"/>',
        '<hyperlink ref="1A" r:id="rIdHyperlink1"/>'
      )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 5);
  assert.equal(report.summary.warnings, 1);
  assert.equal(
    report.issues.some((issue) => issue.code === "MERGE_CELL_COUNT_MISMATCH"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "MERGE_CELL_REF_OUT_OF_BOUNDS"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "DATA_VALIDATION_SQREF_INVALID"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "DATA_VALIDATION_SQREF_OUT_OF_BOUNDS"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "CONDITIONAL_FORMATTING_SQREF_OUT_OF_BOUNDS"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "HYPERLINK_REF_INVALID"),
    true
  );
});

test("validation reports worksheet hyperlinks with missing relationships", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      "</sheetData>",
      '</sheetData><hyperlinks><hyperlink ref="A1" r:id="rIdMissing"/></hyperlinks>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "HYPERLINK_RELATIONSHIP_MISSING");
  assert.equal(report.issues[0]?.target, "rIdMissing");
});

test("validation reports worksheet cells that reference missing style indexes", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText("xl/worksheets/sheet1.xml", worksheetXml.replace('s="1"', 's="99"'));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "CELL_STYLE_INDEX_MISSING");
  assert.equal(report.issues[0]?.target, "99");
});

test("validation reports style count mismatches", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const stylesXml = await pkg.readText("xl/styles.xml");
  pkg.setText("xl/styles.xml", stylesXml.replace('<cellXfs count="5">', '<cellXfs count="4">'));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "STYLE_CELLXFS_COUNT_MISMATCH");
});

test("validation reports style counts above Excel limits", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const stylesXml = await pkg.readText("xl/styles.xml");
  pkg.setText("xl/styles.xml", stylesXml.replace('<cellXfs count="5">', '<cellXfs count="65491">'));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 2);
  assert.equal(
    report.issues.some((issue) => issue.code === "STYLE_CELLXFS_COUNT_MISMATCH"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "STYLE_CELLXFS_COUNT_EXCEEDS_EXCEL_LIMIT"),
    true
  );
});

test("validation reports missing shared string indexes", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ useSharedStrings: true }));
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText("xl/worksheets/sheet1.xml", worksheetXml.replace("<v>0</v>", "<v>9</v>"));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "SHARED_STRING_INDEX_MISSING");
  assert.equal(report.issues[0]?.target, "A1");
});

test("validation reports shared string count mismatches", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ useSharedStrings: true }));
  const sharedStringsXml = await pkg.readText("xl/sharedStrings.xml");
  pkg.setText(
    "xl/sharedStrings.xml",
    sharedStringsXml.replace('count="1" uniqueCount="1"', 'count="0" uniqueCount="2"')
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 2);
  assert.equal(
    report.issues.some((issue) => issue.code === "SHARED_STRINGS_UNIQUE_COUNT_MISMATCH"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "SHARED_STRINGS_COUNT_UNDER_REPORTS_USAGE"),
    true
  );
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

test("validation accepts coherent pivot table parts", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includePivotTable: true }));

  const report = await validateWorkbookPackage(pkg);

  assert.deepEqual(report.summary, { errors: 0, warnings: 0, infos: 0 });
});

test("validation reports pivot table and cache source issues", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includePivotTable: true }));
  const pivotTableXml = await pkg.readText("xl/pivotTables/pivotTable1.xml");
  pkg.setText(
    "xl/pivotTables/pivotTable1.xml",
    pivotTableXml.replace('cacheId="1"', 'cacheId="99"')
  );
  const pivotCacheXml = await pkg.readText("xl/pivotCache/pivotCacheDefinition1.xml");
  pkg.setText(
    "xl/pivotCache/pivotCacheDefinition1.xml",
    pivotCacheXml.replace('ref="A1:B2" sheet="Sheet1"', 'ref="XFE1" sheet="Missing"')
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.summary.warnings, 2);
  assert.equal(
    report.issues.some((issue) => issue.code === "PIVOT_TABLE_CACHE_ID_UNKNOWN"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "PIVOT_CACHE_SOURCE_SHEET_MISSING"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "PIVOT_CACHE_SOURCE_REF_OUT_OF_BOUNDS"),
    true
  );
});

test("validation reports worksheet pivot relationship gaps", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includePivotTable: true }));
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace('r:id="rIdPivotTable1"', 'r:id="rIdMissingPivot"')
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "PIVOT_TABLE_RELATIONSHIP_MISSING");
  assert.equal(report.issues[0]?.target, "rIdMissingPivot");
});

test("validation reports chart formulas with broken references", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeDrawing: true }));
  pkg.setText(
    "xl/charts/chart1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat><c:strRef><c:f>Missing!$A$1:$A$2</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$XFE$1</c:f></c:numRef></c:val>
          <c:tx><c:strRef><c:f>MissingTable[Amount]</c:f></c:strRef></c:tx>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.summary.warnings, 2);
  assert.equal(
    report.issues.some((issue) => issue.code === "CHART_FORMULA_SHEET_MISSING"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "CHART_FORMULA_REFERENCE_OUT_OF_BOUNDS"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "CHART_FORMULA_TABLE_MISSING"),
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

test("validation reports duplicate defined names in the same scope", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeDefinedName: true }));
  const workbookXml = await pkg.readText("xl/workbook.xml");
  pkg.setText(
    "xl/workbook.xml",
    workbookXml.replace(
      "</definedNames>",
      '<definedName name="RevenueRange">Sheet1!$C$1:$D$2</definedName></definedNames>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "DEFINED_NAME_DUPLICATE");
  assert.equal(report.issues[0]?.target, "RevenueRange");
});

test("validation reports defined names with references outside the Excel grid", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeDefinedName: true }));
  const workbookXml = await pkg.readText("xl/workbook.xml");
  pkg.setText("xl/workbook.xml", workbookXml.replace("Sheet1!$A$1:$B$2", "Sheet1!$XFE$1"));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "DEFINED_NAME_REFERENCE_OUT_OF_BOUNDS");
  assert.equal(report.issues[0]?.target, "RevenueRange");
});

test("validation reports defined names that reference missing tables", async () => {
  const pkg = await openPackage(
    await createMinimalWorkbook({ includeDefinedName: true, includeTable: true })
  );
  const workbookXml = await pkg.readText("xl/workbook.xml");
  pkg.setText("xl/workbook.xml", workbookXml.replace("Sheet1!$A$1:$B$2", "MissingTable[Amount]"));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "DEFINED_NAME_TABLE_MISSING");
  assert.equal(report.issues[0]?.target, "RevenueRange");
});

test("validation reports formulas that reference missing sheets", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c>',
      '<c r="A1" s="1"><f>Missing!A1+Sheet1!A1</f><v>1</v></c>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "FORMULA_SHEET_MISSING");
  assert.equal(report.issues[0]?.target, "Missing");
});

test("validation reports formulas with references outside the Excel grid", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c>',
      '<c r="A1" s="1"><f>XFE1+A1048577</f><v>1</v></c>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 2);
  assert.equal(
    report.issues.every((issue) => issue.code === "FORMULA_REFERENCE_OUT_OF_BOUNDS"),
    true
  );
});

test("validation reports formulas that reference missing tables", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeTable: true }));
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<c r="A2" t="inlineStr"><is><t>Old</t></is></c>',
      '<c r="A2"><f>SUM(MissingTable[Amount],RevenueTable[Amount])</f><v>1</v></c>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "FORMULA_TABLE_MISSING");
  assert.equal(report.issues[0]?.target, "MissingTable");
});

test("validation reports shared formula followers without a master", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c></row>',
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" si="9"/><v>2</v></c></row>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "SHARED_FORMULA_MASTER_MISSING");
  assert.equal(report.issues[0]?.target, "B1");
});

test("validation reports invalid shared formula group metadata", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c></row>',
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" si="abc">A1*2</f><v>2</v></c><c r="C1"><f t="shared" si="0" ref="NOPE">A1*3</f><v>3</v></c><c r="D1"><f t="shared" si="0">A1*4</f><v>4</v></c></row>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 3);
  assert.equal(
    report.issues.some((issue) => issue.code === "SHARED_FORMULA_INDEX_INVALID"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "SHARED_FORMULA_REF_INVALID"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "SHARED_FORMULA_MASTER_DUPLICATE"),
    true
  );
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

test("validation accepts calc chains that point to formula cells", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeCalcChain: true }));
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c>',
      '<c r="A1"><f>1+1</f><v>2</v></c>'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.deepEqual(report.summary, { errors: 0, warnings: 0, infos: 0 });
});

test("validation reports calc chains that point to non-formula cells", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeCalcChain: true }));

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "CALC_CHAIN_CELL_NOT_FORMULA");
  assert.equal(report.issues[0]?.target, "xl/worksheets/sheet1.xml!A1");
});

test("validation reports calc chains with missing sheet ids and invalid addresses", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeCalcChain: true }));
  const calcChainXml = await pkg.readText("xl/calcChain.xml");
  pkg.setText(
    "xl/calcChain.xml",
    calcChainXml.replace('<c r="A1" i="1"/>', '<c r="A1" i="99"/><c r="1A"/>')
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.summary.warnings, 1);
  assert.equal(
    report.issues.some((issue) => issue.code === "CALC_CHAIN_SHEET_MISSING"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "CALC_CHAIN_CELL_REF_INVALID"),
    true
  );
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

test("validation reports table column count mismatches", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeTable: true }));
  const tableXml = await pkg.readText("xl/tables/table1.xml");
  pkg.setText(
    "xl/tables/table1.xml",
    tableXml.replace('<tableColumns count="2">', '<tableColumns count="3">')
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "TABLE_COLUMN_COUNT_MISMATCH");
});

test("validation reports table refs whose width does not match columns", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeTable: true }));
  const tableXml = await pkg.readText("xl/tables/table1.xml");
  pkg.setText(
    "xl/tables/table1.xml",
    tableXml.replace('ref="A1:B2"', 'ref="A1:C2"').replace('ref="A1:B2"', 'ref="A1:C2"')
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.warnings, 1);
  assert.equal(report.issues[0]?.code, "TABLE_COLUMN_REF_WIDTH_MISMATCH");
  assert.equal(report.issues[0]?.target, "A1:C2");
});

test("validation reports duplicate table names", async () => {
  const pkg = await openPackage(
    await createMinimalWorkbook({ includeTable: true, includeSecondTable: true })
  );
  const tableXml = await pkg.readText("xl/tables/table2.xml");
  pkg.setText(
    "xl/tables/table2.xml",
    tableXml.replace(
      'name="ExpenseTable" displayName="ExpenseTable"',
      'name="RevenueTable" displayName="RevenueTable"'
    )
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 1);
  assert.equal(report.issues[0]?.code, "TABLE_NAME_DUPLICATE");
  assert.equal(report.issues[0]?.target, "RevenueTable");
});

test("validation reports invalid table column metadata", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeTable: true }));
  const tableXml = await pkg.readText("xl/tables/table1.xml");
  pkg.setText(
    "xl/tables/table1.xml",
    tableXml
      .replace('<tableColumn id="1" name="Name"/>', '<tableColumn id="1"/>')
      .replace('<tableColumn id="2" name="Amount"/>', '<tableColumn id="1" name="Amount"/>')
  );

  const report = await validateWorkbookPackage(pkg);

  assert.equal(report.summary.errors, 2);
  assert.equal(
    report.issues.some((issue) => issue.code === "TABLE_COLUMN_NAME_MISSING"),
    true
  );
  assert.equal(
    report.issues.some((issue) => issue.code === "TABLE_COLUMN_ID_DUPLICATE"),
    true
  );
});
