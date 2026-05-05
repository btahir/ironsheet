import assert from "node:assert/strict";
import test from "node:test";
import { parseZip, readEntryData } from "../packages/core/src/index.ts";
import { openWorkbook, nodeCompressionAdapter } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

const textDecoder = new TextDecoder();

test("inspects workbook sheets from OPC relationships", async () => {
  const data = await createMinimalWorkbook();
  const workbook = await openWorkbook(data);

  assert.deepEqual(workbook.sheets(), [
    {
      name: "Sheet1",
      id: "1",
      relationshipId: "rId1",
      partName: "xl/worksheets/sheet1.xml"
    }
  ]);

  assert.deepEqual((await workbook.inspect()).features, {
    calcChains: 0,
    charts: 0,
    comments: 0,
    conditionalFormats: 0,
    dataValidations: 0,
    definedNames: 0,
    drawings: 0,
    hiddenSheets: 0,
    hyperlinks: 0,
    macros: 0,
    media: 0,
    merges: 0,
    pivotTables: 0,
    sharedStrings: 0,
    tables: 0
  });
});

test("inspect reports workbook feature signals", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({
      includeCalcChain: true,
      includeConditionalFormatting: true,
      includeDataValidation: true,
      includeDrawing: true,
      includeHiddenSheet: true,
      includeHyperlink: true,
      includeMacro: true,
      includeMerge: true,
      includeTable: true,
      useSharedStrings: true
    })
  );

  assert.deepEqual(workbook.sheets(), [
    {
      name: "Sheet1",
      id: "1",
      relationshipId: "rId1",
      partName: "xl/worksheets/sheet1.xml"
    },
    {
      name: "HiddenData",
      id: "2",
      relationshipId: "rIdHidden",
      partName: "xl/worksheets/sheet2.xml",
      state: "hidden"
    }
  ]);

  assert.deepEqual((await workbook.inspect()).features, {
    calcChains: 1,
    charts: 1,
    comments: 0,
    conditionalFormats: 1,
    dataValidations: 1,
    definedNames: 0,
    drawings: 1,
    hiddenSheets: 1,
    hyperlinks: 1,
    macros: 1,
    media: 1,
    merges: 1,
    pivotTables: 0,
    sharedStrings: 1,
    tables: 1
  });
});

test("reads existing cell values and style ids", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  assert.deepEqual(await workbook.readCell("Sheet1", "A1"), {
    address: "A1",
    value: "Original",
    styleId: "1"
  });
  assert.equal(await workbook.readCell("Sheet1", "Z99"), undefined);
});

test("reads shared string cell values", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ useSharedStrings: true }));

  assert.deepEqual(await workbook.readCell("Sheet1", "A1"), {
    address: "A1",
    value: "Original",
    styleId: "1"
  });
});

test("patches and reads ranges", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await workbook.patchRange("Sheet1", "B2", [
    ["Move fast", 12],
    [true, { formula: "=C2*2", result: 24 }]
  ]);

  assert.deepEqual(await workbook.readRange("Sheet1", "B2:C3"), {
    range: "B2:C3",
    cells: [
      [
        { address: "B2", value: "Move fast" },
        { address: "C2", value: 12 }
      ],
      [
        { address: "B3", value: true },
        { address: "C3", value: 24, formula: "C2*2" }
      ]
    ]
  });
});

test("patches one cell and preserves untouched entry payloads", async () => {
  const original = await createMinimalWorkbook();
  const originalZip = parseZip(original);
  const originalStyles = originalZip.entries.find((entry) => entry.name === "xl/styles.xml");
  assert.ok(originalStyles);

  const workbook = await openWorkbook(original);
  await workbook.patchCell("Sheet1", "B2", "Move fast");
  const output = await workbook.write();
  const outputZip = parseZip(output);

  const outputStyles = outputZip.entries.find((entry) => entry.name === "xl/styles.xml");
  assert.ok(outputStyles);
  assert.deepEqual(outputStyles.compressedData, originalStyles.compressedData);

  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));

  assert.match(sheetXml, /<c r="B2" t="inlineStr"><is><t>Move fast<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<dimension ref="A1:B2"\/>/);
});

test("patching an existing styled cell preserves the style id", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await workbook.patchCell("Sheet1", "A1", "Styled replacement");
  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));

  assert.match(sheetXml, /<c r="A1" s="1" t="inlineStr"><is><t>Styled replacement<\/t><\/is><\/c>/);
});

test("patches formulas and marks workbook for recalculation", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await workbook.patchCell("Sheet1", "C3", { formula: "=SUM(A1:B2)" });
  const outputZip = parseZip(await workbook.write());

  const workbookEntry = outputZip.entries.find((entry) => entry.name === "xl/workbook.xml");
  assert.ok(workbookEntry);
  const workbookXml = textDecoder.decode(
    await readEntryData(workbookEntry, nodeCompressionAdapter)
  );

  assert.match(workbookXml, /<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"\/>/);
});

test("formula patches remove stale calculation chain parts", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeCalcChain: true }));

  await workbook.patchCell("Sheet1", "C3", { formula: "=SUM(A1:B2)" });
  const outputZip = parseZip(await workbook.write());
  const names = outputZip.entries.map((entry) => entry.name);

  assert.equal(names.includes("xl/calcChain.xml"), false);

  const relsEntry = outputZip.entries.find((entry) => entry.name === "xl/_rels/workbook.xml.rels");
  assert.ok(relsEntry);
  const relsXml = textDecoder.decode(await readEntryData(relsEntry, nodeCompressionAdapter));
  assert.doesNotMatch(relsXml, /calcChain/);

  const contentTypesEntry = outputZip.entries.find((entry) => entry.name === "[Content_Types].xml");
  assert.ok(contentTypesEntry);
  const contentTypesXml = textDecoder.decode(
    await readEntryData(contentTypesEntry, nodeCompressionAdapter)
  );
  assert.doesNotMatch(contentTypesXml, /calcChain/);
});

test("cell patches preserve macro-enabled workbook parts", async () => {
  const original = await createMinimalWorkbook({ includeMacro: true });
  const originalZip = parseZip(original);
  const originalVba = originalZip.entries.find((entry) => entry.name === "xl/vbaProject.bin");
  assert.ok(originalVba);

  const workbook = await openWorkbook(original);
  await workbook.patchCell("Sheet1", "B2", "macro safe");
  const outputZip = parseZip(await workbook.write());
  const outputVba = outputZip.entries.find((entry) => entry.name === "xl/vbaProject.bin");
  assert.ok(outputVba);

  assert.deepEqual(outputVba.compressedData, originalVba.compressedData);

  const relsEntry = outputZip.entries.find((entry) => entry.name === "xl/_rels/workbook.xml.rels");
  assert.ok(relsEntry);
  const relsXml = textDecoder.decode(await readEntryData(relsEntry, nodeCompressionAdapter));
  assert.match(relsXml, /vbaProject/);

  const contentTypesEntry = outputZip.entries.find((entry) => entry.name === "[Content_Types].xml");
  assert.ok(contentTypesEntry);
  const contentTypesXml = textDecoder.decode(
    await readEntryData(contentTypesEntry, nodeCompressionAdapter)
  );
  assert.match(contentTypesXml, /macroEnabled\.main\+xml/);
  assert.match(contentTypesXml, /vbaProject/);
});

test("cell patches preserve merge cells and hyperlink relationships", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeHyperlink: true, includeMerge: true })
  );

  assert.deepEqual(workbook.diagnostics(), [
    {
      severity: "warning",
      code: "EXTERNAL_RELATIONSHIP_PRESERVED",
      message: "Preserved external relationship rIdHyperlink1",
      part: "xl/worksheets/sheet1.xml"
    }
  ]);

  await workbook.patchCell("Sheet1", "C3", "safe");
  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(sheetXml, /<mergeCell ref="A1:B1"\/>/);
  assert.match(sheetXml, /<hyperlink ref="A1" r:id="rIdHyperlink1"\/>/);

  const rels = outputZip.entries.find(
    (entry) => entry.name === "xl/worksheets/_rels/sheet1.xml.rels"
  );
  assert.ok(rels);
  const relsXml = textDecoder.decode(await readEntryData(rels, nodeCompressionAdapter));
  assert.match(relsXml, /Target="https:\/\/example.com"/);
  assert.match(relsXml, /TargetMode="External"/);
});

test("cell patches preserve validation, formatting, drawing, chart, media, and hidden sheets", async () => {
  const original = await createMinimalWorkbook({
    includeConditionalFormatting: true,
    includeDataValidation: true,
    includeDrawing: true,
    includeHiddenSheet: true
  });
  const originalZip = parseZip(original);
  const originalImage = originalZip.entries.find((entry) => entry.name === "xl/media/image1.png");
  const originalHiddenSheet = originalZip.entries.find(
    (entry) => entry.name === "xl/worksheets/sheet2.xml"
  );
  assert.ok(originalImage);
  assert.ok(originalHiddenSheet);

  const workbook = await openWorkbook(original);
  await workbook.patchCell("Sheet1", "C3", "preserve complex parts");
  const outputZip = parseZip(await workbook.write());

  const outputImage = outputZip.entries.find((entry) => entry.name === "xl/media/image1.png");
  const outputHiddenSheet = outputZip.entries.find(
    (entry) => entry.name === "xl/worksheets/sheet2.xml"
  );
  assert.ok(outputImage);
  assert.ok(outputHiddenSheet);
  assert.deepEqual(outputImage.compressedData, originalImage.compressedData);
  assert.deepEqual(outputHiddenSheet.compressedData, originalHiddenSheet.compressedData);

  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(sheetXml, /<conditionalFormatting sqref="A1:A10">/);
  assert.match(sheetXml, /<dataValidation type="whole" operator="between"/);
  assert.match(sheetXml, /<drawing r:id="rIdDrawing1"\/>/);

  const drawing = outputZip.entries.find((entry) => entry.name === "xl/drawings/drawing1.xml");
  const chart = outputZip.entries.find((entry) => entry.name === "xl/charts/chart1.xml");
  const workbookEntry = outputZip.entries.find((entry) => entry.name === "xl/workbook.xml");
  assert.ok(drawing);
  assert.ok(chart);
  assert.ok(workbookEntry);
  assert.match(
    textDecoder.decode(await readEntryData(drawing, nodeCompressionAdapter)),
    /r:id="rIdChart1"/
  );
  assert.match(
    textDecoder.decode(await readEntryData(chart, nodeCompressionAdapter)),
    /<a:t>Revenue<\/a:t>/
  );
  assert.match(
    textDecoder.decode(await readEntryData(workbookEntry, nodeCompressionAdapter)),
    /state="hidden"/
  );
});

test("diagnostics report macro preservation and calc chain invalidation", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeCalcChain: true, includeMacro: true })
  );

  assert.equal(
    workbook.diagnostics().some((diagnostic) => diagnostic.code === "MACRO_PROJECT_PRESERVED"),
    true
  );

  await workbook.patchCell("Sheet1", "C3", { formula: "=SUM(A1:B2)" });

  assert.equal(
    workbook.diagnostics().some((diagnostic) => diagnostic.code === "FORMULA_CALC_CHAIN_REMOVED"),
    true
  );
});

test("replaces basic table body rows and updates the table ref", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeTable: true }));

  const table = await workbook.replaceTableRows("RevenueTable", [
    ["New", 10],
    ["Growth", 20]
  ]);
  assert.equal(table.ref, "A1:B3");

  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(
    sheetXml,
    /<row r="2"><c r="A2" t="inlineStr"><is><t>New<\/t><\/is><\/c><c r="B2"><v>10<\/v><\/c><\/row>/
  );
  assert.match(
    sheetXml,
    /<row r="3"><c r="A3" t="inlineStr"><is><t>Growth<\/t><\/is><\/c><c r="B3"><v>20<\/v><\/c><\/row>/
  );

  const tableEntry = outputZip.entries.find((entry) => entry.name === "xl/tables/table1.xml");
  assert.ok(tableEntry);
  const tableXml = textDecoder.decode(await readEntryData(tableEntry, nodeCompressionAdapter));
  assert.match(tableXml, /ref="A1:B3"/);
  assert.match(tableXml, /<autoFilter ref="A1:B3"\/>/);
});

test("table row replacement preserves body styles and shrinks dimensions", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({
      includeTable: true,
      styledTableBody: true,
      tableRows: [
        ["Old", 1],
        ["Older", 2],
        ["Oldest", 3]
      ]
    })
  );

  const table = await workbook.replaceTableRows("RevenueTable", [["Fresh", 99]]);
  assert.equal(table.ref, "A1:B2");

  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(sheetXml, /<dimension ref="A1:B2"\/>/);
  assert.match(
    sheetXml,
    /<row r="2" s="2" customFormat="1"><c r="A2" s="3" t="inlineStr"><is><t>Fresh<\/t><\/is><\/c><c r="B2" s="4"><v>99<\/v><\/c><\/row>/
  );
  assert.doesNotMatch(sheetXml, /Oldest/);

  const tableEntry = outputZip.entries.find((entry) => entry.name === "xl/tables/table1.xml");
  assert.ok(tableEntry);
  const tableXml = textDecoder.decode(await readEntryData(tableEntry, nodeCompressionAdapter));
  assert.match(tableXml, /ref="A1:B2"/);
  assert.match(tableXml, /<autoFilter ref="A1:B2"\/>/);
});

test("table formula writes invalidate stale calculation chains", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeCalcChain: true, includeTable: true })
  );

  await workbook.replaceTableRows("RevenueTable", [
    ["Formula", { formula: "=SUM(1,2)", result: 3 }]
  ]);
  const outputZip = parseZip(await workbook.write());
  const names = outputZip.entries.map((entry) => entry.name);

  assert.equal(names.includes("xl/calcChain.xml"), false);
});
