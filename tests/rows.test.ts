import assert from "node:assert/strict";
import test from "node:test";
import { parseZip, shiftFormulaRowReferences } from "../packages/core/src/index.ts";
import { openWorkbook } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("shiftFormulaRowReferences shifts rows on insert", () => {
  const edit = { count: 2, mode: "insert" as const, sheetName: "Sheet1", startRow: 5 };

  assert.equal(
    shiftFormulaRowReferences("SUM(A1:A10)+B7", edit, { defaultSheetName: "Sheet1" }),
    "SUM(A1:A12)+B9"
  );
  assert.equal(
    shiftFormulaRowReferences("$B$4+$B$5", edit, { defaultSheetName: "Sheet1" }),
    "$B$4+$B$7"
  );
  assert.equal(
    shiftFormulaRowReferences("Sheet1!A6+Other!A6", edit, { defaultSheetName: "Other" }),
    "Sheet1!A8+Other!A6"
  );
  assert.equal(
    shiftFormulaRowReferences("'My Sheet'!A6", { ...edit, sheetName: "My Sheet" }),
    "'My Sheet'!A8"
  );
  assert.equal(
    shiftFormulaRowReferences("SUM(A1:A4)", edit, { defaultSheetName: "Sheet1" }),
    "SUM(A1:A4)"
  );
  assert.equal(
    shiftFormulaRowReferences('IF(A9="A9",A9,0)', edit, { defaultSheetName: "Sheet1" }),
    'IF(A11="A9",A11,0)'
  );
});

test("shiftFormulaRowReferences shifts, shrinks, and breaks references on delete", () => {
  const edit = { count: 2, mode: "delete" as const, sheetName: "Sheet1", startRow: 5 };

  assert.equal(
    shiftFormulaRowReferences("SUM(A1:A10)+A8", edit, { defaultSheetName: "Sheet1" }),
    "SUM(A1:A8)+A6"
  );
  assert.equal(
    shiftFormulaRowReferences("A5+A6", edit, { defaultSheetName: "Sheet1" }),
    "#REF!+#REF!"
  );
  assert.equal(
    shiftFormulaRowReferences("SUM(A5:A6)", edit, { defaultSheetName: "Sheet1" }),
    "SUM(#REF!)"
  );
  assert.equal(
    shiftFormulaRowReferences("SUM(A5:A8)", edit, { defaultSheetName: "Sheet1" }),
    "SUM(A5:A6)"
  );
  assert.equal(
    shiftFormulaRowReferences("Sheet1!B6", edit, { defaultSheetName: "Other" }),
    "Sheet1!#REF!"
  );
  assert.equal(
    shiftFormulaRowReferences("Table1[Amount]+A4", edit, { defaultSheetName: "Sheet1" }),
    "Table1[Amount]+A4"
  );
});

test("insertRows shifts rows, values, and dimension", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());
  await workbook.patchRange("Sheet1", "A2", [["second"], ["third"]]);

  await workbook.insertRows("Sheet1", 2, 2);

  assert.equal(await workbook.readCell("Sheet1", "A2"), undefined);
  assert.equal(await workbook.readCell("Sheet1", "A3"), undefined);
  assert.equal((await workbook.readCell("Sheet1", "A4"))?.value, "second");
  assert.equal((await workbook.readCell("Sheet1", "A5"))?.value, "third");
  assert.equal((await workbook.readCell("Sheet1", "A1"))?.value, "Original");
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("deleteRows removes rows and shifts the rest up", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());
  await workbook.patchRange("Sheet1", "A2", [["second"], ["third"], ["fourth"]]);

  await workbook.deleteRows("Sheet1", 2, 2);

  assert.equal((await workbook.readCell("Sheet1", "A1"))?.value, "Original");
  assert.equal((await workbook.readCell("Sheet1", "A2"))?.value, "fourth");
  assert.equal(await workbook.readCell("Sheet1", "A3"), undefined);
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("insertRows rewrites same-sheet and cross-sheet formulas", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeHiddenSheet: true }));
  await workbook.patchCell("Sheet1", "A5", 10);
  await workbook.patchCell("Sheet1", "B1", { formula: "SUM(A5:A6)", result: 10 });
  await workbook.patchCell("HiddenData", "A1", { formula: "Sheet1!A5*2", result: 20 });

  await workbook.insertRows("Sheet1", 3, 2);

  assert.equal((await workbook.readCell("Sheet1", "B1"))?.formula, "SUM(A7:A8)");
  assert.equal((await workbook.readCell("HiddenData", "A1"))?.formula, "Sheet1!A7*2");
  assert.equal((await workbook.readCell("Sheet1", "A7"))?.value, 10);
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("deleteRows marks dead references with #REF!", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());
  await workbook.patchCell("Sheet1", "A5", 10);
  await workbook.patchCell("Sheet1", "B1", { formula: "A5+1", result: 11 });

  await workbook.deleteRows("Sheet1", 5, 1);

  assert.equal((await workbook.readCell("Sheet1", "B1"))?.formula, "#REF!+1");
});

test("row edits update defined names, merges, hyperlinks, and validations", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({
      includeDataValidation: true,
      includeDefinedName: true,
      includeHyperlink: true,
      includeMerge: true
    })
  );

  await workbook.insertRows("Sheet1", 1, 3);

  const names = await workbook.definedNames();
  const revenueRange = names.find((name) => name.name === "RevenueRange");
  assert.equal(revenueRange?.text, "Sheet1!$A$4:$B$5");

  const merges = await workbook.mergedCells("Sheet1");
  assert.deepEqual(
    merges.map((merge) => merge.ref),
    ["A4:B4"]
  );

  const validations = await workbook.dataValidations("Sheet1");
  assert.deepEqual(
    validations.map((validation) => validation.sqref),
    ["B5:B13"]
  );

  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("deleteRows drops merges, hyperlinks, and validations that lose their rows", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeHyperlink: true, includeMerge: true })
  );

  await workbook.deleteRows("Sheet1", 1, 1);

  assert.deepEqual(await workbook.mergedCells("Sheet1"), []);
  assert.deepEqual(await workbook.hyperlinks("Sheet1"), []);
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("insertRows above a table shifts the table range", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeTable: true, tableRows: [["Old", 1]] })
  );

  await assert.rejects(workbook.insertRows("Sheet1", 2, 1), /overlaps table/);

  await workbook.insertRows("Sheet1", 1, 2);
  const tables = await workbook.tables();
  assert.equal(tables[0]?.ref, "A3:B4");
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("deleteRows refuses to intersect a table", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeTable: true, tableRows: [["Old", 1]] })
  );

  await assert.rejects(workbook.deleteRows("Sheet1", 2, 1), /overlaps table/);
});

test("row edits remove stale calculation chains and mark recalculation", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeCalcChain: true, includeFormulaCell: true })
  );

  await workbook.insertRows("Sheet1", 1, 1);

  const outputZip = parseZip(await workbook.write());
  assert.equal(
    outputZip.entries.some((entry) => entry.name === "xl/calcChain.xml"),
    false
  );
});

test("row edits shift comment references", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeComment: true }));
  const before = await workbook.comments();
  assert.equal(before.length, 1);
  const originalRef = before[0]?.ref;
  assert.ok(originalRef);

  await workbook.insertRows("Sheet1", 1, 2);

  const after = await workbook.comments();
  const originalRow = Number.parseInt(originalRef.replace(/^[A-Z]+/i, ""), 10);
  const expectedRef = `${originalRef.replace(/[0-9]+$/, "")}${originalRow + 2}`;
  assert.equal(after[0]?.ref, expectedRef);
});

test("deleteRows refuses to orphan a surviving shared formula group", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B3"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>
    <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>
  </sheetData>
</worksheet>`;
  const pkgWorkbook = workbook as unknown as {
    pkg: { setText: (name: string, xml: string) => void };
  };
  pkgWorkbook.pkg.setText("xl/worksheets/sheet1.xml", sheetXml);

  await assert.rejects(workbook.deleteRows("Sheet1", 1, 1), /shared or array formula/);
  await workbook.deleteRows("Sheet1", 1, 3);
  assert.equal(await workbook.readCell("Sheet1", "B1"), undefined);
});
