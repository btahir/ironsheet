import assert from "node:assert/strict";
import test from "node:test";
import { breakFormulaSheetReferences, parseZip } from "../packages/core/src/index.ts";
import { openWorkbook } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("breakFormulaSheetReferences replaces deleted sheet references with #REF!", () => {
  assert.equal(breakFormulaSheetReferences("SUM(Gone!A1:B2)+A1", "Gone"), "SUM(#REF!)+A1");
  assert.equal(breakFormulaSheetReferences("'Old Data'!$A$1*2", "Old Data"), "#REF!*2");
  assert.equal(breakFormulaSheetReferences("Kept!A1+1", "Gone"), "Kept!A1+1");
  assert.equal(
    breakFormulaSheetReferences('IF(TRUE,"Gone!A1",Gone!A1)', "Gone"),
    'IF(TRUE,"Gone!A1",#REF!)'
  );
});

test("addSheet creates a valid worksheet with relationships and content types", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  const sheet = await workbook.addSheet("Report");

  assert.equal(sheet.name, "Report");
  assert.deepEqual(
    workbook.sheets().map((entry) => entry.name),
    ["Sheet1", "Report"]
  );

  await workbook.patchCell("Report", "B2", "hello");
  assert.equal((await workbook.readCell("Report", "B2"))?.value, "hello");
  assert.equal((await workbook.validate()).summary.errors, 0);

  const reopened = await openWorkbook(await workbook.write());
  assert.deepEqual(
    reopened.sheets().map((entry) => entry.name),
    ["Sheet1", "Report"]
  );
  assert.equal((await reopened.validate()).summary.errors, 0);
});

test("addSheet rejects duplicate and invalid names", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await assert.rejects(workbook.addSheet("Sheet1"), /already used/);
  await assert.rejects(workbook.addSheet("Bad[Name]"), /name/i);
});

test("copySheet duplicates cells and styles without relationship parts", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeTable: true, tableRows: [["Old", 1]] })
  );

  const copy = await workbook.copySheet("Sheet1", "Sheet1 Copy");

  assert.equal(copy.name, "Sheet1 Copy");
  assert.equal(
    (await workbook.readCell("Sheet1 Copy", "A1"))?.value,
    (await workbook.readCell("Sheet1", "A1"))?.value
  );
  assert.equal((await workbook.tables()).length, 1);
  assert.equal(
    workbook
      .diagnostics()
      .some((diagnostic) => diagnostic.code === "SHEET_COPY_FEATURES_NOT_COPIED"),
    true
  );
  assert.equal((await workbook.validate()).summary.errors, 0);

  const reopened = await openWorkbook(await workbook.write());
  assert.equal((await reopened.validate()).summary.errors, 0);
});

test("copySheet keeps external hyperlinks working", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeHyperlink: true }));

  await workbook.copySheet("Sheet1", "Mirror");

  const hyperlinks = await workbook.hyperlinks("Mirror");
  assert.equal(hyperlinks.length, 1);
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("deleteSheet removes the sheet, its parts, and breaks formulas that used it", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeHiddenSheet: true }));
  await workbook.patchCell("Sheet1", "C1", { formula: "HiddenData!A1+1", result: 1 });
  await workbook.showSheet("HiddenData");

  await workbook.deleteSheet("HiddenData");

  assert.deepEqual(
    workbook.sheets().map((entry) => entry.name),
    ["Sheet1"]
  );
  assert.equal((await workbook.readCell("Sheet1", "C1"))?.formula, "#REF!+1");
  assert.equal(
    workbook.diagnostics().some((diagnostic) => diagnostic.code === "SHEET_DELETE_BROKE_FORMULAS"),
    true
  );

  const output = await workbook.write();
  const names = parseZip(output).entries.map((entry) => entry.name);
  assert.equal(names.includes("xl/worksheets/sheet2.xml"), false);

  const reopened = await openWorkbook(output);
  assert.equal((await reopened.validate()).summary.errors, 0);
});

test("deleteSheet refuses to remove the last visible sheet", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeHiddenSheet: true }));

  await assert.rejects(workbook.deleteSheet("Sheet1"), /visible worksheet/);
});

test("deleteSheet cascades to sheet-scoped parts and defined names", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeDefinedName: true, includeHiddenSheet: true })
  );
  await workbook.showSheet("HiddenData");
  await workbook.setDefinedName("HiddenOnly", "HiddenData!$A$1", { sheetName: "HiddenData" });

  await workbook.deleteSheet("HiddenData");

  const names = await workbook.definedNames();
  assert.equal(
    names.some((name) => name.name === "HiddenOnly"),
    false
  );
  assert.equal(
    names.some((name) => name.name === "RevenueRange"),
    true
  );
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("add, copy, and delete compose into a valid workbook round trip", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await workbook.addSheet("Data");
  await workbook.patchRange("Data", "A1", [
    ["Region", "Amount"],
    ["North", 42]
  ]);
  await workbook.copySheet("Data", "Data Copy");
  await workbook.deleteSheet("Data");

  assert.deepEqual(
    workbook.sheets().map((entry) => entry.name),
    ["Sheet1", "Data Copy"]
  );
  assert.equal((await workbook.readCell("Data Copy", "B2"))?.value, 42);

  const reopened = await openWorkbook(await workbook.write());
  assert.equal((await reopened.validate()).summary.errors, 0);
});
