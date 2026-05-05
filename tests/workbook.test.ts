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
