import assert from "node:assert/strict";
import test from "node:test";
import { diffZipPackages } from "../packages/core/src/index.ts";
import { openWorkbook } from "../packages/node/src/index.ts";
import { assertPackageDiff } from "./helpers/package-invariants.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("cell patching only changes the targeted worksheet part", async () => {
  const before = await createMinimalWorkbook({
    includeComment: true,
    includeDrawing: true,
    includeHyperlink: true,
    includeMacro: true
  });
  const workbook = await openWorkbook(before);

  await workbook.patchCell("Sheet1", "D5", "Changed");
  const after = await workbook.write();

  assertPackageDiff(diffZipPackages(before, after), {
    changed: ["xl/worksheets/sheet1.xml"],
    unchanged: [
      "[Content_Types].xml",
      "xl/charts/chart1.xml",
      "xl/comments1.xml",
      "xl/drawings/drawing1.xml",
      "xl/drawings/_rels/drawing1.xml.rels",
      "xl/media/image1.png",
      "xl/vbaProject.bin"
    ]
  });
  const reopened = await openWorkbook(after);
  assert.deepEqual((await reopened.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("table row replacement confines content changes to workbook table and worksheet metadata", async () => {
  const before = await createMinimalWorkbook({
    includeDefinedName: true,
    includeDrawing: true,
    includePivotTable: true,
    includeTable: true,
    includeTableTotals: true,
    styledTableBody: true,
    tableRows: [
      ["North", 1200],
      ["South", 980],
      ["West", 1430]
    ]
  });
  const workbook = await openWorkbook(before);

  await workbook.replaceTableRows("RevenueTable", [
    ["Enterprise", 4300],
    ["Expansion", 2750]
  ]);
  const after = await workbook.write();

  assertPackageDiff(diffZipPackages(before, after), {
    changed: ["xl/tables/table1.xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml"],
    unchanged: [
      "xl/charts/chart1.xml",
      "xl/drawings/drawing1.xml",
      "xl/media/image1.png",
      "xl/pivotCache/pivotCacheDefinition1.xml",
      "xl/pivotTables/pivotTable1.xml"
    ]
  });
  assert.equal(
    workbook
      .diagnostics()
      .some((diagnostic) => diagnostic.code === "PIVOT_TABLES_MAY_NEED_REFRESH"),
    true
  );
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("image insertion into an existing drawing adds only image payload and drawing references", async () => {
  const before = await createMinimalWorkbook({ includeDrawing: true });
  const workbook = await openWorkbook(before);
  const imageData = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 3]);

  await workbook.insertImage("Sheet1", imageData, {
    anchor: {
      kind: "oneCell",
      from: { column: 0, row: 0 },
      ext: { cx: 1000, cy: 1000 }
    }
  });
  const after = await workbook.write();

  assertPackageDiff(diffZipPackages(before, after), {
    added: ["xl/media/image2.png"],
    changed: ["xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels"],
    unchanged: ["[Content_Types].xml", "xl/charts/chart1.xml", "xl/media/image1.png"]
  });
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});
