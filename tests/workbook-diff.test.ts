import assert from "node:assert/strict";
import test from "node:test";
import { diffWorkbooks } from "../packages/core/src/index.ts";
import { openWorkbook } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("diffWorkbooks reports added, changed, and removed cells", async () => {
  const bytes = await createMinimalWorkbook();
  const before = await openWorkbook(bytes);
  const after = await openWorkbook(bytes);

  await after.patchCell("Sheet1", "A1", "Updated");
  await after.patchCell("Sheet1", "B2", 42);
  await after.patchCell("Sheet1", "C3", { formula: "B2*2", result: 84 });

  const diff = await diffWorkbooks(before, after);

  assert.deepEqual(diff.sheets, { added: [], removed: [] });
  assert.equal(diff.summary.changedCells, 1);
  assert.equal(diff.summary.addedCells, 2);
  assert.equal(diff.summary.removedCells, 0);
  assert.equal(diff.summary.truncated, false);

  const changed = diff.cells.find((cell) => cell.kind === "changed");
  assert.equal(changed?.address, "A1");
  assert.equal(changed?.before?.value, "Original");
  assert.equal(changed?.after?.value, "Updated");
  assert.deepEqual(changed?.changed, ["value"]);

  const formulaCell = diff.cells.find((cell) => cell.address === "C3");
  assert.equal(formulaCell?.kind, "added");
  assert.equal(formulaCell?.after?.formula, "B2*2");
});

test("diffWorkbooks reports sheet and defined-name changes", async () => {
  const bytes = await createMinimalWorkbook({ includeDefinedName: true });
  const before = await openWorkbook(bytes);
  const after = await openWorkbook(bytes);

  await after.addSheet("New Sheet");
  await after.setDefinedName("RevenueRange", "Sheet1!$A$1:$C$3");
  await after.setDefinedName("FreshName", "Sheet1!$D$1");

  const diff = await diffWorkbooks(before, after);

  assert.deepEqual(diff.sheets.added, ["New Sheet"]);
  assert.equal(diff.definedNames.changed.includes("RevenueRange"), true);
  assert.equal(diff.definedNames.added.includes("FreshName"), true);
});

test("diffWorkbooks reports table resizes", async () => {
  const bytes = await createMinimalWorkbook({ includeTable: true, tableRows: [["Old", 1]] });
  const before = await openWorkbook(bytes);
  const after = await openWorkbook(bytes);

  await after.replaceTableRows("RevenueTable", [
    ["North", 42],
    ["South", 31]
  ]);

  const diff = await diffWorkbooks(before, after);
  assert.equal(diff.tables.changed.includes("RevenueTable"), true);
});

test("diffWorkbooks detects style-only changes", async () => {
  const bytes = await createMinimalWorkbook();
  const before = await openWorkbook(bytes);
  const after = await openWorkbook(bytes);

  await after.styleCell("Sheet1", "A1", { font: { bold: true } });

  const diff = await diffWorkbooks(before, after);
  const changed = diff.cells.find((cell) => cell.address === "A1");
  assert.deepEqual(changed?.changed, ["style"]);
});

test("diffWorkbooks caps reported cell diffs", async () => {
  const bytes = await createMinimalWorkbook();
  const before = await openWorkbook(bytes);
  const after = await openWorkbook(bytes);

  await after.patchRange("Sheet1", "A10", [
    [1, 2, 3],
    [4, 5, 6]
  ]);

  const diff = await diffWorkbooks(before, after, { maxCellDiffs: 3 });
  assert.equal(diff.cells.length, 3);
  assert.equal(diff.summary.addedCells, 6);
  assert.equal(diff.summary.truncated, true);
});
