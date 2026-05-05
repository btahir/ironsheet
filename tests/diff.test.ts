import assert from "node:assert/strict";
import test from "node:test";
import { diffZipPackages } from "../packages/core/src/index.ts";
import { openWorkbook } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("package diff identifies only changed workbook parts", async () => {
  const before = await createMinimalWorkbook();
  const workbook = await openWorkbook(before);

  await workbook.patchCell("Sheet1", "B2", "Changed");
  const after = await workbook.write();
  const diff = diffZipPackages(before, after);

  assert.equal(diff.summary.added, 0);
  assert.equal(diff.summary.removed, 0);
  assert.equal(diff.entries.find((entry) => entry.name === "xl/styles.xml")?.status, "unchanged");
  assert.equal(
    diff.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml")?.status,
    "changed"
  );
});
