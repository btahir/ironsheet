import assert from "node:assert/strict";
import test from "node:test";
import { diffZipPackages, writeZip } from "../packages/core/src/index.ts";
import { nodeCompressionAdapter, openWorkbook } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

const textEncoder = new TextEncoder();

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

test("package diff separates content changes from ZIP repacking", async () => {
  const data = textEncoder.encode("same content");
  const before = await writeZip([{ name: "xl/workbook.xml", data, compressionMethod: 0 }]);
  const after = await writeZip(
    [{ name: "xl/workbook.xml", data, compressionMethod: 8 }],
    nodeCompressionAdapter
  );

  const diff = diffZipPackages(before, after);
  const entry = diff.entries.find((candidate) => candidate.name === "xl/workbook.xml");

  assert.equal(diff.summary.changed, 0);
  assert.equal(diff.summary.repacked, 1);
  assert.equal(entry?.status, "repacked");
  assert.equal(entry?.contentChanged, false);
  assert.equal(entry?.containerChanged, true);
});
