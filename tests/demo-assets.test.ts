import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import process from "node:process";
import test from "node:test";
import { openWorkbook } from "../packages/node/src/index.ts";

test("browser demo sample proves preservation across rich workbook features", async () => {
  const bytes = new Uint8Array(
    await readFile(join(process.cwd(), "website", "public", "samples", "sales-report.xlsx"))
  );
  const workbook = await openWorkbook(bytes);
  const [inspection, validation] = await Promise.all([workbook.inspect(), workbook.validate()]);

  assert.equal(validation.summary.errors, 0);
  assert.equal(validation.summary.warnings, 0);
  assert.equal(inspection.features.charts, 1);
  assert.equal(inspection.features.drawings, 1);
  assert.equal(inspection.features.formulaCells, 5);
  assert.equal(inspection.features.dataValidations, 1);
  assert.equal(inspection.features.tables, 1);
});
