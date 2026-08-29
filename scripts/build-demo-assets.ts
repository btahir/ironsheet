#!/usr/bin/env tsx
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { openWorkbook } from "../packages/node/src/index.ts";

const outputDirectory = resolve("website/public/samples");
const workbookPath = resolve(outputDirectory, "sales-report.xlsx");
await mkdir(outputDirectory, { recursive: true });

// Keep this application-authored workbook committed as a real preservation target.
// The asset build verifies its launch-critical features instead of replacing it
// with a synthetic OOXML fixture.
const workbook = await openWorkbook(new Uint8Array(await readFile(workbookPath)));
const [inspection, validation, tables] = await Promise.all([
  workbook.inspect(),
  workbook.validate(),
  workbook.tables()
]);

if (validation.summary.errors > 0 || validation.summary.warnings > 0) {
  throw new Error(
    `Demo workbook failed validation (${validation.summary.errors} errors, ${validation.summary.warnings} warnings)`
  );
}

const expectedFeatures = {
  charts: 1,
  dataValidations: 1,
  drawings: 1,
  formulaCells: 5,
  tables: 1
} as const;

for (const [feature, expected] of Object.entries(expectedFeatures)) {
  const actual = inspection.features[feature as keyof typeof expectedFeatures];
  if (actual !== expected) {
    throw new Error(`Demo workbook expected ${expected} ${feature}; found ${actual}`);
  }
}

if (!tables.some((table) => table.name === "RevenueTable")) {
  throw new Error("Demo workbook is missing RevenueTable");
}

const csv = `Name,Amount
North,54000
South,41800
East,36450
West,49750
`;

await writeFile(resolve(outputDirectory, "fresh-sales.csv"), csv, "utf8");

console.log(`demo-assets: verified ${workbookPath} and wrote fresh-sales.csv`);
