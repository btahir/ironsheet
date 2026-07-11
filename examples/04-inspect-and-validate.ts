/**
 * Inspect and validate: read a workbook without mutating it, list its
 * sheets, tables, and defined names, and run semantic validation.
 *
 * Run with: npx tsx examples/04-inspect-and-validate.ts
 */
import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { readWorkbook, validateWorkbookFile } from "@ironsheet/node";
import { createSampleWorkbook } from "./helpers/sample-workbook.ts";

const outputDir = fileURLToPath(new URL("./output/", import.meta.url));
await mkdir(outputDir, { recursive: true });

const inputPath = `${outputDir}04-input.xlsx`;
await writeFile(inputPath, await createSampleWorkbook({ includeTable: true }));

const workbook = await readWorkbook(inputPath);

console.log("sheets", await workbook.sheets());
console.log("tables", await workbook.tables());
console.log("defined names", await workbook.definedNames());
console.log("cell A1", await workbook.readCell("Sheet1", "A1"));

// Full inspection report: package structure plus workbook features.
const inspection = await workbook.inspect();
console.log("inspection", JSON.stringify(inspection, null, 2));

// validateWorkbookFile is the same validator the safe-write flow uses to
// decide whether output bytes may be written.
const validation = await validateWorkbookFile(inputPath);
console.log("validation summary", validation.summary);
if (validation.issues.length > 0) {
  console.log("issues", validation.issues);
}
