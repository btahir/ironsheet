/**
 * Safe cell patch: open a workbook, patch one cell, and only write the
 * output when the mutated workbook still passes validation.
 *
 * Run with: npx tsx examples/01-safe-cell-patch.ts
 */
import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { mutateWorkbookFile, readWorkbookCell } from "@ironsheet/node";
import { createSampleWorkbook } from "./helpers/sample-workbook.ts";

const outputDir = fileURLToPath(new URL("./output/", import.meta.url));
await mkdir(outputDir, { recursive: true });

// Examples generate their own input workbook so they never depend on a
// committed .xlsx fixture. In real usage this file comes from Excel.
const inputPath = `${outputDir}01-input.xlsx`;
const outputPath = `${outputDir}01-patched.xlsx`;
await writeFile(inputPath, await createSampleWorkbook());

// mutateWorkbookFile is the recommended safe-write flow: it validates the
// mutated workbook and returns a package diff before committing bytes.
const report = await mutateWorkbookFile(inputPath, outputPath, async (workbook) => {
  await workbook.patchCell("Sheet1", "A1", "Updated by Ironsheet");
  await workbook.patchCell("Sheet1", "B2", 42000);
});

if (!report.wrote) {
  throw new Error(`Output suppressed: ${report.validation.summary.errors} validation error(s)`);
}

console.log("wrote", outputPath);
console.log("validation summary", report.validation.summary);
console.log("package diff summary", report.diff.summary);
console.log("A1 is now", await readWorkbookCell(outputPath, "Sheet1", "A1"));
