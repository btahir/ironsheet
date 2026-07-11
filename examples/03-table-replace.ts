/**
 * Table replace: swap the body rows of an Excel table while Ironsheet
 * resizes the table range, filter range, and worksheet cells for you.
 *
 * Run with: npx tsx examples/03-table-replace.ts
 */
import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { mutateWorkbookFile, readWorkbook } from "@ironsheet/node";
import { createSampleWorkbook } from "./helpers/sample-workbook.ts";

const outputDir = fileURLToPath(new URL("./output/", import.meta.url));
await mkdir(outputDir, { recursive: true });

const inputPath = `${outputDir}03-input.xlsx`;
const outputPath = `${outputDir}03-replaced.xlsx`;
await writeFile(
  inputPath,
  await createSampleWorkbook({
    includeTable: true,
    tableRows: [
      ["Stale row", 1],
      ["Another stale row", 2]
    ]
  })
);

const report = await mutateWorkbookFile(inputPath, outputPath, async (workbook) => {
  await workbook.replaceTableRows("RevenueTable", [
    ["North", 42000],
    ["South", 31500],
    ["West", 28750],
    ["Enterprise", 61200]
  ]);
});

if (!report.wrote) {
  throw new Error(`Output suppressed: ${report.validation.summary.errors} validation error(s)`);
}

console.log("wrote", outputPath);
console.log("package diff summary", report.diff.summary);

// Read the table back from the written file to show the resized range.
const workbook = await readWorkbook(outputPath);
const tables = await workbook.tables();
console.log("tables after replace", tables);
