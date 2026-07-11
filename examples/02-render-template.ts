/**
 * Template rendering: fill an Excel-authored template through named
 * anchors (cells, named ranges, and tables) in one transactional patch.
 *
 * Run with: npx tsx examples/02-render-template.ts
 */
import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { preflightWorkbookTemplate, renderWorkbookTemplateSafely } from "@ironsheet/node";
import { createSampleWorkbook } from "./helpers/sample-workbook.ts";

const outputDir = fileURLToPath(new URL("./output/", import.meta.url));
await mkdir(outputDir, { recursive: true });

// Generate a template workbook with a RevenueTable and a RevenueRange
// defined name. `npm run templates:build` produces richer starter
// templates under templates/generated/ that work the same way.
const templatePath = `${outputDir}02-template.xlsx`;
const reportPath = `${outputDir}02-report.xlsx`;
await writeFile(templatePath, await createSampleWorkbook({ includeTable: true }));

const patch = {
  names: [
    {
      name: "RevenueRange",
      values: [
        ["Name", "Amount"],
        ["North", 42000]
      ]
    }
  ],
  tables: [
    {
      tableName: "RevenueTable",
      rows: [
        ["North", 42000],
        ["South", 31500],
        ["West", 28750]
      ]
    }
  ]
};

// Preflight resolves every anchor without mutating anything, so a bad
// patch fails before any cell changes.
const preflight = await preflightWorkbookTemplate(templatePath, patch);
console.log("preflight counts", preflight.counts);

const report = await renderWorkbookTemplateSafely(templatePath, reportPath, patch);

if (!report.wrote) {
  throw new Error(`Output suppressed: ${report.validation.summary.errors} validation error(s)`);
}

console.log("wrote", reportPath);
console.log("validation summary", report.validation.summary);
console.log("package diff summary", report.diff.summary);
