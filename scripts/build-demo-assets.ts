#!/usr/bin/env tsx
import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { createMinimalWorkbook } from "../tests/helpers/minimal-xlsx.ts";

const outputDirectory = resolve("website/public/samples");
await mkdir(outputDirectory, { recursive: true });

const workbook = await createMinimalWorkbook({
  includeDefinedName: true,
  includeHiddenSheet: true,
  includeTable: true,
  includeTableTotals: true,
  styledTableBody: true,
  tableRows: [
    ["North", 42_000],
    ["South", 37_500],
    ["East", 31_250],
    ["West", 46_700]
  ]
});

const csv = `Name,Amount
North,54000
South,41800
East,36450
West,49750
International,28500
`;

await writeFile(resolve(outputDirectory, "sales-report.xlsx"), workbook);
await writeFile(resolve(outputDirectory, "fresh-sales.csv"), csv, "utf8");

console.log(`demo-assets: wrote ${outputDirectory}`);
