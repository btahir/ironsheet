#!/usr/bin/env tsx
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import {
  createMinimalWorkbook,
  type MinimalWorkbookOptions
} from "../tests/helpers/minimal-xlsx.ts";

type GeneratedCompatibilityFixture = {
  id: string;
  path: string;
  options: MinimalWorkbookOptions;
};

export const generatedCompatibilityFixtures: GeneratedCompatibilityFixture[] = [
  {
    id: "generated-styled-table-report",
    path: "workbooks/generated/styled-table-report.xlsx",
    options: {
      includeConditionalFormatting: true,
      includeDataValidation: true,
      includeDefinedName: true,
      includeTable: true,
      includeTableTotals: true,
      styledTableBody: true,
      tableRows: [
        ["North", 1200],
        ["South", 980],
        ["West", 1430]
      ]
    }
  },
  {
    id: "generated-macro-enabled-model",
    path: "workbooks/generated/macro-enabled-model.xlsm",
    options: {
      includeCalcChain: true,
      includeDefinedName: true,
      includeFormulaCell: true,
      includeMacro: true,
      useSharedStrings: true
    }
  },
  {
    id: "generated-chart-dashboard",
    path: "workbooks/generated/chart-dashboard.xlsx",
    options: {
      includeDefinedName: true,
      includeDrawing: true,
      includeHiddenSheet: true,
      includeTable: true,
      tableRows: [
        ["Plan", 500],
        ["Actual", 625]
      ]
    }
  },
  {
    id: "generated-pivot-cache-workbook",
    path: "workbooks/generated/pivot-cache-workbook.xlsx",
    options: {
      includeDefinedName: true,
      includePivotTable: true,
      includeTable: true,
      tableRows: [
        ["Hardware", 400],
        ["Services", 700]
      ]
    }
  },
  {
    id: "generated-large-sheet-export",
    path: "workbooks/generated/large-sheet-export.xlsx",
    options: {
      includeTable: true,
      styledTableBody: true,
      tableRows: Array.from({ length: 1024 }, (_, index) => [`Row ${index + 1}`, index + 1])
    }
  },
  {
    id: "generated-cross-feature-torture",
    path: "workbooks/generated/cross-feature-torture.xlsm",
    options: {
      includeComment: true,
      includeConditionalFormatting: true,
      includeDataValidation: true,
      includeDefinedName: true,
      includeDrawing: true,
      includeHiddenSheet: true,
      includeHyperlink: true,
      includeMacro: true,
      includePivotTable: true,
      includeSecondTable: true,
      includeTable: true,
      includeTableTotals: true,
      styledTableBody: true,
      tableRows: [
        ["North", 1200],
        ["South", 980],
        ["Enterprise", 4300],
        ["Expansion", 2750]
      ]
    }
  }
];

export async function buildCompatibilityFixtures(rootDir = process.cwd()): Promise<string[]> {
  const corpusDir = resolve(rootDir, "fixtures/corpus");
  const writtenPaths: string[] = [];

  for (const fixture of generatedCompatibilityFixtures) {
    const outputPath = resolve(corpusDir, fixture.path);
    await mkdir(dirname(outputPath), { recursive: true });
    await writeFile(outputPath, await createMinimalWorkbook(fixture.options));
    writtenPaths.push(outputPath);
  }

  return writtenPaths;
}

async function main(): Promise<void> {
  const writtenPaths = await buildCompatibilityFixtures();
  for (const outputPath of writtenPaths) {
    console.log(`compat: wrote fixture ${outputPath}`);
  }
}

if (process.argv[1] !== undefined && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  await main();
}
