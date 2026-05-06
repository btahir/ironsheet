#!/usr/bin/env tsx
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import {
  createMinimalWorkbook,
  type MinimalWorkbookOptions
} from "../tests/helpers/minimal-xlsx.ts";

type TemplateManifest = {
  schemaVersion: 1;
  templates: Array<{
    id: string;
    path: string;
    title: string;
    description: string;
    patchExample: unknown;
  }>;
};

type TemplateDefinition = {
  id: string;
  options: MinimalWorkbookOptions;
};

const templateDefinitions: TemplateDefinition[] = [
  {
    id: "styled-report-template",
    options: {
      includeAutoFilter: true,
      includeConditionalFormatting: true,
      includeDataValidation: true,
      includeDefinedName: true,
      includeTable: true,
      includeTableTotals: true,
      styledTableBody: true,
      tableRows: [
        ["North", 1200],
        ["South", 980],
        ["West", 1430],
        ["Enterprise", 1890]
      ]
    }
  },
  {
    id: "macro-model-template",
    options: {
      includeCalcChain: true,
      includeDefinedName: true,
      includeFormulaCell: true,
      includeMacro: true,
      styledTableBody: true,
      useSharedStrings: true
    }
  },
  {
    id: "dashboard-template",
    options: {
      includeDefinedName: true,
      includeDrawing: true,
      includeHiddenSheet: true,
      includeTable: true,
      includeTableTotals: true,
      styledTableBody: true,
      tableRows: [
        ["Plan", 500],
        ["Actual", 625],
        ["Forecast", 680]
      ]
    }
  },
  {
    id: "pivot-source-template",
    options: {
      includeDefinedName: true,
      includePivotTable: true,
      includeTable: true,
      styledTableBody: true,
      tableRows: [
        ["Hardware", 400],
        ["Services", 700],
        ["Support", 300]
      ]
    }
  },
  {
    id: "large-export-template",
    options: {
      includeTable: true,
      styledTableBody: true,
      tableRows: Array.from({ length: 2048 }, (_, index) => [`Row ${index + 1}`, index + 1])
    }
  }
];

export async function buildTemplates(rootDir = process.cwd()): Promise<string[]> {
  const manifestPath = resolve(rootDir, "templates/manifest.json");
  const manifest = JSON.parse(await readFile(manifestPath, "utf8")) as TemplateManifest;
  const definitions = new Map(templateDefinitions.map((definition) => [definition.id, definition]));
  const writtenPaths: string[] = [];

  for (const template of manifest.templates) {
    const definition = definitions.get(template.id);
    if (definition === undefined) {
      throw new Error(`Missing template definition for ${template.id}`);
    }

    const outputPath = resolve(rootDir, "templates", template.path);
    await mkdir(dirname(outputPath), { recursive: true });
    await writeFile(outputPath, await createMinimalWorkbook(definition.options));
    writtenPaths.push(outputPath);
  }

  return writtenPaths;
}

async function main(): Promise<void> {
  const writtenPaths = await buildTemplates();
  for (const outputPath of writtenPaths) {
    console.log(`templates: wrote ${outputPath}`);
  }
}

if (process.argv[1] !== undefined && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  await main();
}
