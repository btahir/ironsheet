#!/usr/bin/env tsx
import process from "node:process";
import { readFile } from "node:fs/promises";
import { diffZipPackages } from "@ironsheet/core";
import {
  appendWorkbookRows,
  appendWorkbookTableColumn,
  inspectWorkbookStyles,
  listWorkbookFormulas,
  listWorkbookTables,
  patchWorkbookCell,
  patchWorkbookRange,
  readWorkbook,
  readWorkbookCell,
  readWorkbookRange,
  removeRightmostWorkbookTableColumn,
  renameWorkbookTable,
  renameWorkbookTableColumn,
  replaceWorkbookTableRows,
  retargetWorkbookChartFormulasFile,
  retargetWorkbookPivotCacheSourcesFile,
  styleWorkbookCell,
  validateWorkbookFile
} from "@ironsheet/node";
import type {
  CellInput,
  ChartFormulaRetarget,
  PivotCacheSourceRetarget,
  WorkbookCellStyleInput
} from "@ironsheet/core";

type Command =
  | "inspect"
  | "append-rows"
  | "append-table-column"
  | "formulas"
  | "patch"
  | "patch-range"
  | "read-cell"
  | "read-range"
  | "rename-table-column"
  | "rename-table"
  | "remove-table-column"
  | "replace-table"
  | "retarget-chart"
  | "retarget-pivot"
  | "style-cell"
  | "styles"
  | "tables"
  | "validate"
  | "diff";

function usage(): never {
  console.error(`usage:
  npm run cli -- inspect <workbook.xlsx>
  npm run cli -- tables <workbook.xlsx>
  npm run cli -- formulas <workbook.xlsx>
  npm run cli -- styles <workbook.xlsx>
  npm run cli -- validate <workbook.xlsx>
  npm run cli -- read-cell <workbook.xlsx> <sheet> <cell>
  npm run cli -- read-range <workbook.xlsx> <sheet> <range>
  npm run cli -- patch <input.xlsx> <output.xlsx> <sheet> <cell> <value>
  npm run cli -- style-cell <input.xlsx> <output.xlsx> <sheet> <cell> <jsonStyle>
  npm run cli -- patch-range <input.xlsx> <output.xlsx> <sheet> <startCell> <jsonRows>
  npm run cli -- append-rows <input.xlsx> <output.xlsx> <sheet> <jsonRows>
  npm run cli -- append-table-column <input.xlsx> <output.xlsx> <table> <column> [jsonValues]
  npm run cli -- rename-table <input.xlsx> <output.xlsx> <table> <newName>
  npm run cli -- rename-table-column <input.xlsx> <output.xlsx> <table> <column> <newName>
  npm run cli -- remove-table-column <input.xlsx> <output.xlsx> <table> <rightmostColumn>
  npm run cli -- replace-table <input.xlsx> <output.xlsx> <table> <jsonRows>
  npm run cli -- retarget-chart <input.xlsx> <output.xlsx> <jsonRetargets>
  npm run cli -- retarget-pivot <input.xlsx> <output.xlsx> <jsonRetargets>
  npm run cli -- diff <before.xlsx> <after.xlsx>

value examples:
  hello
  123
  true
  =SUM(A1:A3)`);
  process.exit(2);
}

async function inspect(path: string): Promise<void> {
  const workbook = await readWorkbook(path);
  const result = await workbook.inspect();
  console.log(JSON.stringify(result, null, 2));
}

async function tables(path: string): Promise<void> {
  console.log(JSON.stringify(await listWorkbookTables(path), null, 2));
}

async function formulas(path: string): Promise<void> {
  console.log(JSON.stringify(await listWorkbookFormulas(path), null, 2));
}

async function styles(path: string): Promise<void> {
  console.log(JSON.stringify(await inspectWorkbookStyles(path), null, 2));
}

async function validate(path: string): Promise<void> {
  const report = await validateWorkbookFile(path);
  console.log(JSON.stringify(report, null, 2));

  if (report.summary.errors > 0) {
    process.exitCode = 1;
  }
}

async function readCellCommand(path: string, sheetName: string, address: string): Promise<void> {
  console.log(JSON.stringify(await readWorkbookCell(path, sheetName, address), null, 2));
}

async function readRangeCommand(path: string, sheetName: string, rangeRef: string): Promise<void> {
  console.log(JSON.stringify(await readWorkbookRange(path, sheetName, rangeRef), null, 2));
}

async function diff(beforePath: string, afterPath: string): Promise<void> {
  const before = new Uint8Array(await readFile(beforePath));
  const after = new Uint8Array(await readFile(afterPath));
  console.log(JSON.stringify(diffZipPackages(before, after), null, 2));
}

async function patch(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  address: string,
  rawValue: string
): Promise<void> {
  await patchWorkbookCell(inputPath, outputPath, sheetName, address, parseCliValue(rawValue));
  console.log(`patched ${sheetName}!${address} -> ${outputPath}`);
}

async function styleCellCommand(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  address: string,
  rawStyle: string
): Promise<void> {
  await styleWorkbookCell(inputPath, outputPath, sheetName, address, parseStyle(rawStyle));
  console.log(`styled ${sheetName}!${address} -> ${outputPath}`);
}

async function patchRangeCommand(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  startAddress: string,
  rawRows: string
): Promise<void> {
  await patchWorkbookRange(inputPath, outputPath, sheetName, startAddress, parseRows(rawRows));
  console.log(`patched ${sheetName}!${startAddress} range -> ${outputPath}`);
}

async function appendRowsCommand(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  rawRows: string
): Promise<void> {
  await appendWorkbookRows(inputPath, outputPath, sheetName, parseRows(rawRows));
  console.log(`appended rows to ${sheetName} -> ${outputPath}`);
}

async function appendTableColumn(
  inputPath: string,
  outputPath: string,
  tableName: string,
  columnName: string,
  rawValues: string | undefined
): Promise<void> {
  await appendWorkbookTableColumn(
    inputPath,
    outputPath,
    tableName,
    columnName,
    rawValues === undefined ? [] : parseRowValues(rawValues)
  );
  console.log(`appended ${tableName}[${columnName}] -> ${outputPath}`);
}

async function replaceTable(
  inputPath: string,
  outputPath: string,
  tableName: string,
  rawRows: string
): Promise<void> {
  await replaceWorkbookTableRows(inputPath, outputPath, tableName, parseRows(rawRows));
  console.log(`replaced ${tableName} rows -> ${outputPath}`);
}

async function renameTable(
  inputPath: string,
  outputPath: string,
  tableName: string,
  nextName: string
): Promise<void> {
  await renameWorkbookTable(inputPath, outputPath, tableName, nextName);
  console.log(`renamed ${tableName} to ${nextName} -> ${outputPath}`);
}

async function renameTableColumn(
  inputPath: string,
  outputPath: string,
  tableName: string,
  columnName: string,
  nextName: string
): Promise<void> {
  await renameWorkbookTableColumn(inputPath, outputPath, tableName, columnName, nextName);
  console.log(`renamed ${tableName}[${columnName}] to ${nextName} -> ${outputPath}`);
}

async function removeTableColumn(
  inputPath: string,
  outputPath: string,
  tableName: string,
  columnName: string
): Promise<void> {
  await removeRightmostWorkbookTableColumn(inputPath, outputPath, tableName, columnName);
  console.log(`removed ${tableName}[${columnName}] -> ${outputPath}`);
}

async function retargetChart(
  inputPath: string,
  outputPath: string,
  rawRetargets: string
): Promise<void> {
  await retargetWorkbookChartFormulasFile(inputPath, outputPath, parseChartRetargets(rawRetargets));
  console.log(`retargeted chart formulas -> ${outputPath}`);
}

async function retargetPivot(
  inputPath: string,
  outputPath: string,
  rawRetargets: string
): Promise<void> {
  await retargetWorkbookPivotCacheSourcesFile(
    inputPath,
    outputPath,
    parsePivotRetargets(rawRetargets)
  );
  console.log(`retargeted pivot cache sources -> ${outputPath}`);
}

function parseCliValue(value: string): CellInput {
  if (value.startsWith("=")) {
    return { formula: value };
  }

  if (value === "true") {
    return true;
  }

  if (value === "false") {
    return false;
  }

  if (/^-?\d+(\.\d+)?$/.test(value)) {
    return Number(value);
  }

  return value;
}

function parseRows(rawRows: string): CellInput[][] {
  const parsed: unknown = JSON.parse(rawRows);
  if (!Array.isArray(parsed) || !parsed.every((row) => Array.isArray(row))) {
    throw new Error("jsonRows must be an array of row arrays");
  }

  return parsed.map((row) => row.map((cell) => parseJsonCell(cell as unknown)));
}

function parseRowValues(rawValues: string): CellInput[] {
  const parsed: unknown = JSON.parse(rawValues);
  if (!Array.isArray(parsed)) {
    throw new Error("jsonValues must be an array");
  }

  return parsed.map((cell) => parseJsonCell(cell as unknown));
}

function parseStyle(rawStyle: string): WorkbookCellStyleInput {
  const parsed: unknown = JSON.parse(rawStyle);
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    throw new Error("jsonStyle must be an object");
  }

  return parsed as WorkbookCellStyleInput;
}

function parseChartRetargets(rawRetargets: string): ChartFormulaRetarget[] {
  const parsed: unknown = JSON.parse(rawRetargets);
  if (
    !Array.isArray(parsed) ||
    !parsed.every(
      (retarget) =>
        typeof retarget === "object" &&
        retarget !== null &&
        "from" in retarget &&
        typeof retarget.from === "string" &&
        "to" in retarget &&
        typeof retarget.to === "string"
    )
  ) {
    throw new Error("jsonRetargets must be an array of { from, to } objects");
  }

  return parsed as ChartFormulaRetarget[];
}

function parsePivotRetargets(rawRetargets: string): PivotCacheSourceRetarget[] {
  const parsed: unknown = JSON.parse(rawRetargets);
  if (!Array.isArray(parsed)) {
    throw new Error("jsonRetargets must be an array");
  }

  return parsed as PivotCacheSourceRetarget[];
}

function parseJsonCell(value: unknown): CellInput {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }

  if (
    typeof value === "object" &&
    value !== null &&
    "formula" in value &&
    typeof value.formula === "string"
  ) {
    return value as { formula: string };
  }

  throw new Error(`Unsupported table cell value ${JSON.stringify(value)}`);
}

const [command, ...args] = process.argv.slice(2) as [Command | undefined, ...string[]];

try {
  if (command === "inspect") {
    const [path] = args;
    if (path === undefined) {
      usage();
    }
    await inspect(path);
  } else if (command === "tables") {
    const [path] = args;
    if (path === undefined) {
      usage();
    }
    await tables(path);
  } else if (command === "formulas") {
    const [path] = args;
    if (path === undefined) {
      usage();
    }
    await formulas(path);
  } else if (command === "styles") {
    const [path] = args;
    if (path === undefined) {
      usage();
    }
    await styles(path);
  } else if (command === "validate") {
    const [path] = args;
    if (path === undefined) {
      usage();
    }
    await validate(path);
  } else if (command === "read-cell") {
    const [path, sheetName, address] = args;
    if (path === undefined || sheetName === undefined || address === undefined) {
      usage();
    }
    await readCellCommand(path, sheetName, address);
  } else if (command === "read-range") {
    const [path, sheetName, rangeRef] = args;
    if (path === undefined || sheetName === undefined || rangeRef === undefined) {
      usage();
    }
    await readRangeCommand(path, sheetName, rangeRef);
  } else if (command === "diff") {
    const [beforePath, afterPath] = args;
    if (beforePath === undefined || afterPath === undefined) {
      usage();
    }
    await diff(beforePath, afterPath);
  } else if (command === "patch") {
    const [inputPath, outputPath, sheetName, address, rawValue] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      address === undefined ||
      rawValue === undefined
    ) {
      usage();
    }
    await patch(inputPath, outputPath, sheetName, address, rawValue);
  } else if (command === "patch-range") {
    const [inputPath, outputPath, sheetName, startAddress, rawRows] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      startAddress === undefined ||
      rawRows === undefined
    ) {
      usage();
    }
    await patchRangeCommand(inputPath, outputPath, sheetName, startAddress, rawRows);
  } else if (command === "style-cell") {
    const [inputPath, outputPath, sheetName, address, rawStyle] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      address === undefined ||
      rawStyle === undefined
    ) {
      usage();
    }
    await styleCellCommand(inputPath, outputPath, sheetName, address, rawStyle);
  } else if (command === "append-rows") {
    const [inputPath, outputPath, sheetName, rawRows] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      rawRows === undefined
    ) {
      usage();
    }
    await appendRowsCommand(inputPath, outputPath, sheetName, rawRows);
  } else if (command === "append-table-column") {
    const [inputPath, outputPath, tableName, columnName, rawValues] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      tableName === undefined ||
      columnName === undefined
    ) {
      usage();
    }
    await appendTableColumn(inputPath, outputPath, tableName, columnName, rawValues);
  } else if (command === "replace-table") {
    const [inputPath, outputPath, tableName, rawRows] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      tableName === undefined ||
      rawRows === undefined
    ) {
      usage();
    }
    await replaceTable(inputPath, outputPath, tableName, rawRows);
  } else if (command === "rename-table") {
    const [inputPath, outputPath, tableName, nextName] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      tableName === undefined ||
      nextName === undefined
    ) {
      usage();
    }
    await renameTable(inputPath, outputPath, tableName, nextName);
  } else if (command === "rename-table-column") {
    const [inputPath, outputPath, tableName, columnName, nextName] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      tableName === undefined ||
      columnName === undefined ||
      nextName === undefined
    ) {
      usage();
    }
    await renameTableColumn(inputPath, outputPath, tableName, columnName, nextName);
  } else if (command === "remove-table-column") {
    const [inputPath, outputPath, tableName, columnName] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      tableName === undefined ||
      columnName === undefined
    ) {
      usage();
    }
    await removeTableColumn(inputPath, outputPath, tableName, columnName);
  } else if (command === "retarget-chart") {
    const [inputPath, outputPath, rawRetargets] = args;
    if (inputPath === undefined || outputPath === undefined || rawRetargets === undefined) {
      usage();
    }
    await retargetChart(inputPath, outputPath, rawRetargets);
  } else if (command === "retarget-pivot") {
    const [inputPath, outputPath, rawRetargets] = args;
    if (inputPath === undefined || outputPath === undefined || rawRetargets === undefined) {
      usage();
    }
    await retargetPivot(inputPath, outputPath, rawRetargets);
  } else {
    usage();
  }
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
}
