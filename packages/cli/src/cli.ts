#!/usr/bin/env tsx
import process from "node:process";
import { readFile } from "node:fs/promises";
import { diffZipPackages } from "../../core/src/index.ts";
import {
  appendWorkbookRows,
  inspectWorkbookStyles,
  listWorkbookFormulas,
  listWorkbookTables,
  patchWorkbookCell,
  patchWorkbookRange,
  readWorkbook,
  readWorkbookCell,
  readWorkbookRange,
  renameWorkbookTable,
  renameWorkbookTableColumn,
  replaceWorkbookTableRows,
  validateWorkbookFile
} from "../../node/src/index.ts";
import type { CellInput } from "../../core/src/index.ts";

type Command =
  | "inspect"
  | "append-rows"
  | "formulas"
  | "patch"
  | "patch-range"
  | "read-cell"
  | "read-range"
  | "rename-table-column"
  | "rename-table"
  | "replace-table"
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
  npm run cli -- patch-range <input.xlsx> <output.xlsx> <sheet> <startCell> <jsonRows>
  npm run cli -- append-rows <input.xlsx> <output.xlsx> <sheet> <jsonRows>
  npm run cli -- rename-table <input.xlsx> <output.xlsx> <table> <newName>
  npm run cli -- rename-table-column <input.xlsx> <output.xlsx> <table> <column> <newName>
  npm run cli -- replace-table <input.xlsx> <output.xlsx> <table> <jsonRows>
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
  console.log(JSON.stringify(await validateWorkbookFile(path), null, 2));
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
  } else {
    usage();
  }
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
}
