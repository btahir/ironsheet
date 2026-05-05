#!/usr/bin/env tsx
import process from "node:process";
import { readFile } from "node:fs/promises";
import { diffZipPackages } from "../../core/src/index.ts";
import { patchWorkbookCell, readWorkbook, readWorkbookCell } from "../../node/src/index.ts";
import type { CellInput } from "../../core/src/index.ts";

type Command = "inspect" | "patch" | "read-cell" | "diff";

function usage(): never {
  console.error(`usage:
  npm run cli -- inspect <workbook.xlsx>
  npm run cli -- read-cell <workbook.xlsx> <sheet> <cell>
  npm run cli -- patch <input.xlsx> <output.xlsx> <sheet> <cell> <value>
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

async function readCellCommand(path: string, sheetName: string, address: string): Promise<void> {
  console.log(JSON.stringify(await readWorkbookCell(path, sheetName, address), null, 2));
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

const [command, ...args] = process.argv.slice(2) as [Command | undefined, ...string[]];

try {
  if (command === "inspect") {
    const [path] = args;
    if (path === undefined) {
      usage();
    }
    await inspect(path);
  } else if (command === "read-cell") {
    const [path, sheetName, address] = args;
    if (path === undefined || sheetName === undefined || address === undefined) {
      usage();
    }
    await readCellCommand(path, sheetName, address);
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
  } else {
    usage();
  }
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
}
