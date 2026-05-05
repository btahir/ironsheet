import { readFile as nodeReadFile, writeFile as nodeWriteFile } from "node:fs/promises";
import { deflateRawSync, inflateRawSync } from "node:zlib";
import {
  OoxmlPackage,
  Workbook,
  type CompressionAdapter,
  type CellInput
} from "../../core/src/index.ts";

export const nodeCompressionAdapter: CompressionAdapter = {
  inflateRaw(data) {
    return new Uint8Array(inflateRawSync(data));
  },
  deflateRaw(data) {
    return new Uint8Array(deflateRawSync(data));
  }
};

export async function openPackage(data: Uint8Array): Promise<OoxmlPackage> {
  return OoxmlPackage.open(data, nodeCompressionAdapter);
}

export async function openWorkbook(data: Uint8Array): Promise<Workbook> {
  return Workbook.fromPackage(await openPackage(data));
}

export async function readWorkbook(path: string): Promise<Workbook> {
  const data = await nodeReadFile(path);
  return openWorkbook(new Uint8Array(data));
}

export async function writeWorkbook(workbook: Workbook, path: string): Promise<void> {
  await nodeWriteFile(path, await workbook.write());
}

export async function patchWorkbookCell(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  address: string,
  value: CellInput
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.patchCell(sheetName, address, value);
  await writeWorkbook(workbook, outputPath);
}

export async function readWorkbookCell(
  inputPath: string,
  sheetName: string,
  address: string
): Promise<Awaited<ReturnType<Workbook["readCell"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.readCell(sheetName, address);
}
