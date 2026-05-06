import { readFile as nodeReadFile, writeFile as nodeWriteFile } from "node:fs/promises";
import { deflateRawSync, inflateRawSync } from "node:zlib";
import {
  OoxmlPackage,
  Workbook,
  type CompressionAdapter,
  type CellInput,
  type CellPatch,
  type ChartFormulaRetarget,
  type PivotCacheSourceRetarget,
  type WorkbookCellStyleInput
} from "@ironsheet/core";

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

export async function patchWorkbookCells(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  patches: CellPatch[]
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.patchCells(sheetName, patches);
  await writeWorkbook(workbook, outputPath);
}

export async function patchWorkbookRange(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  startAddress: string,
  values: CellInput[][]
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.patchRange(sheetName, startAddress, values);
  await writeWorkbook(workbook, outputPath);
}

export async function styleWorkbookCell(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  address: string,
  style: WorkbookCellStyleInput
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.styleCell(sheetName, address, style);
  await writeWorkbook(workbook, outputPath);
}

export async function appendWorkbookRows(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  rows: CellInput[][]
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.appendRows(sheetName, rows);
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

export async function readWorkbookRange(
  inputPath: string,
  sheetName: string,
  rangeRef: string
): Promise<Awaited<ReturnType<Workbook["readRange"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.readRange(sheetName, rangeRef);
}

export async function listWorkbookFormulas(
  inputPath: string
): Promise<Awaited<ReturnType<Workbook["formulas"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.formulas();
}

export async function listWorkbookTables(
  inputPath: string
): Promise<Awaited<ReturnType<Workbook["tables"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.tables();
}

export async function inspectWorkbookStyles(
  inputPath: string
): Promise<Awaited<ReturnType<Workbook["styles"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.styles();
}

export async function validateWorkbookFile(
  inputPath: string
): Promise<Awaited<ReturnType<Workbook["validate"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.validate();
}

export async function replaceWorkbookTableRows(
  inputPath: string,
  outputPath: string,
  tableName: string,
  rows: CellInput[][]
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.replaceTableRows(tableName, rows);
  await writeWorkbook(workbook, outputPath);
}

export async function appendWorkbookTableColumn(
  inputPath: string,
  outputPath: string,
  tableName: string,
  columnName: string,
  values: CellInput[] = []
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.appendTableColumn(tableName, columnName, values);
  await writeWorkbook(workbook, outputPath);
}

export async function renameWorkbookTable(
  inputPath: string,
  outputPath: string,
  tableName: string,
  nextName: string
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.renameTable(tableName, nextName);
  await writeWorkbook(workbook, outputPath);
}

export async function renameWorkbookSheet(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  nextName: string
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.renameSheet(sheetName, nextName);
  await writeWorkbook(workbook, outputPath);
}

export async function removeRightmostWorkbookTableColumn(
  inputPath: string,
  outputPath: string,
  tableName: string,
  columnName: string
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.removeRightmostTableColumn(tableName, columnName);
  await writeWorkbook(workbook, outputPath);
}

export async function renameWorkbookTableColumn(
  inputPath: string,
  outputPath: string,
  tableName: string,
  columnName: string,
  nextName: string
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.renameTableColumn(tableName, columnName, nextName);
  await writeWorkbook(workbook, outputPath);
}

export async function retargetWorkbookChartFormulasFile(
  inputPath: string,
  outputPath: string,
  retargets: ChartFormulaRetarget[]
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.retargetChartFormulas(retargets);
  await writeWorkbook(workbook, outputPath);
}

export async function retargetWorkbookPivotCacheSourcesFile(
  inputPath: string,
  outputPath: string,
  retargets: PivotCacheSourceRetarget[]
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.retargetPivotCacheSources(retargets);
  await writeWorkbook(workbook, outputPath);
}
