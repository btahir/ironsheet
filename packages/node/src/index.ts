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
  type WorkbookCellStyleInput,
  type WorkbookSheetState,
  type WorksheetConditionalFormat,
  type WorksheetDataValidation
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

export async function listWorkbookHyperlinks(
  inputPath: string,
  sheetName?: string
): Promise<Awaited<ReturnType<Workbook["hyperlinks"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.hyperlinks(sheetName);
}

export async function listWorkbookMergedCells(
  inputPath: string,
  sheetName?: string
): Promise<Awaited<ReturnType<Workbook["mergedCells"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.mergedCells(sheetName);
}

export async function listWorkbookDataValidations(
  inputPath: string,
  sheetName?: string
): Promise<Awaited<ReturnType<Workbook["dataValidations"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.dataValidations(sheetName);
}

export async function listWorkbookConditionalFormats(
  inputPath: string,
  sheetName?: string
): Promise<Awaited<ReturnType<Workbook["conditionalFormats"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.conditionalFormats(sheetName);
}

export async function listWorkbookComments(
  inputPath: string,
  sheetName?: string
): Promise<Awaited<ReturnType<Workbook["comments"]>>> {
  const workbook = await readWorkbook(inputPath);
  return workbook.comments(sheetName);
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

export async function setWorkbookDefinedName(
  inputPath: string,
  outputPath: string,
  name: string,
  text: string,
  options: Parameters<Workbook["setDefinedName"]>[2] = {}
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.setDefinedName(name, text, options);
  await writeWorkbook(workbook, outputPath);
}

export async function deleteWorkbookDefinedName(
  inputPath: string,
  outputPath: string,
  name: string,
  options: Parameters<Workbook["deleteDefinedName"]>[1] = {}
): Promise<boolean> {
  const workbook = await readWorkbook(inputPath);
  const deleted = await workbook.deleteDefinedName(name, options);
  await writeWorkbook(workbook, outputPath);
  return deleted;
}

export async function mergeWorkbookCells(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  ref: string
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.mergeCells(sheetName, ref);
  await writeWorkbook(workbook, outputPath);
}

export async function unmergeWorkbookCells(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  ref: string
): Promise<boolean> {
  const workbook = await readWorkbook(inputPath);
  const unmerged = await workbook.unmergeCells(sheetName, ref);
  await writeWorkbook(workbook, outputPath);
  return unmerged;
}

export async function setWorkbookDataValidation(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  dataValidation: WorksheetDataValidation
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.setDataValidation(sheetName, dataValidation);
  await writeWorkbook(workbook, outputPath);
}

export async function deleteWorkbookDataValidation(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  sqref: string
): Promise<boolean> {
  const workbook = await readWorkbook(inputPath);
  const deleted = await workbook.deleteDataValidation(sheetName, sqref);
  await writeWorkbook(workbook, outputPath);
  return deleted;
}

export async function setWorkbookConditionalFormat(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  conditionalFormat: WorksheetConditionalFormat
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.setConditionalFormat(sheetName, conditionalFormat);
  await writeWorkbook(workbook, outputPath);
}

export async function deleteWorkbookConditionalFormat(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  sqref: string
): Promise<boolean> {
  const workbook = await readWorkbook(inputPath);
  const deleted = await workbook.deleteConditionalFormat(sheetName, sqref);
  await writeWorkbook(workbook, outputPath);
  return deleted;
}

export async function setWorkbookHyperlink(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  ref: string,
  target: string,
  options: Parameters<Workbook["setHyperlink"]>[3] = {}
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.setHyperlink(sheetName, ref, target, options);
  await writeWorkbook(workbook, outputPath);
}

export async function deleteWorkbookHyperlink(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  ref: string
): Promise<boolean> {
  const workbook = await readWorkbook(inputPath);
  const deleted = await workbook.deleteHyperlink(sheetName, ref);
  await writeWorkbook(workbook, outputPath);
  return deleted;
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

export async function hideWorkbookSheet(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  state: WorkbookSheetState = "hidden"
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.hideSheet(sheetName, state);
  await writeWorkbook(workbook, outputPath);
}

export async function showWorkbookSheet(
  inputPath: string,
  outputPath: string,
  sheetName: string
): Promise<void> {
  const workbook = await readWorkbook(inputPath);
  await workbook.showSheet(sheetName);
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
