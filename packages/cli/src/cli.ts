#!/usr/bin/env tsx
import { Buffer } from "node:buffer";
import process from "node:process";
import { readFile } from "node:fs/promises";
import { diffZipPackages } from "@ironsheet/core";
import {
  appendWorkbookRows,
  appendWorkbookTableColumn,
  deleteWorkbookAutoFilter,
  deleteWorkbookConditionalFormat,
  deleteWorkbookDataValidation,
  deleteWorkbookDefinedName,
  deleteWorkbookHyperlink,
  hideWorkbookSheet,
  inspectWorkbookStyles,
  listWorkbookAutoFilters,
  listWorkbookComments,
  listWorkbookConditionalFormats,
  listWorkbookDataValidations,
  listWorkbookFormulas,
  listWorkbookHyperlinks,
  listWorkbookImages,
  listWorkbookMergedCells,
  listWorkbookNamedRanges,
  listWorkbookTables,
  mergeWorkbookCells,
  patchWorkbookCell,
  patchWorkbookNamedRange,
  patchWorkbookRange,
  readWorkbook,
  readWorkbookCell,
  readWorkbookNamedRange,
  readWorkbookRange,
  removeRightmostWorkbookTableColumn,
  renameWorkbookSheet,
  renameWorkbookTable,
  renameWorkbookTableColumn,
  replaceWorkbookImageFile,
  replaceWorkbookTableRows,
  renderWorkbookTemplate,
  renderWorkbookTemplateSafely,
  retargetWorkbookChartFormulasFile,
  retargetWorkbookPivotCacheSourcesFile,
  setWorkbookAutoFilter,
  setWorkbookConditionalFormat,
  setWorkbookDataValidation,
  setWorkbookDefinedName,
  setWorkbookHyperlink,
  showWorkbookSheet,
  styleWorkbookCell,
  unmergeWorkbookCells,
  validateWorkbookFile
} from "@ironsheet/node";
import type {
  CellInput,
  ChartFormulaRetarget,
  PivotCacheSourceRetarget,
  WorkbookCellStyleInput,
  WorkbookSheetState,
  WorkbookTemplatePatch,
  WorksheetAutoFilter,
  WorksheetConditionalFormat,
  WorksheetDataValidation
} from "@ironsheet/core";

type Command =
  | "inspect"
  | "append-rows"
  | "append-table-column"
  | "auto-filters"
  | "comments"
  | "conditional-formats"
  | "delete-auto-filter"
  | "delete-conditional-format"
  | "data-validations"
  | "delete-data-validation"
  | "delete-defined-name"
  | "delete-hyperlink"
  | "formulas"
  | "hide-sheet"
  | "hyperlinks"
  | "images"
  | "patch"
  | "merge-cells"
  | "merged-cells"
  | "named-ranges"
  | "patch-range"
  | "patch-named-range"
  | "read-cell"
  | "read-named-range"
  | "read-range"
  | "rename-sheet"
  | "rename-table-column"
  | "rename-table"
  | "remove-table-column"
  | "replace-image"
  | "replace-table"
  | "render-template"
  | "render-template-safe"
  | "retarget-chart"
  | "retarget-pivot"
  | "set-auto-filter"
  | "set-conditional-format"
  | "set-data-validation"
  | "set-defined-name"
  | "set-hyperlink"
  | "show-sheet"
  | "style-cell"
  | "styles"
  | "tables"
  | "unmerge-cells"
  | "validate"
  | "diff";

function usage(): never {
  console.error(`usage:
  npm run cli -- inspect <workbook.xlsx>
  npm run cli -- tables <workbook.xlsx>
  npm run cli -- formulas <workbook.xlsx>
  npm run cli -- auto-filters <workbook.xlsx> [sheet]
  npm run cli -- comments <workbook.xlsx> [sheet]
  npm run cli -- conditional-formats <workbook.xlsx> [sheet]
  npm run cli -- data-validations <workbook.xlsx> [sheet]
  npm run cli -- hyperlinks <workbook.xlsx> [sheet]
  npm run cli -- images <workbook.xlsx> [sheet]
  npm run cli -- merged-cells <workbook.xlsx> [sheet]
  npm run cli -- named-ranges <workbook.xlsx> [name]
  npm run cli -- styles <workbook.xlsx>
  npm run cli -- validate <workbook.xlsx>
  npm run cli -- read-cell <workbook.xlsx> <sheet> <cell>
  npm run cli -- read-range <workbook.xlsx> <sheet> <range>
  npm run cli -- read-named-range <workbook.xlsx> <name> [sheet]
  npm run cli -- patch <input.xlsx> <output.xlsx> <sheet> <cell> <value>
  npm run cli -- style-cell <input.xlsx> <output.xlsx> <sheet> <cell> <jsonStyle>
  npm run cli -- patch-range <input.xlsx> <output.xlsx> <sheet> <startCell> <jsonRows>
  npm run cli -- patch-named-range <input.xlsx> <output.xlsx> <name> <jsonRows> [sheet]
  npm run cli -- append-rows <input.xlsx> <output.xlsx> <sheet> <jsonRows>
  npm run cli -- append-table-column <input.xlsx> <output.xlsx> <table> <column> [jsonValues]
  npm run cli -- set-auto-filter <input.xlsx> <output.xlsx> <sheet> <jsonAutoFilter>
  npm run cli -- delete-auto-filter <input.xlsx> <output.xlsx> <sheet>
  npm run cli -- replace-image <input.xlsx> <output.xlsx> <imagePartName> <imageFile>
  npm run cli -- render-template <input.xlsx> <output.xlsx> <jsonPatch|@patch.json>
  npm run cli -- render-template-safe <input.xlsx> <output.xlsx> <jsonPatch|@patch.json>
  npm run cli -- set-conditional-format <input.xlsx> <output.xlsx> <sheet> <jsonConditionalFormat>
  npm run cli -- delete-conditional-format <input.xlsx> <output.xlsx> <sheet> <sqref>
  npm run cli -- set-data-validation <input.xlsx> <output.xlsx> <sheet> <jsonValidation>
  npm run cli -- delete-data-validation <input.xlsx> <output.xlsx> <sheet> <sqref>
  npm run cli -- set-defined-name <input.xlsx> <output.xlsx> <name> <formula> [jsonOptions]
  npm run cli -- delete-defined-name <input.xlsx> <output.xlsx> <name> [jsonOptions]
  npm run cli -- hide-sheet <input.xlsx> <output.xlsx> <sheet> [hidden|veryHidden]
  npm run cli -- show-sheet <input.xlsx> <output.xlsx> <sheet>
  npm run cli -- merge-cells <input.xlsx> <output.xlsx> <sheet> <range>
  npm run cli -- unmerge-cells <input.xlsx> <output.xlsx> <sheet> <range>
  npm run cli -- set-hyperlink <input.xlsx> <output.xlsx> <sheet> <ref> <target> [jsonOptions]
  npm run cli -- delete-hyperlink <input.xlsx> <output.xlsx> <sheet> <ref>
  npm run cli -- rename-sheet <input.xlsx> <output.xlsx> <sheet> <newName>
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

async function autoFilters(path: string, sheetName: string | undefined): Promise<void> {
  console.log(JSON.stringify(await listWorkbookAutoFilters(path, sheetName), null, 2));
}

async function comments(path: string, sheetName: string | undefined): Promise<void> {
  console.log(JSON.stringify(await listWorkbookComments(path, sheetName), null, 2));
}

async function conditionalFormats(path: string, sheetName: string | undefined): Promise<void> {
  console.log(JSON.stringify(await listWorkbookConditionalFormats(path, sheetName), null, 2));
}

async function dataValidations(path: string, sheetName: string | undefined): Promise<void> {
  console.log(JSON.stringify(await listWorkbookDataValidations(path, sheetName), null, 2));
}

async function hyperlinks(path: string, sheetName: string | undefined): Promise<void> {
  console.log(JSON.stringify(await listWorkbookHyperlinks(path, sheetName), null, 2));
}

async function images(path: string, sheetName: string | undefined): Promise<void> {
  console.log(JSON.stringify(await listWorkbookImages(path, sheetName), null, 2));
}

async function mergedCells(path: string, sheetName: string | undefined): Promise<void> {
  console.log(JSON.stringify(await listWorkbookMergedCells(path, sheetName), null, 2));
}

async function namedRanges(path: string, name: string | undefined): Promise<void> {
  console.log(JSON.stringify(await listWorkbookNamedRanges(path, name), null, 2));
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

async function readNamedRangeCommand(
  path: string,
  name: string,
  sheetName: string | undefined
): Promise<void> {
  console.log(
    JSON.stringify(
      await readWorkbookNamedRange(path, name, sheetName === undefined ? {} : { sheetName }),
      null,
      2
    )
  );
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

async function patchNamedRangeCommand(
  inputPath: string,
  outputPath: string,
  name: string,
  rawRows: string,
  sheetName: string | undefined
): Promise<void> {
  await patchWorkbookNamedRange(
    inputPath,
    outputPath,
    name,
    parseRows(rawRows),
    sheetName === undefined ? {} : { sheetName }
  );
  console.log(`patched named range ${name} -> ${outputPath}`);
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

async function setDefinedName(
  inputPath: string,
  outputPath: string,
  name: string,
  text: string,
  rawOptions: string | undefined
): Promise<void> {
  await setWorkbookDefinedName(
    inputPath,
    outputPath,
    name,
    text,
    parseDefinedNameOptions(rawOptions)
  );
  console.log(`set defined name ${name} -> ${outputPath}`);
}

async function deleteDefinedName(
  inputPath: string,
  outputPath: string,
  name: string,
  rawOptions: string | undefined
): Promise<void> {
  const deleted = await deleteWorkbookDefinedName(
    inputPath,
    outputPath,
    name,
    parseDefinedNameDeleteOptions(rawOptions)
  );
  console.log(`${deleted ? "deleted" : "did not find"} defined name ${name} -> ${outputPath}`);
}

async function setAutoFilter(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  rawAutoFilter: string
): Promise<void> {
  const autoFilter = parseAutoFilter(rawAutoFilter);
  await setWorkbookAutoFilter(inputPath, outputPath, sheetName, autoFilter);
  console.log(`set auto filter ${sheetName}!${autoFilter.ref} -> ${outputPath}`);
}

async function deleteAutoFilter(
  inputPath: string,
  outputPath: string,
  sheetName: string
): Promise<void> {
  const deleted = await deleteWorkbookAutoFilter(inputPath, outputPath, sheetName);
  console.log(`${deleted ? "deleted" : "did not find"} auto filter ${sheetName} -> ${outputPath}`);
}

async function setConditionalFormat(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  rawConditionalFormat: string
): Promise<void> {
  const conditionalFormat = parseConditionalFormat(rawConditionalFormat);
  await setWorkbookConditionalFormat(inputPath, outputPath, sheetName, conditionalFormat);
  console.log(`set conditional format ${sheetName}!${conditionalFormat.sqref} -> ${outputPath}`);
}

async function deleteConditionalFormat(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  sqref: string
): Promise<void> {
  const deleted = await deleteWorkbookConditionalFormat(inputPath, outputPath, sheetName, sqref);
  console.log(
    `${deleted ? "deleted" : "did not find"} conditional format ${sheetName}!${sqref} -> ${outputPath}`
  );
}

async function setDataValidation(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  rawValidation: string
): Promise<void> {
  const dataValidation = parseDataValidation(rawValidation);
  await setWorkbookDataValidation(inputPath, outputPath, sheetName, dataValidation);
  console.log(`set data validation ${sheetName}!${dataValidation.sqref} -> ${outputPath}`);
}

async function deleteDataValidation(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  sqref: string
): Promise<void> {
  const deleted = await deleteWorkbookDataValidation(inputPath, outputPath, sheetName, sqref);
  console.log(
    `${deleted ? "deleted" : "did not find"} data validation ${sheetName}!${sqref} -> ${outputPath}`
  );
}

async function mergeCells(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  ref: string
): Promise<void> {
  await mergeWorkbookCells(inputPath, outputPath, sheetName, ref);
  console.log(`merged ${sheetName}!${ref} -> ${outputPath}`);
}

async function unmergeCells(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  ref: string
): Promise<void> {
  const unmerged = await unmergeWorkbookCells(inputPath, outputPath, sheetName, ref);
  console.log(`${unmerged ? "unmerged" : "did not find"} ${sheetName}!${ref} -> ${outputPath}`);
}

async function setHyperlink(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  ref: string,
  target: string,
  rawOptions: string | undefined
): Promise<void> {
  await setWorkbookHyperlink(
    inputPath,
    outputPath,
    sheetName,
    ref,
    target,
    parseHyperlinkOptions(rawOptions)
  );
  console.log(`set hyperlink ${sheetName}!${ref} -> ${outputPath}`);
}

async function deleteHyperlink(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  ref: string
): Promise<void> {
  const deleted = await deleteWorkbookHyperlink(inputPath, outputPath, sheetName, ref);
  console.log(
    `${deleted ? "deleted" : "did not find"} hyperlink ${sheetName}!${ref} -> ${outputPath}`
  );
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

async function replaceImage(
  inputPath: string,
  outputPath: string,
  imagePartName: string,
  imagePath: string
): Promise<void> {
  await replaceWorkbookImageFile(inputPath, outputPath, imagePartName, imagePath);
  console.log(`replaced ${imagePartName} from ${imagePath} -> ${outputPath}`);
}

async function renderTemplate(
  inputPath: string,
  outputPath: string,
  rawPatch: string
): Promise<void> {
  const result = await renderWorkbookTemplate(
    inputPath,
    outputPath,
    await parseTemplatePatch(rawPatch)
  );
  console.log(JSON.stringify(result, null, 2));
}

async function renderTemplateSafe(
  inputPath: string,
  outputPath: string,
  rawPatch: string
): Promise<void> {
  const result = await renderWorkbookTemplateSafely(
    inputPath,
    outputPath,
    await parseTemplatePatch(rawPatch)
  );
  console.log(JSON.stringify(result, null, 2));

  if (!result.wrote) {
    process.exitCode = 1;
  }
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

async function renameSheet(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  nextName: string
): Promise<void> {
  await renameWorkbookSheet(inputPath, outputPath, sheetName, nextName);
  console.log(`renamed ${sheetName} to ${nextName} -> ${outputPath}`);
}

async function hideSheet(
  inputPath: string,
  outputPath: string,
  sheetName: string,
  rawState: string | undefined
): Promise<void> {
  const state = parseSheetState(rawState);
  await hideWorkbookSheet(inputPath, outputPath, sheetName, state);
  console.log(`hid ${sheetName} as ${state} -> ${outputPath}`);
}

async function showSheet(inputPath: string, outputPath: string, sheetName: string): Promise<void> {
  await showWorkbookSheet(inputPath, outputPath, sheetName);
  console.log(`showed ${sheetName} -> ${outputPath}`);
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

  return parseJsonRows(parsed);
}

function parseJsonRows(rows: unknown[]): CellInput[][] {
  if (!rows.every((row) => Array.isArray(row))) {
    throw new Error("rows must be an array of row arrays");
  }

  return rows.map((row) => row.map((cell) => parseJsonCell(cell as unknown)));
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

function parseAutoFilter(rawAutoFilter: string): WorksheetAutoFilter {
  const parsed: unknown = JSON.parse(rawAutoFilter);
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    throw new Error("jsonAutoFilter must be an object");
  }

  if (!("ref" in parsed) || typeof parsed.ref !== "string") {
    throw new Error("jsonAutoFilter must include a string ref");
  }

  return parsed as WorksheetAutoFilter;
}

async function parseTemplatePatch(rawPatch: string): Promise<WorkbookTemplatePatch> {
  const parsed: unknown = JSON.parse(await readJsonArgument(rawPatch));
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    throw new Error("jsonPatch must be an object");
  }

  const patch = parsed as Record<string, unknown>;
  return {
    ...parseTemplateCells(patch.cells),
    ...(await parseTemplateImages(patch.images)),
    ...parseTemplateNames(patch.names),
    ...parseTemplateRanges(patch.ranges),
    ...parseTemplateTables(patch.tables)
  };
}

async function readJsonArgument(raw: string): Promise<string> {
  if (!raw.startsWith("@")) {
    return raw;
  }

  return readFile(raw.slice(1), "utf8");
}

function parseTemplateCells(cells: unknown): Pick<WorkbookTemplatePatch, "cells"> {
  if (cells === undefined) {
    return {};
  }

  if (!Array.isArray(cells)) {
    throw new Error("jsonPatch.cells must be an array");
  }

  return {
    cells: cells.map((cell) => {
      if (typeof cell !== "object" || cell === null || Array.isArray(cell)) {
        throw new Error("jsonPatch.cells entries must be objects");
      }

      const candidate = cell as Record<string, unknown>;
      if (
        typeof candidate.sheetName !== "string" ||
        typeof candidate.address !== "string" ||
        !("value" in candidate)
      ) {
        throw new Error("jsonPatch.cells entries need sheetName, address, and value");
      }

      return {
        sheetName: candidate.sheetName,
        address: candidate.address,
        value: parseJsonCell(candidate.value)
      };
    })
  };
}

async function parseTemplateImages(
  images: unknown
): Promise<Pick<WorkbookTemplatePatch, "images">> {
  if (images === undefined) {
    return {};
  }

  if (!Array.isArray(images)) {
    throw new Error("jsonPatch.images must be an array");
  }

  return {
    images: await Promise.all(
      images.map(async (image) => {
        if (typeof image !== "object" || image === null || Array.isArray(image)) {
          throw new Error("jsonPatch.images entries must be objects");
        }

        const candidate = image as Record<string, unknown>;
        if (typeof candidate.imagePartName !== "string") {
          throw new Error("jsonPatch.images entries need imagePartName");
        }

        if (typeof candidate.path === "string") {
          return {
            imagePartName: candidate.imagePartName,
            data: new Uint8Array(await readFile(candidate.path))
          };
        }

        if (typeof candidate.dataBase64 === "string") {
          return {
            imagePartName: candidate.imagePartName,
            data: new Uint8Array(Buffer.from(candidate.dataBase64, "base64"))
          };
        }

        throw new Error("jsonPatch.images entries need path or dataBase64");
      })
    )
  };
}

function parseTemplateNames(names: unknown): Pick<WorkbookTemplatePatch, "names"> {
  if (names === undefined) {
    return {};
  }

  if (!Array.isArray(names)) {
    throw new Error("jsonPatch.names must be an array");
  }

  return {
    names: names.map((name) => {
      if (typeof name !== "object" || name === null || Array.isArray(name)) {
        throw new Error("jsonPatch.names entries must be objects");
      }

      const candidate = name as Record<string, unknown>;
      if (typeof candidate.name !== "string" || !Array.isArray(candidate.values)) {
        throw new Error("jsonPatch.names entries need name and values");
      }

      if (candidate.sheetName !== undefined && typeof candidate.sheetName !== "string") {
        throw new Error("jsonPatch.names sheetName must be a string when provided");
      }

      if (
        candidate.allowOutsideRange !== undefined &&
        typeof candidate.allowOutsideRange !== "boolean"
      ) {
        throw new Error("jsonPatch.names allowOutsideRange must be a boolean when provided");
      }

      return {
        name: candidate.name,
        values: parseJsonRows(candidate.values),
        ...(candidate.allowOutsideRange === undefined
          ? {}
          : { allowOutsideRange: candidate.allowOutsideRange }),
        ...(candidate.sheetName === undefined ? {} : { sheetName: candidate.sheetName })
      };
    })
  };
}

function parseTemplateRanges(ranges: unknown): Pick<WorkbookTemplatePatch, "ranges"> {
  if (ranges === undefined) {
    return {};
  }

  if (!Array.isArray(ranges)) {
    throw new Error("jsonPatch.ranges must be an array");
  }

  return {
    ranges: ranges.map((range) => {
      if (typeof range !== "object" || range === null || Array.isArray(range)) {
        throw new Error("jsonPatch.ranges entries must be objects");
      }

      const candidate = range as Record<string, unknown>;
      if (
        typeof candidate.sheetName !== "string" ||
        typeof candidate.startAddress !== "string" ||
        !Array.isArray(candidate.values)
      ) {
        throw new Error("jsonPatch.ranges entries need sheetName, startAddress, and values");
      }

      return {
        sheetName: candidate.sheetName,
        startAddress: candidate.startAddress,
        values: parseJsonRows(candidate.values)
      };
    })
  };
}

function parseTemplateTables(tables: unknown): Pick<WorkbookTemplatePatch, "tables"> {
  if (tables === undefined) {
    return {};
  }

  if (!Array.isArray(tables)) {
    throw new Error("jsonPatch.tables must be an array");
  }

  return {
    tables: tables.map((table) => {
      if (typeof table !== "object" || table === null || Array.isArray(table)) {
        throw new Error("jsonPatch.tables entries must be objects");
      }

      const candidate = table as Record<string, unknown>;
      if (typeof candidate.tableName !== "string" || !Array.isArray(candidate.rows)) {
        throw new Error("jsonPatch.tables entries need tableName and rows");
      }

      return {
        tableName: candidate.tableName,
        rows: parseJsonRows(candidate.rows)
      };
    })
  };
}

function parseConditionalFormat(rawConditionalFormat: string): WorksheetConditionalFormat {
  const parsed: unknown = JSON.parse(rawConditionalFormat);
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    throw new Error("jsonConditionalFormat must be an object");
  }

  if (!("sqref" in parsed) || typeof parsed.sqref !== "string") {
    throw new Error("jsonConditionalFormat must include a string sqref");
  }

  if (!("rules" in parsed) || !Array.isArray(parsed.rules)) {
    throw new Error("jsonConditionalFormat must include a rules array");
  }

  return parsed as WorksheetConditionalFormat;
}

function parseDataValidation(rawValidation: string): WorksheetDataValidation {
  const parsed: unknown = JSON.parse(rawValidation);
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    throw new Error("jsonValidation must be an object");
  }

  if (!("sqref" in parsed) || typeof parsed.sqref !== "string") {
    throw new Error("jsonValidation must include a string sqref");
  }

  return parsed as WorksheetDataValidation;
}

function parseDefinedNameOptions(rawOptions: string | undefined): {
  comment?: string;
  hidden?: boolean;
  sheetName?: string;
} {
  if (rawOptions === undefined) {
    return {};
  }

  const parsed: unknown = JSON.parse(rawOptions);
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    throw new Error("jsonOptions must be an object");
  }

  const options = parsed as Record<string, unknown>;
  return {
    ...(typeof options.comment === "string" ? { comment: options.comment } : {}),
    ...(typeof options.hidden === "boolean" ? { hidden: options.hidden } : {}),
    ...(typeof options.sheetName === "string" ? { sheetName: options.sheetName } : {})
  };
}

function parseDefinedNameDeleteOptions(rawOptions: string | undefined): { sheetName?: string } {
  const options = parseDefinedNameOptions(rawOptions);
  return options.sheetName === undefined ? {} : { sheetName: options.sheetName };
}

function parseHyperlinkOptions(rawOptions: string | undefined): {
  display?: string;
  tooltip?: string;
} {
  if (rawOptions === undefined) {
    return {};
  }

  const parsed: unknown = JSON.parse(rawOptions);
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    throw new Error("jsonOptions must be an object");
  }

  const options = parsed as Record<string, unknown>;
  return {
    ...(typeof options.display === "string" ? { display: options.display } : {}),
    ...(typeof options.tooltip === "string" ? { tooltip: options.tooltip } : {})
  };
}

function parseSheetState(rawState: string | undefined): WorkbookSheetState {
  if (rawState === undefined || rawState === "hidden") {
    return "hidden";
  }

  if (rawState === "veryHidden") {
    return "veryHidden";
  }

  throw new Error("sheet state must be hidden or veryHidden");
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
  } else if (command === "auto-filters") {
    const [path, sheetName] = args;
    if (path === undefined) {
      usage();
    }
    await autoFilters(path, sheetName);
  } else if (command === "comments") {
    const [path, sheetName] = args;
    if (path === undefined) {
      usage();
    }
    await comments(path, sheetName);
  } else if (command === "conditional-formats") {
    const [path, sheetName] = args;
    if (path === undefined) {
      usage();
    }
    await conditionalFormats(path, sheetName);
  } else if (command === "data-validations") {
    const [path, sheetName] = args;
    if (path === undefined) {
      usage();
    }
    await dataValidations(path, sheetName);
  } else if (command === "hyperlinks") {
    const [path, sheetName] = args;
    if (path === undefined) {
      usage();
    }
    await hyperlinks(path, sheetName);
  } else if (command === "images") {
    const [path, sheetName] = args;
    if (path === undefined) {
      usage();
    }
    await images(path, sheetName);
  } else if (command === "merged-cells") {
    const [path, sheetName] = args;
    if (path === undefined) {
      usage();
    }
    await mergedCells(path, sheetName);
  } else if (command === "named-ranges") {
    const [path, name] = args;
    if (path === undefined) {
      usage();
    }
    await namedRanges(path, name);
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
  } else if (command === "read-named-range") {
    const [path, name, sheetName] = args;
    if (path === undefined || name === undefined) {
      usage();
    }
    await readNamedRangeCommand(path, name, sheetName);
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
  } else if (command === "patch-named-range") {
    const [inputPath, outputPath, name, rawRows, sheetName] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      name === undefined ||
      rawRows === undefined
    ) {
      usage();
    }
    await patchNamedRangeCommand(inputPath, outputPath, name, rawRows, sheetName);
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
  } else if (command === "set-auto-filter") {
    const [inputPath, outputPath, sheetName, rawAutoFilter] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      rawAutoFilter === undefined
    ) {
      usage();
    }
    await setAutoFilter(inputPath, outputPath, sheetName, rawAutoFilter);
  } else if (command === "delete-auto-filter") {
    const [inputPath, outputPath, sheetName] = args;
    if (inputPath === undefined || outputPath === undefined || sheetName === undefined) {
      usage();
    }
    await deleteAutoFilter(inputPath, outputPath, sheetName);
  } else if (command === "set-conditional-format") {
    const [inputPath, outputPath, sheetName, rawConditionalFormat] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      rawConditionalFormat === undefined
    ) {
      usage();
    }
    await setConditionalFormat(inputPath, outputPath, sheetName, rawConditionalFormat);
  } else if (command === "delete-conditional-format") {
    const [inputPath, outputPath, sheetName, sqref] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      sqref === undefined
    ) {
      usage();
    }
    await deleteConditionalFormat(inputPath, outputPath, sheetName, sqref);
  } else if (command === "set-data-validation") {
    const [inputPath, outputPath, sheetName, rawValidation] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      rawValidation === undefined
    ) {
      usage();
    }
    await setDataValidation(inputPath, outputPath, sheetName, rawValidation);
  } else if (command === "delete-data-validation") {
    const [inputPath, outputPath, sheetName, sqref] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      sqref === undefined
    ) {
      usage();
    }
    await deleteDataValidation(inputPath, outputPath, sheetName, sqref);
  } else if (command === "set-defined-name") {
    const [inputPath, outputPath, name, text, rawOptions] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      name === undefined ||
      text === undefined
    ) {
      usage();
    }
    await setDefinedName(inputPath, outputPath, name, text, rawOptions);
  } else if (command === "delete-defined-name") {
    const [inputPath, outputPath, name, rawOptions] = args;
    if (inputPath === undefined || outputPath === undefined || name === undefined) {
      usage();
    }
    await deleteDefinedName(inputPath, outputPath, name, rawOptions);
  } else if (command === "hide-sheet") {
    const [inputPath, outputPath, sheetName, rawState] = args;
    if (inputPath === undefined || outputPath === undefined || sheetName === undefined) {
      usage();
    }
    await hideSheet(inputPath, outputPath, sheetName, rawState);
  } else if (command === "show-sheet") {
    const [inputPath, outputPath, sheetName] = args;
    if (inputPath === undefined || outputPath === undefined || sheetName === undefined) {
      usage();
    }
    await showSheet(inputPath, outputPath, sheetName);
  } else if (command === "merge-cells") {
    const [inputPath, outputPath, sheetName, ref] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      ref === undefined
    ) {
      usage();
    }
    await mergeCells(inputPath, outputPath, sheetName, ref);
  } else if (command === "unmerge-cells") {
    const [inputPath, outputPath, sheetName, ref] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      ref === undefined
    ) {
      usage();
    }
    await unmergeCells(inputPath, outputPath, sheetName, ref);
  } else if (command === "set-hyperlink") {
    const [inputPath, outputPath, sheetName, ref, target, rawOptions] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      ref === undefined ||
      target === undefined
    ) {
      usage();
    }
    await setHyperlink(inputPath, outputPath, sheetName, ref, target, rawOptions);
  } else if (command === "delete-hyperlink") {
    const [inputPath, outputPath, sheetName, ref] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      ref === undefined
    ) {
      usage();
    }
    await deleteHyperlink(inputPath, outputPath, sheetName, ref);
  } else if (command === "replace-image") {
    const [inputPath, outputPath, imagePartName, imagePath] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      imagePartName === undefined ||
      imagePath === undefined
    ) {
      usage();
    }
    await replaceImage(inputPath, outputPath, imagePartName, imagePath);
  } else if (command === "render-template") {
    const [inputPath, outputPath, rawPatch] = args;
    if (inputPath === undefined || outputPath === undefined || rawPatch === undefined) {
      usage();
    }
    await renderTemplate(inputPath, outputPath, rawPatch);
  } else if (command === "render-template-safe") {
    const [inputPath, outputPath, rawPatch] = args;
    if (inputPath === undefined || outputPath === undefined || rawPatch === undefined) {
      usage();
    }
    await renderTemplateSafe(inputPath, outputPath, rawPatch);
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
  } else if (command === "rename-sheet") {
    const [inputPath, outputPath, sheetName, nextName] = args;
    if (
      inputPath === undefined ||
      outputPath === undefined ||
      sheetName === undefined ||
      nextName === undefined
    ) {
      usage();
    }
    await renameSheet(inputPath, outputPath, sheetName, nextName);
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
