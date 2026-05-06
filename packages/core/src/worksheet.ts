import {
  compareCellAddresses,
  formatCellAddress,
  parseCellAddress,
  parseCellRange,
  type CellRange
} from "./address.ts";
import { WorksheetError } from "./errors.ts";
import {
  decodeXml,
  escapeXmlAttribute,
  escapeXmlText,
  findElementCloseStart,
  findElementEnd,
  findFirstStartTag,
  findStartTags,
  streamXmlElements,
  tokenizeXmlChunks,
  xmlTokenRawText
} from "./xml.ts";

export type CellPrimitive = string | number | boolean | Date | null;

export type FormulaValue = {
  formula: string;
  result?: CellPrimitive;
};

export type CellInput = CellPrimitive | FormulaValue;

export type ReadCellResult = {
  address: string;
  value: CellPrimitive;
  formula?: string;
  styleId?: string;
};

export type ReadRangeResult = {
  range: string;
  cells: Array<Array<ReadCellResult | undefined>>;
};

export type CellPatch = {
  address: string;
  value: CellInput;
};

export type PatchCellResult = {
  xml: string;
  formulaChanged: boolean;
  affectedRanges: CellRange[];
};

type ExistingCell = {
  start: number;
  end: number;
  raw: string;
  styleId?: string;
};

type RowElement = {
  rowNumber: number;
  start: number;
  end: number;
  tag: ReturnType<typeof findStartTags>[number];
};

type RowTemplate = {
  attributes: Record<string, string>;
  stylesByColumn: Map<number, string>;
};

export type WorksheetRowXml = {
  attributes: Record<string, string>;
  raw: string;
  rowNumber?: number;
  start: number;
  end: number;
};

export type WorksheetRowReplacement = {
  rowNumber: number;
  xml: string;
};

export type WorksheetMergedCell = {
  ref: string;
};

export type MergeWorksheetCellsResult = {
  merged: boolean;
  merge: WorksheetMergedCell;
  xml: string;
};

export type UnmergeWorksheetCellsResult = {
  unmerged: boolean;
  xml: string;
};

export type WorksheetDataValidation = {
  sqref: string;
  allowBlank?: boolean;
  error?: string;
  errorStyle?: string;
  errorTitle?: string;
  formula1?: string;
  formula2?: string;
  operator?: string;
  prompt?: string;
  promptTitle?: string;
  showDropDown?: boolean;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
  type?: string;
};

export type SetWorksheetDataValidationResult = {
  dataValidation: WorksheetDataValidation;
  replaced: boolean;
  xml: string;
};

export type DeleteWorksheetDataValidationResult = {
  deleted: boolean;
  xml: string;
};

export type WorksheetConditionalFormatRule = {
  aboveAverage?: boolean;
  attributes?: Record<string, string | number | boolean>;
  bottom?: boolean;
  dxfId?: string;
  equalAverage?: boolean;
  formulas?: string[];
  operator?: string;
  percent?: boolean;
  priority?: string;
  rank?: string;
  rawXml?: string;
  stdDev?: string;
  stopIfTrue?: boolean;
  text?: string;
  timePeriod?: string;
  type?: string;
};

export type WorksheetConditionalFormat = {
  sqref: string;
  pivot?: boolean;
  rules: WorksheetConditionalFormatRule[];
};

export type SetWorksheetConditionalFormatResult = {
  conditionalFormat: WorksheetConditionalFormat;
  replaced: boolean;
  xml: string;
};

export type DeleteWorksheetConditionalFormatResult = {
  deleted: boolean;
  xml: string;
};

export type WorksheetAutoFilter = {
  ref: string;
  rawXml?: string;
};

export type SetWorksheetAutoFilterResult = {
  autoFilter: WorksheetAutoFilter;
  replaced: boolean;
  xml: string;
};

export type DeleteWorksheetAutoFilterResult = {
  deleted: boolean;
  xml: string;
};

export type WorksheetHyperlink = {
  ref: string;
  display?: string;
  location?: string;
  relationshipId?: string;
  tooltip?: string;
};

export type SetWorksheetHyperlinkResult = {
  hyperlink: WorksheetHyperlink;
  replacedRelationshipId?: string;
  xml: string;
};

export type DeleteWorksheetHyperlinkResult = {
  deleted: boolean;
  relationshipIds: string[];
  xml: string;
};

export type EnsureWorksheetDrawingResult = {
  inserted: boolean;
  xml: string;
};

export function readCell(
  xml: string,
  address: string,
  options: { sharedStrings?: string[] } = {}
): ReadCellResult | undefined {
  const parsedAddress = parseCellAddress(address);
  const existing = findCellElement(xml, parsedAddress.address);

  if (existing === undefined) {
    return undefined;
  }

  const result: ReadCellResult = {
    address: parsedAddress.address,
    value: readCellValue(existing.raw, options.sharedStrings ?? [])
  };

  if (existing.styleId !== undefined) {
    result.styleId = existing.styleId;
  }

  const formula = readTagText(existing.raw, "f");
  if (formula !== undefined) {
    result.formula = formula;
  }

  return result;
}

export function readRange(
  xml: string,
  rangeRef: string,
  options: { sharedStrings?: string[] } = {}
): ReadRangeResult {
  const range = parseCellRange(rangeRef);
  const cells: Array<Array<ReadCellResult | undefined>> = [];

  for (let row = range.start.row; row <= range.end.row; row += 1) {
    const rowCells: Array<ReadCellResult | undefined> = [];

    for (let column = range.start.column; column <= range.end.column; column += 1) {
      rowCells.push(readCell(xml, formatCellAddress(column, row), options));
    }

    cells.push(rowCells);
  }

  return {
    range: range.ref,
    cells
  };
}

export function patchCell(xml: string, address: string, value: CellInput): PatchCellResult {
  const parsedAddress = parseCellAddress(address);
  const existing = findCellElement(xml, parsedAddress.address);
  const prefix = inferWorksheetPrefix(xml);
  const existingFormula =
    existing === undefined ? false : findFirstStartTag(existing.raw, "f") !== undefined;
  const cellXml = createCellXml(
    parsedAddress.address,
    value,
    existing?.styleId === undefined ? { prefix } : { prefix, styleId: existing.styleId }
  );
  const formulaChanged = isFormulaValue(value) || existingFormula;
  const withCell =
    existing === undefined
      ? insertCell(xml, parsedAddress.address, parsedAddress.row, cellXml)
      : `${xml.slice(0, existing.start)}${cellXml}${xml.slice(existing.end)}`;

  return {
    xml: updateDimension(withCell, parsedAddress.address),
    formulaChanged,
    affectedRanges: [parseCellRange(parsedAddress.address)]
  };
}

export function patchCells(xml: string, patches: CellPatch[]): PatchCellResult {
  let nextXml = xml;
  let formulaChanged = false;
  const affectedRanges: CellRange[] = [];

  for (const patch of patches) {
    const result = patchCell(nextXml, patch.address, patch.value);
    nextXml = result.xml;
    formulaChanged = formulaChanged || result.formulaChanged;
    affectedRanges.push(...result.affectedRanges);
  }

  return {
    xml: nextXml,
    formulaChanged,
    affectedRanges
  };
}

export function patchRange(
  xml: string,
  startAddress: string,
  values: CellInput[][]
): PatchCellResult {
  const start = parseCellAddress(startAddress);
  const patches: CellPatch[] = [];

  for (const [rowIndex, row] of values.entries()) {
    for (const [columnIndex, value] of row.entries()) {
      patches.push({
        address: formatCellAddress(start.column + columnIndex, start.row + rowIndex),
        value
      });
    }
  }

  return patchCells(xml, patches);
}

export function applyCellStyle(xml: string, address: string, styleId: string): PatchCellResult {
  const parsedAddress = parseCellAddress(address);
  const existing = findCellElement(xml, parsedAddress.address);
  const styled =
    existing === undefined
      ? insertCell(
          xml,
          parsedAddress.address,
          parsedAddress.row,
          createCellXml(parsedAddress.address, null, {
            prefix: inferWorksheetPrefix(xml),
            styleId
          })
        )
      : `${xml.slice(0, existing.start)}${upsertCellStyle(existing.raw, styleId)}${xml.slice(existing.end)}`;

  return {
    xml: updateDimension(styled, parsedAddress.address),
    formulaChanged: false,
    affectedRanges: [parseCellRange(parsedAddress.address)]
  };
}

export function removeCellsInRange(
  xml: string,
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
): PatchCellResult {
  let nextXml = xml;
  const cells = findStartTags(xml, "c")
    .map((tag) => {
      const address = tag.attributes.r;
      if (address === undefined) {
        return undefined;
      }

      const parsed = parseCellAddress(address);
      if (
        parsed.row < range.startRow ||
        parsed.row > range.endRow ||
        parsed.column < range.startColumn ||
        parsed.column > range.endColumn
      ) {
        return undefined;
      }

      return {
        start: tag.start,
        end: tag.selfClosing ? tag.end : findElementEnd(xml, tag),
        address: parsed.address
      };
    })
    .filter((cell): cell is { start: number; end: number; address: string } => cell !== undefined)
    .sort((left, right) => right.start - left.start);

  for (const cell of cells) {
    nextXml = `${nextXml.slice(0, cell.start)}${nextXml.slice(cell.end)}`;
  }

  const affectedRef = `${formatCellAddress(range.startColumn, range.startRow)}:${formatCellAddress(range.endColumn, range.endRow)}`;
  return {
    xml: recalculateDimension(nextXml),
    formulaChanged: false,
    affectedRanges: [parseCellRange(affectedRef)]
  };
}

export function appendRows(
  xml: string,
  rows: CellInput[][],
  options: { startColumn?: number } = {}
): PatchCellResult {
  if (rows.length === 0) {
    return {
      xml,
      formulaChanged: false,
      affectedRanges: []
    };
  }

  const startColumn = options.startColumn ?? 1;
  const startRow = findMaxUsedRow(xml) + 1;
  const rowXml = createRowsXml(rows, { prefix: inferWorksheetPrefix(xml), startColumn, startRow });
  const inserted = insertRowsBeforeSheetDataClose(xml, rowXml);
  const lastRow = Math.max(startRow, startRow + rows.length - 1);
  const lastColumn = Math.max(
    startColumn,
    ...rows.map((row) => startColumn + Math.max(0, row.length - 1))
  );

  return {
    xml: updateDimension(inserted, formatCellAddress(lastColumn, lastRow)),
    formulaChanged: rows.some((row) => row.some(isFormulaValue)),
    affectedRanges: [
      parseCellRange(
        `${formatCellAddress(startColumn, startRow)}:${formatCellAddress(lastColumn, lastRow)}`
      )
    ]
  };
}

export function replaceRowsInRange(
  xml: string,
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  rows: CellInput[][],
  options: { preserveStyles?: boolean; trailingRows?: number } = {}
): string {
  const sheetData = findFirstStartTag(xml, "sheetData");
  if (sheetData === undefined) {
    throw new WorksheetError("Worksheet is missing sheetData");
  }

  findElementCloseStart(xml, sheetData);

  const preserveStyles = options.preserveStyles ?? true;
  const prefix = inferWorksheetPrefix(xml);
  const template = preserveStyles ? collectRowTemplate(xml, range) : undefined;
  const trailingRows = options.trailingRows ?? 0;
  const trailingRowElements = findRowElements(xml).filter(
    (row) => row.rowNumber > range.endRow && row.rowNumber <= range.endRow + trailingRows
  );
  const rowElements = findRowElements(xml)
    .filter(
      (row) => row.rowNumber >= range.startRow && row.rowNumber <= range.endRow + trailingRows
    )
    .slice()
    .reverse();

  let nextXml = xml;
  for (const row of rowElements) {
    nextXml = `${nextXml.slice(0, row.start)}${nextXml.slice(row.end)}`;
  }

  const insertionPoint = findRowInsertionPoint(nextXml, range.startRow);
  const rowXml = rows
    .map((row, rowIndex) => {
      const rowNumber = range.startRow + rowIndex;
      const cells = row
        .slice(0, range.endColumn - range.startColumn + 1)
        .map((cell, columnIndex) => {
          const column = range.startColumn + columnIndex;
          const styleId = template?.stylesByColumn.get(column);
          return createCellXml(
            formatCellAddress(column, rowNumber),
            cell,
            styleId === undefined ? { prefix } : { prefix, styleId }
          );
        })
        .join("");

      return `<${qualifiedName(prefix, "row")} r="${rowNumber}"${rowAttributesXml(template?.attributes)}>${cells}</${qualifiedName(prefix, "row")}>`;
    })
    .join("");
  const shiftedTrailingRows = trailingRowElements
    .map((row, index) =>
      shiftRowXml(xml.slice(row.start, row.end), {
        oldRow: row.rowNumber,
        newRow: range.startRow + rows.length + index,
        oldBodyStartRow: range.startRow,
        oldBodyEndRow: range.endRow,
        newBodyEndRow: range.startRow + rows.length - 1
      })
    )
    .join("");

  const updated = `${nextXml.slice(0, insertionPoint)}${rowXml}${shiftedTrailingRows}${nextXml.slice(insertionPoint)}`;
  return recalculateDimension(updated);
}

export function createCellXml(
  address: string,
  value: CellInput,
  options: { prefix?: string | undefined; styleId?: string } = {}
): string {
  const attributes = createCellAttributes(address, value, options);
  const c = qualifiedName(options.prefix, "c");

  if (isFormulaValue(value)) {
    const f = qualifiedName(options.prefix, "f");
    const result =
      value.result === undefined ? "" : createFormulaResultXml(value.result, options.prefix);
    return `<${c} ${attributes}><${f}>${escapeXmlText(value.formula.replace(/^=/, ""))}</${f}>${result}</${c}>`;
  }

  if (value === null) {
    return `<${c} ${attributes}/>`;
  }

  if (typeof value === "number") {
    const v = qualifiedName(options.prefix, "v");
    return `<${c} ${attributes}><${v}>${String(value)}</${v}></${c}>`;
  }

  if (typeof value === "boolean") {
    const v = qualifiedName(options.prefix, "v");
    return `<${c} ${attributes}><${v}>${value ? "1" : "0"}</${v}></${c}>`;
  }

  if (value instanceof Date) {
    const v = qualifiedName(options.prefix, "v");
    return `<${c} ${attributes}><${v}>${dateToExcelSerial(value)}</${v}></${c}>`;
  }

  const is = qualifiedName(options.prefix, "is");
  const t = qualifiedName(options.prefix, "t");
  return `<${c} ${attributes}><${is}><${t}>${escapeXmlText(value)}</${t}></${is}></${c}>`;
}

export function createRowsXml(
  rows: CellInput[][],
  options: { prefix?: string | undefined; startColumn?: number; startRow: number }
): string {
  const startColumn = options.startColumn ?? 1;
  const rowTag = qualifiedName(options.prefix, "row");
  return rows
    .map((row, rowIndex) => {
      const rowNumber = options.startRow + rowIndex;
      const cells = row
        .map((cell, columnIndex) =>
          createCellXml(formatCellAddress(startColumn + columnIndex, rowNumber), cell, {
            prefix: options.prefix
          })
        )
        .join("");

      return `<${rowTag} r="${rowNumber}">${cells}</${rowTag}>`;
    })
    .join("");
}

export async function* streamRowsXml(
  rows: Iterable<CellInput[]> | AsyncIterable<CellInput[]>,
  options: { startColumn?: number; startRow: number }
): AsyncGenerator<string> {
  const startColumn = options.startColumn ?? 1;
  let rowNumber = options.startRow;

  for await (const row of rows) {
    yield createRowsXml([row], { startColumn, startRow: rowNumber });
    rowNumber += 1;
  }
}

export async function* streamWorksheetRowsXml(
  chunks: Iterable<string> | AsyncIterable<string>
): AsyncGenerator<WorksheetRowXml> {
  for await (const row of streamXmlElements(chunks, "row")) {
    const rowNumber = Number.parseInt(row.tag.attributes.r ?? "", 10);
    yield {
      attributes: row.tag.attributes,
      raw: row.raw,
      ...(Number.isInteger(rowNumber) ? { rowNumber } : {}),
      start: row.start,
      end: row.end
    };
  }
}

export async function* streamReplaceWorksheetRowsXml(
  chunks: Iterable<string> | AsyncIterable<string>,
  replacements: Iterable<WorksheetRowReplacement>
): AsyncGenerator<string> {
  const replacementsByRow = new Map(
    [...replacements].map((replacement) => [replacement.rowNumber, replacement.xml])
  );
  let skippingRowDepth = 0;

  for await (const token of tokenizeXmlChunks(chunks)) {
    if (skippingRowDepth > 0) {
      if (token.kind === "start" && token.tag.localName === "row" && !token.tag.selfClosing) {
        skippingRowDepth += 1;
      } else if (token.kind === "end" && token.localName === "row") {
        skippingRowDepth -= 1;
      }
      continue;
    }

    if (token.kind !== "start" || token.tag.localName !== "row") {
      yield xmlTokenRawText(token);
      continue;
    }

    const rowNumber = Number.parseInt(token.tag.attributes.r ?? "", 10);
    const replacement = Number.isInteger(rowNumber) ? replacementsByRow.get(rowNumber) : undefined;
    if (replacement === undefined) {
      yield xmlTokenRawText(token);
      continue;
    }

    yield replacement;
    if (!token.tag.selfClosing) {
      skippingRowDepth = 1;
    }
  }
}

export function listWorksheetDataValidations(xml: string): WorksheetDataValidation[] {
  return findStartTags(xml, "dataValidation").map((tag) => dataValidationFromTag(xml, tag));
}

export function setWorksheetDataValidation(
  xml: string,
  dataValidation: WorksheetDataValidation
): SetWorksheetDataValidationResult {
  const normalized = normalizeDataValidation(dataValidation);
  const dataValidationXml = createDataValidationXml(xml, normalized);
  const existing = findMatchingDataValidation(xml, normalized.sqref);
  if (existing !== undefined) {
    const end = existing.selfClosing ? existing.end : findElementEnd(xml, existing);
    return {
      dataValidation: normalized,
      replaced: true,
      xml: `${xml.slice(0, existing.start)}${dataValidationXml}${xml.slice(end)}`
    };
  }

  const dataValidations = findFirstStartTag(xml, "dataValidations");
  if (dataValidations !== undefined) {
    if (dataValidations.selfClosing) {
      const opening = upsertTagAttribute(dataValidations.raw, "count", "1").replace(/\/>$/, ">");
      return {
        dataValidation: normalized,
        replaced: false,
        xml: `${xml.slice(0, dataValidations.start)}${opening}${dataValidationXml}</${dataValidations.name}>${xml.slice(dataValidations.end)}`
      };
    }

    const count = findStartTags(xml, "dataValidation").length + 1;
    const opening = upsertTagAttribute(dataValidations.raw, "count", String(count));
    const close = findElementCloseStart(xml, dataValidations);
    return {
      dataValidation: normalized,
      replaced: false,
      xml: `${xml.slice(0, dataValidations.start)}${opening}${xml.slice(dataValidations.end, close)}${dataValidationXml}${xml.slice(close)}`
    };
  }

  const insertOffset = dataValidationContainerInsertOffset(xml);
  const dataValidationsTag = qualifiedName(inferWorksheetPrefix(xml), "dataValidations");
  return {
    dataValidation: normalized,
    replaced: false,
    xml: `${xml.slice(0, insertOffset)}<${dataValidationsTag} count="1">${dataValidationXml}</${dataValidationsTag}>${xml.slice(insertOffset)}`
  };
}

export function deleteWorksheetDataValidation(
  xml: string,
  sqref: string
): DeleteWorksheetDataValidationResult {
  const normalizedSqref = normalizeSqref(sqref);
  const removals = findStartTags(xml, "dataValidation")
    .filter((tag) => normalizeSqref(tag.attributes.sqref ?? "") === normalizedSqref)
    .map((tag) => ({
      tag,
      end: tag.selfClosing ? tag.end : findElementEnd(xml, tag)
    }));

  if (removals.length === 0) {
    return { deleted: false, xml };
  }

  let nextXml = xml;
  for (const removal of removals.toReversed()) {
    nextXml = `${nextXml.slice(0, removal.tag.start)}${nextXml.slice(removal.end)}`;
  }

  const dataValidations = findFirstStartTag(nextXml, "dataValidations");
  if (dataValidations === undefined) {
    return { deleted: true, xml: nextXml };
  }

  const remaining = findStartTags(nextXml, "dataValidation").length;
  if (remaining === 0) {
    return {
      deleted: true,
      xml: `${nextXml.slice(0, dataValidations.start)}${nextXml.slice(findElementEnd(nextXml, dataValidations))}`
    };
  }

  const opening = upsertTagAttribute(dataValidations.raw, "count", String(remaining));
  return {
    deleted: true,
    xml: `${nextXml.slice(0, dataValidations.start)}${opening}${nextXml.slice(dataValidations.end)}`
  };
}

export function listWorksheetConditionalFormats(xml: string): WorksheetConditionalFormat[] {
  return findStartTags(xml, "conditionalFormatting").map((tag) =>
    conditionalFormatFromTag(xml, tag)
  );
}

export function setWorksheetConditionalFormat(
  xml: string,
  conditionalFormat: WorksheetConditionalFormat
): SetWorksheetConditionalFormatResult {
  const normalized = normalizeConditionalFormat(conditionalFormat);
  const conditionalFormatXml = createConditionalFormatXml(xml, normalized);
  const existing = findMatchingConditionalFormat(xml, normalized.sqref);
  if (existing !== undefined) {
    const end = existing.selfClosing ? existing.end : findElementEnd(xml, existing);
    return {
      conditionalFormat: normalized,
      replaced: true,
      xml: `${xml.slice(0, existing.start)}${conditionalFormatXml}${xml.slice(end)}`
    };
  }

  const insertOffset = conditionalFormatInsertOffset(xml);
  return {
    conditionalFormat: normalized,
    replaced: false,
    xml: `${xml.slice(0, insertOffset)}${conditionalFormatXml}${xml.slice(insertOffset)}`
  };
}

export function deleteWorksheetConditionalFormat(
  xml: string,
  sqref: string
): DeleteWorksheetConditionalFormatResult {
  const normalizedSqref = normalizeSqref(sqref, "Conditional formatting sqref");
  const removals = findStartTags(xml, "conditionalFormatting")
    .filter((tag) => normalizeSqref(tag.attributes.sqref ?? "") === normalizedSqref)
    .map((tag) => ({
      tag,
      end: tag.selfClosing ? tag.end : findElementEnd(xml, tag)
    }));

  if (removals.length === 0) {
    return { deleted: false, xml };
  }

  let nextXml = xml;
  for (const removal of removals.toReversed()) {
    nextXml = `${nextXml.slice(0, removal.tag.start)}${nextXml.slice(removal.end)}`;
  }

  return {
    deleted: true,
    xml: nextXml
  };
}

export function listWorksheetAutoFilters(xml: string): WorksheetAutoFilter[] {
  return findStartTags(xml, "autoFilter").map((tag) => autoFilterFromTag(xml, tag));
}

export function setWorksheetAutoFilter(
  xml: string,
  autoFilter: WorksheetAutoFilter
): SetWorksheetAutoFilterResult {
  const normalized = normalizeAutoFilter(autoFilter);
  const autoFilterXml = createAutoFilterXml(xml, normalized);
  const existing = findFirstStartTag(xml, "autoFilter");
  if (existing !== undefined) {
    const end = existing.selfClosing ? existing.end : findElementEnd(xml, existing);
    return {
      autoFilter: normalized,
      replaced: true,
      xml: `${xml.slice(0, existing.start)}${autoFilterXml}${xml.slice(end)}`
    };
  }

  const insertOffset = autoFilterInsertOffset(xml);
  return {
    autoFilter: normalized,
    replaced: false,
    xml: `${xml.slice(0, insertOffset)}${autoFilterXml}${xml.slice(insertOffset)}`
  };
}

export function deleteWorksheetAutoFilter(xml: string): DeleteWorksheetAutoFilterResult {
  const removals = findStartTags(xml, "autoFilter").map((tag) => ({
    tag,
    end: tag.selfClosing ? tag.end : findElementEnd(xml, tag)
  }));

  if (removals.length === 0) {
    return { deleted: false, xml };
  }

  let nextXml = xml;
  for (const removal of removals.toReversed()) {
    nextXml = `${nextXml.slice(0, removal.tag.start)}${nextXml.slice(removal.end)}`;
  }

  return {
    deleted: true,
    xml: nextXml
  };
}

export function listWorksheetMergedCells(xml: string): WorksheetMergedCell[] {
  return findStartTags(xml, "mergeCell").map((tag) => ({
    ref: normalizeMergeRef(tag.attributes.ref ?? "")
  }));
}

export function mergeWorksheetCells(xml: string, ref: string): MergeWorksheetCellsResult {
  const merge = { ref: normalizeMergeRef(ref) };
  const mergeRange = parseCellRange(merge.ref);
  if (
    mergeRange.start.column === mergeRange.end.column &&
    mergeRange.start.row === mergeRange.end.row
  ) {
    throw new WorksheetError(`Cannot merge a single cell range ${merge.ref}`);
  }

  const existingMerges = listWorksheetMergedCells(xml);
  if (existingMerges.some((existing) => existing.ref === merge.ref)) {
    return { merged: false, merge, xml };
  }

  const overlap = existingMerges.find((existing) =>
    cellRangesIntersect(parseCellRange(existing.ref), mergeRange)
  );
  if (overlap !== undefined) {
    throw new WorksheetError(`Merged range ${merge.ref} overlaps existing merge ${overlap.ref}`);
  }

  const mergeCellXml = createMergeCellXml(xml, merge);
  const mergeCells = findFirstStartTag(xml, "mergeCells");
  if (mergeCells !== undefined) {
    if (mergeCells.selfClosing) {
      const opening = upsertTagAttribute(mergeCells.raw, "count", "1").replace(/\/>$/, ">");
      return {
        merged: true,
        merge,
        xml: `${xml.slice(0, mergeCells.start)}${opening}${mergeCellXml}</${mergeCells.name}>${xml.slice(mergeCells.end)}`
      };
    }

    const opening = upsertTagAttribute(mergeCells.raw, "count", String(existingMerges.length + 1));
    const close = findElementCloseStart(xml, mergeCells);
    return {
      merged: true,
      merge,
      xml: `${xml.slice(0, mergeCells.start)}${opening}${xml.slice(mergeCells.end, close)}${mergeCellXml}${xml.slice(close)}`
    };
  }

  const insertOffset = mergeContainerInsertOffset(xml);
  const mergeCellsTag = qualifiedName(inferWorksheetPrefix(xml), "mergeCells");
  return {
    merged: true,
    merge,
    xml: `${xml.slice(0, insertOffset)}<${mergeCellsTag} count="1">${mergeCellXml}</${mergeCellsTag}>${xml.slice(insertOffset)}`
  };
}

export function unmergeWorksheetCells(xml: string, ref: string): UnmergeWorksheetCellsResult {
  const normalizedRef = normalizeMergeRef(ref);
  const removals = findStartTags(xml, "mergeCell")
    .filter((tag) => normalizeMergeRef(tag.attributes.ref ?? "") === normalizedRef)
    .map((tag) => ({
      tag,
      end: tag.selfClosing ? tag.end : findElementEnd(xml, tag)
    }));

  if (removals.length === 0) {
    return { unmerged: false, xml };
  }

  let nextXml = xml;
  for (const removal of removals.toReversed()) {
    nextXml = `${nextXml.slice(0, removal.tag.start)}${nextXml.slice(removal.end)}`;
  }

  const mergeCells = findFirstStartTag(nextXml, "mergeCells");
  if (mergeCells === undefined) {
    return { unmerged: true, xml: nextXml };
  }

  const remaining = findStartTags(nextXml, "mergeCell").length;
  if (remaining === 0) {
    return {
      unmerged: true,
      xml: `${nextXml.slice(0, mergeCells.start)}${nextXml.slice(findElementEnd(nextXml, mergeCells))}`
    };
  }

  const opening = upsertTagAttribute(mergeCells.raw, "count", String(remaining));
  return {
    unmerged: true,
    xml: `${nextXml.slice(0, mergeCells.start)}${opening}${nextXml.slice(mergeCells.end)}`
  };
}

export function listWorksheetHyperlinks(xml: string): WorksheetHyperlink[] {
  return findStartTags(xml, "hyperlink").map((tag) => {
    const hyperlink: WorksheetHyperlink = {
      ref: normalizeHyperlinkRef(tag.attributes.ref ?? "")
    };

    if (tag.attributes.display !== undefined) {
      hyperlink.display = tag.attributes.display;
    }
    if (tag.attributes.location !== undefined) {
      hyperlink.location = tag.attributes.location;
    }
    if (tag.attributes["r:id"] !== undefined) {
      hyperlink.relationshipId = tag.attributes["r:id"];
    }
    if (tag.attributes.tooltip !== undefined) {
      hyperlink.tooltip = tag.attributes.tooltip;
    }

    return hyperlink;
  });
}

export function setWorksheetHyperlink(
  xml: string,
  hyperlink: WorksheetHyperlink
): SetWorksheetHyperlinkResult {
  const normalized: WorksheetHyperlink = {
    ref: normalizeHyperlinkRef(hyperlink.ref),
    ...(hyperlink.display === undefined ? {} : { display: hyperlink.display }),
    ...(hyperlink.location === undefined ? {} : { location: hyperlink.location }),
    ...(hyperlink.relationshipId === undefined ? {} : { relationshipId: hyperlink.relationshipId }),
    ...(hyperlink.tooltip === undefined ? {} : { tooltip: hyperlink.tooltip })
  };
  const nextXml =
    normalized.relationshipId === undefined ? xml : ensureWorksheetRelationshipNamespace(xml);
  const hyperlinkXml = createHyperlinkXml(nextXml, normalized);
  const existing = findMatchingHyperlink(nextXml, normalized.ref);
  if (existing !== undefined) {
    const end = existing.selfClosing ? existing.end : findElementEnd(nextXml, existing);
    const replacedRelationshipId = existing.attributes["r:id"];
    return {
      hyperlink: normalized,
      ...(replacedRelationshipId === undefined ? {} : { replacedRelationshipId }),
      xml: `${nextXml.slice(0, existing.start)}${hyperlinkXml}${nextXml.slice(end)}`
    };
  }

  const hyperlinks = findFirstStartTag(nextXml, "hyperlinks");
  if (hyperlinks !== undefined) {
    if (hyperlinks.selfClosing) {
      const hyperlinksTag = qualifiedName(inferWorksheetPrefix(nextXml), "hyperlinks");
      return {
        hyperlink: normalized,
        xml: `${nextXml.slice(0, hyperlinks.start)}<${hyperlinksTag}>${hyperlinkXml}</${hyperlinksTag}>${nextXml.slice(hyperlinks.end)}`
      };
    }

    const insertOffset = findElementCloseStart(nextXml, hyperlinks);
    return {
      hyperlink: normalized,
      xml: `${nextXml.slice(0, insertOffset)}${hyperlinkXml}${nextXml.slice(insertOffset)}`
    };
  }

  const insertOffset = hyperlinkContainerInsertOffset(nextXml);
  const hyperlinksTag = qualifiedName(inferWorksheetPrefix(nextXml), "hyperlinks");
  return {
    hyperlink: normalized,
    xml: `${nextXml.slice(0, insertOffset)}<${hyperlinksTag}>${hyperlinkXml}</${hyperlinksTag}>${nextXml.slice(insertOffset)}`
  };
}

export function deleteWorksheetHyperlink(xml: string, ref: string): DeleteWorksheetHyperlinkResult {
  const normalizedRef = normalizeHyperlinkRef(ref);
  const removals = findStartTags(xml, "hyperlink")
    .filter((tag) => normalizeHyperlinkRef(tag.attributes.ref ?? "") === normalizedRef)
    .map((tag) => ({
      tag,
      end: tag.selfClosing ? tag.end : findElementEnd(xml, tag),
      relationshipId: tag.attributes["r:id"]
    }));

  if (removals.length === 0) {
    return { deleted: false, relationshipIds: [], xml };
  }

  let nextXml = xml;
  for (const removal of removals.toReversed()) {
    nextXml = `${nextXml.slice(0, removal.tag.start)}${nextXml.slice(removal.end)}`;
  }

  const hyperlinks = findFirstStartTag(nextXml, "hyperlinks");
  if (hyperlinks !== undefined && !hyperlinks.selfClosing) {
    const body = nextXml.slice(hyperlinks.end, findElementCloseStart(nextXml, hyperlinks));
    if (findStartTags(body, "hyperlink").length === 0) {
      nextXml = `${nextXml.slice(0, hyperlinks.start)}${nextXml.slice(findElementEnd(nextXml, hyperlinks))}`;
    }
  }

  return {
    deleted: true,
    relationshipIds: removals
      .map((removal) => removal.relationshipId)
      .filter((relationshipId): relationshipId is string => relationshipId !== undefined),
    xml: nextXml
  };
}

export function ensureWorksheetDrawing(
  xml: string,
  relationshipId: string
): EnsureWorksheetDrawingResult {
  const nextXml = ensureWorksheetRelationshipNamespace(xml);
  if (findFirstStartTag(nextXml, "drawing") !== undefined) {
    return { inserted: false, xml: nextXml };
  }

  const drawingXml = `<${qualifiedName(inferWorksheetPrefix(nextXml), "drawing")} r:id="${escapeXmlAttribute(relationshipId)}"/>`;
  const insertOffset = drawingInsertOffset(nextXml);
  return {
    inserted: true,
    xml: `${nextXml.slice(0, insertOffset)}${drawingXml}${nextXml.slice(insertOffset)}`
  };
}

function findCellElement(xml: string, address: string): ExistingCell | undefined {
  const tags = findStartTags(xml, "c");
  const target = address.toUpperCase();

  for (const tag of tags) {
    if ((tag.attributes.r ?? "").toUpperCase() !== target) {
      continue;
    }

    if (tag.selfClosing) {
      return existingCell(xml, tag.start, tag.end, tag.attributes.s);
    }

    return existingCell(xml, tag.start, findElementEnd(xml, tag), tag.attributes.s);
  }

  return undefined;
}

function existingCell(
  xml: string,
  start: number,
  end: number,
  styleId: string | undefined
): ExistingCell {
  if (styleId === undefined) {
    return { start, end, raw: xml.slice(start, end) };
  }

  return { start, end, raw: xml.slice(start, end), styleId };
}

function insertCell(xml: string, address: string, rowNumber: number, cellXml: string): string {
  const row = findRowElement(xml, rowNumber);

  if (row === undefined) {
    return insertRow(xml, rowNumber, cellXml);
  }

  const rowXml = xml.slice(row.start, row.end);
  const cells = findStartTags(rowXml, "c")
    .map((tag) => ({
      tag,
      address: tag.attributes.r
    }))
    .filter(
      (cell): cell is { tag: typeof cell.tag; address: string } => cell.address !== undefined
    );

  for (const cell of cells) {
    if (compareCellAddresses(address, cell.address) < 0) {
      const insertionPoint = row.start + cell.tag.start;
      return `${xml.slice(0, insertionPoint)}${cellXml}${xml.slice(insertionPoint)}`;
    }
  }

  const close = findElementCloseStart(xml, row.tag);

  return `${xml.slice(0, close)}${cellXml}${xml.slice(close)}`;
}

function insertRow(xml: string, rowNumber: number, cellXml: string): string {
  const insertionPoint = findRowInsertionPoint(xml, rowNumber);
  const rowTag = qualifiedName(inferWorksheetPrefix(xml), "row");
  const rowXml = `<${rowTag} r="${rowNumber}">${cellXml}</${rowTag}>`;
  return `${xml.slice(0, insertionPoint)}${rowXml}${xml.slice(insertionPoint)}`;
}

function insertRowsBeforeSheetDataClose(xml: string, rowXml: string): string {
  const sheetData = findFirstStartTag(xml, "sheetData");
  if (sheetData === undefined) {
    throw new WorksheetError("Worksheet is missing sheetData");
  }

  const close = findElementCloseStart(xml, sheetData);

  return `${xml.slice(0, close)}${rowXml}${xml.slice(close)}`;
}

function findRowElement(xml: string, rowNumber: number): RowElement | undefined {
  const target = String(rowNumber);
  return findRowElements(xml).find((row) => row.tag.attributes.r === target);
}

function findRowElements(xml: string): RowElement[] {
  return findStartTags(xml, "row").map((row) => {
    const rowNumber = Number.parseInt(row.attributes.r ?? "", 10);
    if (!Number.isInteger(rowNumber)) {
      throw new WorksheetError("Row is missing a numeric r attribute");
    }
    if (row.selfClosing) {
      return { rowNumber, start: row.start, end: row.end, tag: row };
    }

    return { rowNumber, start: row.start, end: findElementEnd(xml, row), tag: row };
  });
}

function findRowInsertionPoint(xml: string, rowNumber: number): number {
  const nextRow = findRowElements(xml).find((row) => row.rowNumber > rowNumber);
  if (nextRow !== undefined) {
    return nextRow.start;
  }

  const sheetData = findFirstStartTag(xml, "sheetData");
  if (sheetData === undefined) {
    throw new WorksheetError("Worksheet is missing sheetData");
  }

  return findElementCloseStart(xml, sheetData);
}

function findMaxUsedRow(xml: string): number {
  const rowNumbers = findRowElements(xml).map((row) => row.rowNumber);
  const cellRows = findStartTags(xml, "c")
    .map((tag) => tag.attributes.r)
    .filter((address): address is string => address !== undefined)
    .map((address) => parseCellAddress(address).row);

  return Math.max(0, ...rowNumbers, ...cellRows);
}

function collectRowTemplate(
  xml: string,
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
): RowTemplate | undefined {
  const templateRow = findRowElements(xml).find(
    (row) => row.rowNumber >= range.startRow && row.rowNumber <= range.endRow
  );
  if (templateRow === undefined) {
    return undefined;
  }

  const attributes = Object.fromEntries(
    Object.entries(templateRow.tag.attributes).filter(([name]) => name !== "r")
  );

  const stylesByColumn = new Map<number, string>();
  const rowXml = xml.slice(templateRow.start, templateRow.end);
  for (const cell of findStartTags(rowXml, "c")) {
    const address = cell.attributes.r;
    const styleId = cell.attributes.s;
    if (address === undefined || styleId === undefined) {
      continue;
    }

    const column = parseCellAddress(address).column;
    if (column >= range.startColumn && column <= range.endColumn) {
      stylesByColumn.set(column, styleId);
    }
  }

  return {
    attributes,
    stylesByColumn
  };
}

function rowAttributesXml(attributes: Record<string, string> | undefined): string {
  if (attributes === undefined) {
    return "";
  }

  return Object.entries(attributes)
    .map(([name, value]) => ` ${name}="${escapeXmlAttribute(value)}"`)
    .join("");
}

function shiftRowXml(
  rowXml: string,
  options: {
    oldRow: number;
    newRow: number;
    oldBodyStartRow: number;
    oldBodyEndRow: number;
    newBodyEndRow: number;
  }
): string {
  let nextXml = rowXml.replace(/\s(r)=["']([1-9][0-9]*)["']/, ` r="${options.newRow}"`);
  nextXml = nextXml.replace(
    /\s(r)=["']([A-Z]+)([1-9][0-9]*)["']/g,
    (_match, attributeName: string, columnLabel: string, rowLabel: string) => {
      const row = Number.parseInt(rowLabel, 10);
      return ` ${attributeName}="${columnLabel}${row === options.oldRow ? options.newRow : row}"`;
    }
  );

  if (options.oldBodyEndRow >= options.oldBodyStartRow) {
    const rangePattern = new RegExp(
      `(\\$?[A-Z]+\\$?)${options.oldBodyStartRow}:(\\$?[A-Z]+\\$?)${options.oldBodyEndRow}`,
      "g"
    );
    nextXml = nextXml.replace(
      rangePattern,
      `$1${options.oldBodyStartRow}:$2${Math.max(options.oldBodyStartRow, options.newBodyEndRow)}`
    );
  }

  return nextXml;
}

function updateDimension(xml: string, address: string): string {
  const dimension = findFirstStartTag(xml, "dimension");
  if (dimension === undefined) {
    return xml;
  }

  const ref = dimension.attributes.ref;
  if (ref === undefined) {
    return xml;
  }

  const expanded = expandRange(ref, address);
  const replacement = upsertRefAttribute(dimension.raw, expanded);

  return `${xml.slice(0, dimension.start)}${replacement}${xml.slice(dimension.end)}`;
}

function recalculateDimension(xml: string): string {
  const dimension = findFirstStartTag(xml, "dimension");
  if (dimension === undefined) {
    return xml;
  }

  const cells = findStartTags(xml, "c")
    .map((tag) => tag.attributes.r)
    .filter((address): address is string => address !== undefined)
    .map((address) => parseCellAddress(address));
  const ref =
    cells.length === 0
      ? "A1"
      : `${formatCellAddress(Math.min(...cells.map((cell) => cell.column)), Math.min(...cells.map((cell) => cell.row)))}:${formatCellAddress(Math.max(...cells.map((cell) => cell.column)), Math.max(...cells.map((cell) => cell.row)))}`;
  const replacement = upsertRefAttribute(dimension.raw, ref);

  return `${xml.slice(0, dimension.start)}${replacement}${xml.slice(dimension.end)}`;
}

function expandRange(ref: string, address: string): string {
  const parts = ref.includes(":") ? ref.split(":") : [ref, ref];
  const [rawStart, rawEnd] = parts;
  if (rawStart === undefined || rawEnd === undefined) {
    throw new WorksheetError(`Invalid worksheet dimension ${ref}`);
  }

  const start = parseCellAddress(rawStart);
  const end = parseCellAddress(rawEnd);
  const next = parseCellAddress(address);

  const minColumn = Math.min(start.column, end.column, next.column);
  const minRow = Math.min(start.row, end.row, next.row);
  const maxColumn = Math.max(start.column, end.column, next.column);
  const maxRow = Math.max(start.row, end.row, next.row);

  return `${formatCellAddress(minColumn, minRow)}:${formatCellAddress(maxColumn, maxRow)}`;
}

function upsertRefAttribute(rawTag: string, ref: string): string {
  if (/\sref=(["']).*?\1/.test(rawTag)) {
    return rawTag.replace(/\sref=(["']).*?\1/, ` ref="${escapeXmlAttribute(ref)}"`);
  }

  const closing = rawTag.endsWith("/>") ? "/>" : ">";
  return `${rawTag.slice(0, -closing.length)} ref="${escapeXmlAttribute(ref)}"${closing}`;
}

function upsertCellStyle(cellXml: string, styleId: string): string {
  const cell = findFirstStartTag(cellXml, "c");
  if (cell === undefined) {
    throw new WorksheetError("Cell XML is missing c tag");
  }

  return `${cellXml.slice(0, cell.start)}${upsertTagAttribute(cell.raw, "s", styleId)}${cellXml.slice(cell.end)}`;
}

function upsertTagAttribute(rawTag: string, name: string, value: string): string {
  const escapedValue = escapeXmlAttribute(value);
  const pattern = new RegExp(`\\s${name}=(["']).*?\\1`);
  if (pattern.test(rawTag)) {
    return rawTag.replace(pattern, ` ${name}="${escapedValue}"`);
  }

  const closing = rawTag.endsWith("/>") ? "/>" : ">";
  return `${rawTag.slice(0, -closing.length)} ${name}="${escapedValue}"${closing}`;
}

function createFormulaResultXml(value: CellPrimitive, prefix: string | undefined): string {
  if (value === null) {
    return "";
  }

  const v = qualifiedName(prefix, "v");
  if (typeof value === "boolean") {
    return `<${v}>${value ? "1" : "0"}</${v}>`;
  }

  if (typeof value === "number") {
    return `<${v}>${String(value)}</${v}>`;
  }

  if (value instanceof Date) {
    return `<${v}>${dateToExcelSerial(value)}</${v}>`;
  }

  return `<${v}>${escapeXmlText(value)}</${v}>`;
}

function createCellAttributes(
  address: string,
  value: CellInput,
  options: { styleId?: string }
): string {
  const attributes: string[] = [`r="${escapeXmlAttribute(address)}"`];

  if (options.styleId !== undefined) {
    attributes.push(`s="${escapeXmlAttribute(options.styleId)}"`);
  }

  if (isFormulaValue(value)) {
    if (typeof value.result === "boolean") {
      attributes.push('t="b"');
    } else if (typeof value.result === "string") {
      attributes.push('t="str"');
    }
  } else if (typeof value === "boolean") {
    attributes.push('t="b"');
  } else if (typeof value === "string") {
    attributes.push('t="inlineStr"');
  }

  return attributes.join(" ");
}

function readCellValue(cellXml: string, sharedStrings: string[]): CellPrimitive {
  if (/t=(["'])inlineStr\1/.test(cellXml)) {
    return readTextRuns(cellXml).join("");
  }

  if (/t=(["'])s\1/.test(cellXml)) {
    const index = Number.parseInt(readTagText(cellXml, "v") ?? "", 10);
    return Number.isInteger(index) ? (sharedStrings[index] ?? null) : null;
  }

  if (/t=(["'])b\1/.test(cellXml)) {
    return readTagText(cellXml, "v") === "1";
  }

  if (/t=(["'])(?:str|e)\1/.test(cellXml)) {
    return readTagText(cellXml, "v") ?? "";
  }

  const value = readTagText(cellXml, "v");
  if (value === undefined) {
    return null;
  }

  const numericValue = Number(value);
  return Number.isFinite(numericValue) ? numericValue : value;
}

function readTextRuns(xml: string): string[] {
  return findStartTags(xml, "t").map((tag) => {
    if (tag.selfClosing) {
      return "";
    }

    return decodeXml(xml.slice(tag.end, findElementCloseStart(xml, tag)));
  });
}

function readTagText(xml: string, localName: string): string | undefined {
  const tag = findFirstStartTag(xml, localName);
  if (tag === undefined) {
    return undefined;
  }

  return decodeXml(xml.slice(tag.end, findElementCloseStart(xml, tag)));
}

function dateToExcelSerial(date: Date): number {
  const epoch = Date.UTC(1899, 11, 30);
  return (date.getTime() - epoch) / 86_400_000;
}

function inferWorksheetPrefix(xml: string): string | undefined {
  for (const localName of ["worksheet", "sheetData", "row", "c"]) {
    const tag = findFirstStartTag(xml, localName);
    if (tag === undefined) {
      continue;
    }

    const colon = tag.name.indexOf(":");
    return colon === -1 ? undefined : tag.name.slice(0, colon);
  }

  return undefined;
}

function normalizeHyperlinkRef(ref: string): string {
  return parseCellRange(ref).ref;
}

function normalizeSqref(sqref: string, label = "Data validation sqref"): string {
  const refs = sqref
    .trim()
    .split(/\s+/)
    .filter((ref) => ref.length > 0)
    .map((ref) => parseCellRange(ref.replaceAll("$", "")).ref);

  if (refs.length === 0) {
    throw new WorksheetError(`${label} cannot be empty`);
  }

  return refs.join(" ");
}

function normalizeDataValidation(dataValidation: WorksheetDataValidation): WorksheetDataValidation {
  return {
    sqref: normalizeSqref(dataValidation.sqref),
    ...(dataValidation.allowBlank === undefined ? {} : { allowBlank: dataValidation.allowBlank }),
    ...(dataValidation.error === undefined ? {} : { error: dataValidation.error }),
    ...(dataValidation.errorStyle === undefined ? {} : { errorStyle: dataValidation.errorStyle }),
    ...(dataValidation.errorTitle === undefined ? {} : { errorTitle: dataValidation.errorTitle }),
    ...(dataValidation.formula1 === undefined ? {} : { formula1: dataValidation.formula1 }),
    ...(dataValidation.formula2 === undefined ? {} : { formula2: dataValidation.formula2 }),
    ...(dataValidation.operator === undefined ? {} : { operator: dataValidation.operator }),
    ...(dataValidation.prompt === undefined ? {} : { prompt: dataValidation.prompt }),
    ...(dataValidation.promptTitle === undefined
      ? {}
      : { promptTitle: dataValidation.promptTitle }),
    ...(dataValidation.showDropDown === undefined
      ? {}
      : { showDropDown: dataValidation.showDropDown }),
    ...(dataValidation.showErrorMessage === undefined
      ? {}
      : { showErrorMessage: dataValidation.showErrorMessage }),
    ...(dataValidation.showInputMessage === undefined
      ? {}
      : { showInputMessage: dataValidation.showInputMessage }),
    ...(dataValidation.type === undefined ? {} : { type: dataValidation.type })
  };
}

function dataValidationFromTag(
  xml: string,
  tag: ReturnType<typeof findStartTags>[number]
): WorksheetDataValidation {
  const dataValidation: WorksheetDataValidation = {
    sqref: normalizeSqref(tag.attributes.sqref ?? "")
  };

  for (const [attribute, key] of [
    ["error", "error"],
    ["errorStyle", "errorStyle"],
    ["errorTitle", "errorTitle"],
    ["operator", "operator"],
    ["prompt", "prompt"],
    ["promptTitle", "promptTitle"],
    ["type", "type"]
  ] as const) {
    if (tag.attributes[attribute] !== undefined) {
      dataValidation[key] = tag.attributes[attribute];
    }
  }

  for (const [attribute, key] of [
    ["allowBlank", "allowBlank"],
    ["showDropDown", "showDropDown"],
    ["showErrorMessage", "showErrorMessage"],
    ["showInputMessage", "showInputMessage"]
  ] as const) {
    if (tag.attributes[attribute] !== undefined) {
      dataValidation[key] = xmlBoolean(tag.attributes[attribute]);
    }
  }

  const raw = xml.slice(tag.start, tag.selfClosing ? tag.end : findElementEnd(xml, tag));
  const formula1 = readTagText(raw, "formula1");
  if (formula1 !== undefined) {
    dataValidation.formula1 = formula1;
  }

  const formula2 = readTagText(raw, "formula2");
  if (formula2 !== undefined) {
    dataValidation.formula2 = formula2;
  }

  return dataValidation;
}

function createDataValidationXml(xml: string, dataValidation: WorksheetDataValidation): string {
  const prefix = inferWorksheetPrefix(xml);
  const tagName = qualifiedName(prefix, "dataValidation");
  const attributes = [
    `sqref="${escapeXmlAttribute(dataValidation.sqref)}"`,
    dataValidation.type === undefined
      ? undefined
      : `type="${escapeXmlAttribute(dataValidation.type)}"`,
    dataValidation.operator === undefined
      ? undefined
      : `operator="${escapeXmlAttribute(dataValidation.operator)}"`,
    dataValidation.allowBlank === undefined
      ? undefined
      : `allowBlank="${dataValidation.allowBlank ? "1" : "0"}"`,
    dataValidation.showDropDown === undefined
      ? undefined
      : `showDropDown="${dataValidation.showDropDown ? "1" : "0"}"`,
    dataValidation.showErrorMessage === undefined
      ? undefined
      : `showErrorMessage="${dataValidation.showErrorMessage ? "1" : "0"}"`,
    dataValidation.showInputMessage === undefined
      ? undefined
      : `showInputMessage="${dataValidation.showInputMessage ? "1" : "0"}"`,
    dataValidation.errorStyle === undefined
      ? undefined
      : `errorStyle="${escapeXmlAttribute(dataValidation.errorStyle)}"`,
    dataValidation.errorTitle === undefined
      ? undefined
      : `errorTitle="${escapeXmlAttribute(dataValidation.errorTitle)}"`,
    dataValidation.error === undefined
      ? undefined
      : `error="${escapeXmlAttribute(dataValidation.error)}"`,
    dataValidation.promptTitle === undefined
      ? undefined
      : `promptTitle="${escapeXmlAttribute(dataValidation.promptTitle)}"`,
    dataValidation.prompt === undefined
      ? undefined
      : `prompt="${escapeXmlAttribute(dataValidation.prompt)}"`
  ]
    .filter((attribute): attribute is string => attribute !== undefined)
    .join(" ");

  const formulaXml = [
    dataValidation.formula1 === undefined
      ? undefined
      : `<${qualifiedName(prefix, "formula1")}>${escapeXmlText(dataValidation.formula1)}</${qualifiedName(prefix, "formula1")}>`,
    dataValidation.formula2 === undefined
      ? undefined
      : `<${qualifiedName(prefix, "formula2")}>${escapeXmlText(dataValidation.formula2)}</${qualifiedName(prefix, "formula2")}>`
  ]
    .filter((formula): formula is string => formula !== undefined)
    .join("");

  return formulaXml.length === 0
    ? `<${tagName} ${attributes}/>`
    : `<${tagName} ${attributes}>${formulaXml}</${tagName}>`;
}

function findMatchingDataValidation(
  xml: string,
  sqref: string
): ReturnType<typeof findStartTags>[number] | undefined {
  return findStartTags(xml, "dataValidation").find(
    (tag) => normalizeSqref(tag.attributes.sqref ?? "") === sqref
  );
}

function dataValidationContainerInsertOffset(xml: string): number {
  for (const localName of [
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst"
  ]) {
    const tag = findFirstStartTag(xml, localName);
    if (tag !== undefined) {
      return tag.start;
    }
  }

  const worksheet = findFirstStartTag(xml, "worksheet");
  if (worksheet === undefined) {
    throw new WorksheetError("Worksheet is missing worksheet root");
  }

  return findElementCloseStart(xml, worksheet);
}

function normalizeAutoFilter(autoFilter: WorksheetAutoFilter): WorksheetAutoFilter {
  return {
    ref: parseCellRange(autoFilter.ref.replaceAll("$", "")).ref,
    ...(autoFilter.rawXml === undefined ? {} : { rawXml: autoFilter.rawXml })
  };
}

function autoFilterFromTag(
  xml: string,
  tag: ReturnType<typeof findStartTags>[number]
): WorksheetAutoFilter {
  return {
    ref: parseCellRange((tag.attributes.ref ?? "").replaceAll("$", "")).ref,
    rawXml: xml.slice(tag.start, tag.selfClosing ? tag.end : findElementEnd(xml, tag))
  };
}

function createAutoFilterXml(xml: string, autoFilter: WorksheetAutoFilter): string {
  if (autoFilter.rawXml !== undefined) {
    return normalizeAutoFilterRawXml(autoFilter.rawXml, autoFilter.ref);
  }

  return `<${qualifiedName(inferWorksheetPrefix(xml), "autoFilter")} ref="${escapeXmlAttribute(autoFilter.ref)}"/>`;
}

function normalizeAutoFilterRawXml(rawXml: string, ref: string): string {
  const tag = findFirstStartTag(rawXml, "autoFilter");
  if (tag === undefined || tag.start !== 0) {
    throw new WorksheetError("Auto filter rawXml must start with an autoFilter element");
  }

  const end = tag.selfClosing ? tag.end : findElementEnd(rawXml, tag);
  if (end !== rawXml.length) {
    throw new WorksheetError("Auto filter rawXml must contain exactly one autoFilter element");
  }

  return `${rawXml.slice(0, tag.start)}${upsertRefAttribute(tag.raw, ref)}${rawXml.slice(tag.end)}`;
}

function autoFilterInsertOffset(xml: string): number {
  for (const localName of [
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst"
  ]) {
    const tag = findFirstStartTag(xml, localName);
    if (tag !== undefined) {
      return tag.start;
    }
  }

  const worksheet = findFirstStartTag(xml, "worksheet");
  if (worksheet === undefined) {
    throw new WorksheetError("Worksheet is missing worksheet root");
  }

  return findElementCloseStart(xml, worksheet);
}

function normalizeConditionalFormat(
  conditionalFormat: WorksheetConditionalFormat
): WorksheetConditionalFormat {
  if (conditionalFormat.rules.length === 0) {
    throw new WorksheetError("Conditional format must include at least one rule");
  }

  return {
    sqref: normalizeSqref(conditionalFormat.sqref, "Conditional formatting sqref"),
    ...(conditionalFormat.pivot === undefined ? {} : { pivot: conditionalFormat.pivot }),
    rules: conditionalFormat.rules.map(normalizeConditionalFormatRule)
  };
}

function normalizeConditionalFormatRule(
  rule: WorksheetConditionalFormatRule
): WorksheetConditionalFormatRule {
  return {
    ...(rule.aboveAverage === undefined ? {} : { aboveAverage: rule.aboveAverage }),
    ...(rule.attributes === undefined ? {} : { attributes: { ...rule.attributes } }),
    ...(rule.bottom === undefined ? {} : { bottom: rule.bottom }),
    ...(rule.dxfId === undefined ? {} : { dxfId: rule.dxfId }),
    ...(rule.equalAverage === undefined ? {} : { equalAverage: rule.equalAverage }),
    ...(rule.formulas === undefined ? {} : { formulas: [...rule.formulas] }),
    ...(rule.operator === undefined ? {} : { operator: rule.operator }),
    ...(rule.percent === undefined ? {} : { percent: rule.percent }),
    ...(rule.priority === undefined ? {} : { priority: rule.priority }),
    ...(rule.rank === undefined ? {} : { rank: rule.rank }),
    ...(rule.rawXml === undefined ? {} : { rawXml: rule.rawXml }),
    ...(rule.stdDev === undefined ? {} : { stdDev: rule.stdDev }),
    ...(rule.stopIfTrue === undefined ? {} : { stopIfTrue: rule.stopIfTrue }),
    ...(rule.text === undefined ? {} : { text: rule.text }),
    ...(rule.timePeriod === undefined ? {} : { timePeriod: rule.timePeriod }),
    ...(rule.type === undefined ? {} : { type: rule.type })
  };
}

function conditionalFormatFromTag(
  xml: string,
  tag: ReturnType<typeof findStartTags>[number]
): WorksheetConditionalFormat {
  const rawXml = xml.slice(tag.start, tag.selfClosing ? tag.end : findElementEnd(xml, tag));
  const conditionalFormat: WorksheetConditionalFormat = {
    sqref: normalizeSqref(tag.attributes.sqref ?? "", "Conditional formatting sqref"),
    rules: findStartTags(rawXml, "cfRule").map((rule) => conditionalFormatRuleFromTag(rawXml, rule))
  };

  if (tag.attributes.pivot !== undefined) {
    conditionalFormat.pivot = xmlBoolean(tag.attributes.pivot);
  }

  return conditionalFormat;
}

function conditionalFormatRuleFromTag(
  xml: string,
  tag: ReturnType<typeof findStartTags>[number]
): WorksheetConditionalFormatRule {
  const rawXml = xml.slice(tag.start, tag.selfClosing ? tag.end : findElementEnd(xml, tag));
  const rule: WorksheetConditionalFormatRule = {
    attributes: { ...tag.attributes },
    rawXml
  };

  for (const [attribute, key] of [
    ["dxfId", "dxfId"],
    ["operator", "operator"],
    ["priority", "priority"],
    ["rank", "rank"],
    ["stdDev", "stdDev"],
    ["text", "text"],
    ["timePeriod", "timePeriod"],
    ["type", "type"]
  ] as const) {
    if (tag.attributes[attribute] !== undefined) {
      rule[key] = tag.attributes[attribute];
    }
  }

  for (const [attribute, key] of [
    ["aboveAverage", "aboveAverage"],
    ["bottom", "bottom"],
    ["equalAverage", "equalAverage"],
    ["percent", "percent"],
    ["stopIfTrue", "stopIfTrue"]
  ] as const) {
    if (tag.attributes[attribute] !== undefined) {
      rule[key] = xmlBoolean(tag.attributes[attribute]);
    }
  }

  const formulas = findStartTags(rawXml, "formula").map((formula) =>
    formula.selfClosing
      ? ""
      : decodeXml(rawXml.slice(formula.end, findElementCloseStart(rawXml, formula)))
  );
  if (formulas.length > 0) {
    rule.formulas = formulas;
  }

  return rule;
}

function createConditionalFormatXml(
  xml: string,
  conditionalFormat: WorksheetConditionalFormat
): string {
  const prefix = inferWorksheetPrefix(xml);
  const tagName = qualifiedName(prefix, "conditionalFormatting");
  const attributes = [
    `sqref="${escapeXmlAttribute(conditionalFormat.sqref)}"`,
    conditionalFormat.pivot === undefined
      ? undefined
      : `pivot="${conditionalFormat.pivot ? "1" : "0"}"`
  ]
    .filter((attribute): attribute is string => attribute !== undefined)
    .join(" ");
  const rulesXml = conditionalFormat.rules
    .map((rule) => createConditionalFormatRuleXml(prefix, rule))
    .join("");

  return `<${tagName} ${attributes}>${rulesXml}</${tagName}>`;
}

function createConditionalFormatRuleXml(
  prefix: string | undefined,
  rule: WorksheetConditionalFormatRule
): string {
  if (rule.rawXml !== undefined) {
    validateConditionalFormatRuleXml(rule.rawXml);
    return rule.rawXml;
  }

  const attributes = conditionalFormatRuleAttributes(rule);
  if (attributes.type === undefined || attributes.priority === undefined) {
    throw new WorksheetError("Conditional format rule must include type and priority");
  }

  const tagName = qualifiedName(prefix, "cfRule");
  const attributesXml = conditionalFormatRuleAttributeXml(attributes);
  const formulasXml = (rule.formulas ?? [])
    .map(
      (formula) =>
        `<${qualifiedName(prefix, "formula")}>${escapeXmlText(formula)}</${qualifiedName(prefix, "formula")}>`
    )
    .join("");

  return formulasXml.length === 0
    ? `<${tagName} ${attributesXml}/>`
    : `<${tagName} ${attributesXml}>${formulasXml}</${tagName}>`;
}

function validateConditionalFormatRuleXml(rawXml: string): void {
  const tag = findFirstStartTag(rawXml, "cfRule");
  if (tag === undefined || tag.start !== 0) {
    throw new WorksheetError("Conditional format rawXml must start with a cfRule element");
  }

  const end = tag.selfClosing ? tag.end : findElementEnd(rawXml, tag);
  if (end !== rawXml.length) {
    throw new WorksheetError("Conditional format rawXml must contain exactly one cfRule element");
  }
}

function conditionalFormatRuleAttributes(
  rule: WorksheetConditionalFormatRule
): Record<string, string> {
  const attributes: Record<string, string> = {};
  for (const [name, value] of Object.entries(rule.attributes ?? {})) {
    attributes[name] = conditionalFormatAttributeValue(value);
  }

  for (const [name, value] of [
    ["type", rule.type],
    ["priority", rule.priority],
    ["operator", rule.operator],
    ["dxfId", rule.dxfId],
    ["stopIfTrue", rule.stopIfTrue],
    ["aboveAverage", rule.aboveAverage],
    ["percent", rule.percent],
    ["bottom", rule.bottom],
    ["equalAverage", rule.equalAverage],
    ["rank", rule.rank],
    ["stdDev", rule.stdDev],
    ["text", rule.text],
    ["timePeriod", rule.timePeriod]
  ] as const) {
    if (value !== undefined) {
      attributes[name] = conditionalFormatAttributeValue(value);
    }
  }

  return attributes;
}

function conditionalFormatRuleAttributeXml(attributes: Record<string, string>): string {
  const order = [
    "type",
    "priority",
    "operator",
    "dxfId",
    "stopIfTrue",
    "aboveAverage",
    "percent",
    "bottom",
    "equalAverage",
    "rank",
    "stdDev",
    "text",
    "timePeriod"
  ];
  const emitted = new Set<string>();
  const parts: string[] = [];

  for (const name of order) {
    const value = attributes[name];
    if (value === undefined) {
      continue;
    }

    emitted.add(name);
    parts.push(`${name}="${escapeXmlAttribute(value)}"`);
  }

  for (const name of Object.keys(attributes).sort()) {
    if (emitted.has(name)) {
      continue;
    }

    parts.push(`${name}="${escapeXmlAttribute(attributes[name] ?? "")}"`);
  }

  return parts.join(" ");
}

function conditionalFormatAttributeValue(value: string | number | boolean): string {
  if (typeof value === "boolean") {
    return value ? "1" : "0";
  }

  return String(value);
}

function findMatchingConditionalFormat(
  xml: string,
  sqref: string
): ReturnType<typeof findStartTags>[number] | undefined {
  return findStartTags(xml, "conditionalFormatting").find(
    (tag) => normalizeSqref(tag.attributes.sqref ?? "", "Conditional formatting sqref") === sqref
  );
}

function conditionalFormatInsertOffset(xml: string): number {
  for (const localName of [
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst"
  ]) {
    const tag = findFirstStartTag(xml, localName);
    if (tag !== undefined) {
      return tag.start;
    }
  }

  const worksheet = findFirstStartTag(xml, "worksheet");
  if (worksheet === undefined) {
    throw new WorksheetError("Worksheet is missing worksheet root");
  }

  return findElementCloseStart(xml, worksheet);
}

function xmlBoolean(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

function normalizeMergeRef(ref: string): string {
  return parseCellRange(ref.replaceAll("$", "")).ref;
}

function createMergeCellXml(xml: string, merge: WorksheetMergedCell): string {
  return `<${qualifiedName(inferWorksheetPrefix(xml), "mergeCell")} ref="${escapeXmlAttribute(merge.ref)}"/>`;
}

function mergeContainerInsertOffset(xml: string): number {
  for (const localName of [
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst"
  ]) {
    const tag = findFirstStartTag(xml, localName);
    if (tag !== undefined) {
      return tag.start;
    }
  }

  const worksheet = findFirstStartTag(xml, "worksheet");
  if (worksheet === undefined) {
    throw new WorksheetError("Worksheet is missing worksheet root");
  }

  return findElementCloseStart(xml, worksheet);
}

function cellRangesIntersect(left: CellRange, right: CellRange): boolean {
  return (
    left.start.column <= right.end.column &&
    left.end.column >= right.start.column &&
    left.start.row <= right.end.row &&
    left.end.row >= right.start.row
  );
}

function createHyperlinkXml(xml: string, hyperlink: WorksheetHyperlink): string {
  const tagName = qualifiedName(inferWorksheetPrefix(xml), "hyperlink");
  const attributes = [
    `ref="${escapeXmlAttribute(hyperlink.ref)}"`,
    hyperlink.relationshipId === undefined
      ? undefined
      : `r:id="${escapeXmlAttribute(hyperlink.relationshipId)}"`,
    hyperlink.location === undefined
      ? undefined
      : `location="${escapeXmlAttribute(hyperlink.location)}"`,
    hyperlink.display === undefined
      ? undefined
      : `display="${escapeXmlAttribute(hyperlink.display)}"`,
    hyperlink.tooltip === undefined
      ? undefined
      : `tooltip="${escapeXmlAttribute(hyperlink.tooltip)}"`
  ]
    .filter((attribute): attribute is string => attribute !== undefined)
    .join(" ");

  return `<${tagName} ${attributes}/>`;
}

function findMatchingHyperlink(
  xml: string,
  ref: string
): ReturnType<typeof findStartTags>[number] | undefined {
  return findStartTags(xml, "hyperlink").find(
    (tag) => normalizeHyperlinkRef(tag.attributes.ref ?? "") === ref
  );
}

function hyperlinkContainerInsertOffset(xml: string): number {
  for (const localName of [
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst"
  ]) {
    const tag = findFirstStartTag(xml, localName);
    if (tag !== undefined) {
      return tag.start;
    }
  }

  const worksheet = findFirstStartTag(xml, "worksheet");
  if (worksheet === undefined) {
    throw new WorksheetError("Worksheet is missing worksheet root");
  }

  return findElementCloseStart(xml, worksheet);
}

function drawingInsertOffset(xml: string): number {
  for (const localName of [
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst"
  ]) {
    const tag = findFirstStartTag(xml, localName);
    if (tag !== undefined) {
      return tag.start;
    }
  }

  const worksheet = findFirstStartTag(xml, "worksheet");
  if (worksheet === undefined) {
    throw new WorksheetError("Worksheet is missing worksheet root");
  }

  return findElementCloseStart(xml, worksheet);
}

function ensureWorksheetRelationshipNamespace(xml: string): string {
  const worksheet = findFirstStartTag(xml, "worksheet");
  if (worksheet === undefined) {
    throw new WorksheetError("Worksheet is missing worksheet root");
  }

  if (worksheet.attributes["xmlns:r"] !== undefined) {
    return xml;
  }

  const updated = upsertTagAttribute(
    worksheet.raw,
    "xmlns:r",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  );
  return `${xml.slice(0, worksheet.start)}${updated}${xml.slice(worksheet.end)}`;
}

function qualifiedName(prefix: string | undefined, localName: string): string {
  return prefix === undefined ? localName : `${prefix}:${localName}`;
}

function isFormulaValue(value: CellInput): value is FormulaValue {
  return (
    typeof value === "object" && value !== null && !(value instanceof Date) && "formula" in value
  );
}
