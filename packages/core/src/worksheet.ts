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

function qualifiedName(prefix: string | undefined, localName: string): string {
  return prefix === undefined ? localName : `${prefix}:${localName}`;
}

function isFormulaValue(value: CellInput): value is FormulaValue {
  return (
    typeof value === "object" && value !== null && !(value instanceof Date) && "formula" in value
  );
}
