import {
  compareCellAddresses,
  formatCellAddress,
  parseCellAddress,
  parseCellRange
} from "./address.ts";
import { WorksheetError } from "./errors.ts";
import {
  decodeXml,
  escapeXmlAttribute,
  escapeXmlText,
  findFirstStartTag,
  findStartTags
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
  const cellXml = createCellXml(
    parsedAddress.address,
    value,
    existing?.styleId === undefined ? {} : { styleId: existing.styleId }
  );
  const formulaChanged = isFormulaValue(value);
  const withCell =
    existing === undefined
      ? insertCell(xml, parsedAddress.address, parsedAddress.row, cellXml)
      : `${xml.slice(0, existing.start)}${cellXml}${xml.slice(existing.end)}`;

  return {
    xml: updateDimension(withCell, parsedAddress.address),
    formulaChanged
  };
}

export function patchCells(xml: string, patches: CellPatch[]): PatchCellResult {
  let nextXml = xml;
  let formulaChanged = false;

  for (const patch of patches) {
    const result = patchCell(nextXml, patch.address, patch.value);
    nextXml = result.xml;
    formulaChanged = formulaChanged || result.formulaChanged;
  }

  return {
    xml: nextXml,
    formulaChanged
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

export function appendRows(
  xml: string,
  rows: CellInput[][],
  options: { startColumn?: number } = {}
): PatchCellResult {
  const startColumn = options.startColumn ?? 1;
  const startRow = findMaxUsedRow(xml) + 1;
  return patchRange(xml, formatCellAddress(startColumn, startRow), rows);
}

export function replaceRowsInRange(
  xml: string,
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  rows: CellInput[][],
  options: { preserveStyles?: boolean } = {}
): string {
  const sheetData = findFirstStartTag(xml, "sheetData");
  if (sheetData === undefined) {
    throw new WorksheetError("Worksheet is missing sheetData");
  }

  const sheetDataClose = xml.indexOf("</sheetData>", sheetData.end);
  if (sheetDataClose === -1) {
    throw new WorksheetError("Worksheet sheetData is missing a closing tag");
  }

  const preserveStyles = options.preserveStyles ?? true;
  const template = preserveStyles ? collectRowTemplate(xml, range) : undefined;
  const rowElements = findRowElements(xml)
    .filter((row) => row.rowNumber >= range.startRow && row.rowNumber <= range.endRow)
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
            styleId === undefined ? {} : { styleId }
          );
        })
        .join("");

      return `<row r="${rowNumber}"${rowAttributesXml(template?.attributes)}>${cells}</row>`;
    })
    .join("");

  const updated = `${nextXml.slice(0, insertionPoint)}${rowXml}${nextXml.slice(insertionPoint)}`;
  return recalculateDimension(updated);
}

export function createCellXml(
  address: string,
  value: CellInput,
  options: { styleId?: string } = {}
): string {
  const attributes = createCellAttributes(address, value, options);

  if (isFormulaValue(value)) {
    const result = value.result === undefined ? "" : createFormulaResultXml(value.result);
    return `<c ${attributes}><f>${escapeXmlText(value.formula.replace(/^=/, ""))}</f>${result}</c>`;
  }

  if (value === null) {
    return `<c ${attributes}/>`;
  }

  if (typeof value === "number") {
    return `<c ${attributes}><v>${String(value)}</v></c>`;
  }

  if (typeof value === "boolean") {
    return `<c ${attributes}><v>${value ? "1" : "0"}</v></c>`;
  }

  if (value instanceof Date) {
    return `<c ${attributes}><v>${dateToExcelSerial(value)}</v></c>`;
  }

  return `<c ${attributes}><is><t>${escapeXmlText(value)}</t></is></c>`;
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

    const close = xml.indexOf("</c>", tag.end);
    if (close === -1) {
      throw new WorksheetError(`Cell ${address} is missing a closing </c> tag`);
    }

    return existingCell(xml, tag.start, close + "</c>".length, tag.attributes.s);
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

  const close = xml.lastIndexOf("</row>", row.end);
  if (close === -1 || close < row.start) {
    throw new WorksheetError(`Row ${rowNumber} is missing a closing </row> tag`);
  }

  return `${xml.slice(0, close)}${cellXml}${xml.slice(close)}`;
}

function insertRow(xml: string, rowNumber: number, cellXml: string): string {
  const sheetData = findFirstStartTag(xml, "sheetData");
  if (sheetData === undefined) {
    throw new WorksheetError("Worksheet is missing sheetData");
  }

  const close = xml.indexOf("</sheetData>", sheetData.end);
  if (close === -1) {
    throw new WorksheetError("Worksheet sheetData is missing a closing tag");
  }

  const rowXml = `<row r="${rowNumber}">${cellXml}</row>`;
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

    const close = xml.indexOf("</row>", row.end);
    if (close === -1) {
      throw new WorksheetError(`Row ${rowNumber} is missing a closing </row> tag`);
    }

    return { rowNumber, start: row.start, end: close + "</row>".length, tag: row };
  });
}

function findRowInsertionPoint(xml: string, rowNumber: number): number {
  const nextRow = findRowElements(xml).find((row) => row.rowNumber > rowNumber);
  if (nextRow !== undefined) {
    return nextRow.start;
  }

  const sheetDataClose = xml.indexOf("</sheetData>");
  if (sheetDataClose === -1) {
    throw new WorksheetError("Worksheet sheetData is missing a closing tag");
  }

  return sheetDataClose;
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

function createFormulaResultXml(value: CellPrimitive): string {
  if (value === null) {
    return "";
  }

  if (typeof value === "boolean") {
    return `<v>${value ? "1" : "0"}</v>`;
  }

  if (typeof value === "number") {
    return `<v>${String(value)}</v>`;
  }

  if (value instanceof Date) {
    return `<v>${dateToExcelSerial(value)}</v>`;
  }

  return `<v>${escapeXmlText(value)}</v>`;
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

  if (typeof value === "boolean") {
    attributes.push('t="b"');
  } else if (typeof value === "string") {
    attributes.push('t="inlineStr"');
  }

  return attributes.join(" ");
}

function readCellValue(cellXml: string, sharedStrings: string[]): CellPrimitive {
  if (/t=(["'])inlineStr\1/.test(cellXml)) {
    return readTagText(cellXml, "t") ?? "";
  }

  if (/t=(["'])s\1/.test(cellXml)) {
    const index = Number.parseInt(readTagText(cellXml, "v") ?? "", 10);
    return Number.isInteger(index) ? (sharedStrings[index] ?? null) : null;
  }

  if (/t=(["'])b\1/.test(cellXml)) {
    return readTagText(cellXml, "v") === "1";
  }

  const value = readTagText(cellXml, "v");
  if (value === undefined) {
    return null;
  }

  const numericValue = Number(value);
  return Number.isFinite(numericValue) ? numericValue : value;
}

function readTagText(xml: string, localName: string): string | undefined {
  const open = new RegExp(`<(?:[A-Za-z0-9_]+:)?${localName}(?:\\s[^>]*)?>`).exec(xml);
  if (open === null || open.index === undefined) {
    return undefined;
  }

  const start = open.index + open[0].length;
  const close = xml.indexOf(`</${localName}>`, start);
  const prefixedClose = close === -1 ? findPrefixedClose(xml, localName, start) : close;

  if (prefixedClose === -1) {
    return undefined;
  }

  return decodeXml(xml.slice(start, prefixedClose));
}

function findPrefixedClose(xml: string, localName: string, start: number): number {
  const close = new RegExp(`</[A-Za-z0-9_]+:${localName}>`).exec(xml.slice(start));
  return close?.index === undefined ? -1 : start + close.index;
}

function dateToExcelSerial(date: Date): number {
  const epoch = Date.UTC(1899, 11, 30);
  return (date.getTime() - epoch) / 86_400_000;
}

function isFormulaValue(value: CellInput): value is FormulaValue {
  return (
    typeof value === "object" && value !== null && !(value instanceof Date) && "formula" in value
  );
}
