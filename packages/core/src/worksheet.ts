import { compareCellAddresses, parseCellAddress } from "./address.ts";
import { WorksheetError } from "./errors.ts";
import { escapeXmlAttribute, escapeXmlText, findFirstStartTag, findStartTags } from "./xml.ts";

export type CellPrimitive = string | number | boolean | Date | null;

export type FormulaValue = {
  formula: string;
  result?: CellPrimitive;
};

export type CellInput = CellPrimitive | FormulaValue;

export type PatchCellResult = {
  xml: string;
  formulaChanged: boolean;
};

export function patchCell(xml: string, address: string, value: CellInput): PatchCellResult {
  const parsedAddress = parseCellAddress(address);
  const cellXml = createCellXml(parsedAddress.address, value);
  const formulaChanged = isFormulaValue(value);
  const existing = findCellElement(xml, parsedAddress.address);
  const withCell =
    existing === undefined
      ? insertCell(xml, parsedAddress.address, parsedAddress.row, cellXml)
      : `${xml.slice(0, existing.start)}${cellXml}${xml.slice(existing.end)}`;

  return {
    xml: updateDimension(withCell, parsedAddress.address),
    formulaChanged
  };
}

export function createCellXml(address: string, value: CellInput): string {
  if (isFormulaValue(value)) {
    const result = value.result === undefined ? "" : createFormulaResultXml(value.result);
    return `<c r="${escapeXmlAttribute(address)}"><f>${escapeXmlText(value.formula.replace(/^=/, ""))}</f>${result}</c>`;
  }

  if (value === null) {
    return `<c r="${escapeXmlAttribute(address)}"/>`;
  }

  if (typeof value === "number") {
    return `<c r="${escapeXmlAttribute(address)}"><v>${String(value)}</v></c>`;
  }

  if (typeof value === "boolean") {
    return `<c r="${escapeXmlAttribute(address)}" t="b"><v>${value ? "1" : "0"}</v></c>`;
  }

  if (value instanceof Date) {
    return `<c r="${escapeXmlAttribute(address)}"><v>${dateToExcelSerial(value)}</v></c>`;
  }

  return `<c r="${escapeXmlAttribute(address)}" t="inlineStr"><is><t>${escapeXmlText(value)}</t></is></c>`;
}

function findCellElement(xml: string, address: string): { start: number; end: number } | undefined {
  const tags = findStartTags(xml, "c");
  const target = address.toUpperCase();

  for (const tag of tags) {
    if ((tag.attributes.r ?? "").toUpperCase() !== target) {
      continue;
    }

    if (tag.selfClosing) {
      return { start: tag.start, end: tag.end };
    }

    const close = xml.indexOf("</c>", tag.end);
    if (close === -1) {
      throw new WorksheetError(`Cell ${address} is missing a closing </c> tag`);
    }

    return { start: tag.start, end: close + "</c>".length };
  }

  return undefined;
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

function findRowElement(
  xml: string,
  rowNumber: number
): { start: number; end: number } | undefined {
  const rows = findStartTags(xml, "row");
  const target = String(rowNumber);

  for (const row of rows) {
    if (row.attributes.r !== target) {
      continue;
    }

    if (row.selfClosing) {
      return { start: row.start, end: row.end };
    }

    const close = xml.indexOf("</row>", row.end);
    if (close === -1) {
      throw new WorksheetError(`Row ${rowNumber} is missing a closing </row> tag`);
    }

    return { start: row.start, end: close + "</row>".length };
  }

  return undefined;
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
  const replacement = dimension.raw.replace(
    /\sref=(["'])(.*?)\1/,
    ` ref="${escapeXmlAttribute(expanded)}"`
  );

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

  return `${columnLabel(minColumn)}${minRow}:${columnLabel(maxColumn)}${maxRow}`;
}

function columnLabel(column: number): string {
  let remaining = column;
  let label = "";

  while (remaining > 0) {
    remaining -= 1;
    label = String.fromCharCode(65 + (remaining % 26)) + label;
    remaining = Math.floor(remaining / 26);
  }

  return label;
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

function dateToExcelSerial(date: Date): number {
  const epoch = Date.UTC(1899, 11, 30);
  return (date.getTime() - epoch) / 86_400_000;
}

function isFormulaValue(value: CellInput): value is FormulaValue {
  return (
    typeof value === "object" && value !== null && !(value instanceof Date) && "formula" in value
  );
}
