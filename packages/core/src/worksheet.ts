import { compareCellAddresses, parseCellAddress } from "./address.ts";
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
