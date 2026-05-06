import {
  escapeXmlAttribute,
  findElementCloseStart,
  findElementEnd,
  findFirstStartTag,
  findStartTags,
  type XmlTag
} from "./xml.ts";

export const excelCellFormatLimit = 65_490;
export const excelCellFormatWarningThreshold = 60_000;

export type WorkbookStyles = {
  cellStyleXfs: WorkbookCellFormat[];
  cellXfs: WorkbookCellFormat[];
  counts: {
    borders: number;
    cellStyleXfs: number;
    cellXfs: number;
    fills: number;
    fonts: number;
    numFmts: number;
  };
  numberFormats: WorkbookNumberFormat[];
};

export type WorkbookCellFormat = {
  applyAlignment?: string;
  applyBorder?: string;
  applyFill?: string;
  applyFont?: string;
  applyNumberFormat?: string;
  borderId?: string;
  fillId?: string;
  fontId?: string;
  numFmtId?: string;
  xfId?: string;
};

export type WorkbookCellStyleInput = Partial<WorkbookCellFormat> & {
  numberFormat?: string;
};

export type WorkbookNumberFormat = {
  numFmtId: string;
  formatCode: string;
};

export type EnsureWorkbookCellFormatResult = {
  created: boolean;
  styleId: string;
  xml: string;
};

export type EnsureWorkbookNumberFormatResult = {
  created: boolean;
  numFmtId: string;
  xml: string;
};

export function parseWorkbookStyles(xml: string): WorkbookStyles {
  return {
    cellStyleXfs: parseCellFormats(xml, "cellStyleXfs"),
    cellXfs: parseCellFormats(xml, "cellXfs"),
    counts: {
      borders: countChildren(xml, "borders", "border"),
      cellStyleXfs: countChildren(xml, "cellStyleXfs", "xf"),
      cellXfs: countChildren(xml, "cellXfs", "xf"),
      fills: countChildren(xml, "fills", "fill"),
      fonts: countChildren(xml, "fonts", "font"),
      numFmts: countChildren(xml, "numFmts", "numFmt")
    },
    numberFormats: parseNumberFormats(xml)
  };
}

export function ensureWorkbookCellFormat(
  xml: string,
  format: WorkbookCellFormat
): EnsureWorkbookCellFormatResult {
  const normalizedFormat = normalizeCellFormat(format);
  const existing = parseCellFormats(xml, "cellXfs");
  const existingIndex = existing.findIndex(
    (candidate) => cellFormatKey(candidate) === cellFormatKey(normalizedFormat)
  );

  if (existingIndex !== -1) {
    return {
      created: false,
      styleId: String(existingIndex),
      xml
    };
  }

  const cellXfs = findFirstStartTag(xml, "cellXfs");
  if (cellXfs === undefined) {
    return insertNewCellXfs(xml, normalizedFormat);
  }

  const nextStyleId = existing.length;
  const xfXml = cellFormatXml(normalizedFormat);
  const close = findElementCloseStart(xml, cellXfs);
  const withFormat = `${xml.slice(0, close)}${xfXml}${xml.slice(close)}`;
  return {
    created: true,
    styleId: String(nextStyleId),
    xml: replaceStartTag(withFormat, cellXfs, upsertCount(cellXfs.raw, nextStyleId + 1))
  };
}

export function ensureWorkbookNumberFormat(
  xml: string,
  formatCode: string
): EnsureWorkbookNumberFormatResult {
  const existing = parseNumberFormats(xml).find((format) => format.formatCode === formatCode);
  if (existing !== undefined) {
    return {
      created: false,
      numFmtId: existing.numFmtId,
      xml
    };
  }

  const nextNumFmtId = String(nextCustomNumberFormatId(xml));
  const numFmtXml = `<numFmt numFmtId="${escapeXmlAttribute(nextNumFmtId)}" formatCode="${escapeXmlAttribute(formatCode)}"/>`;
  const numFmts = findFirstStartTag(xml, "numFmts");
  if (numFmts === undefined) {
    return insertNewNumberFormats(xml, numFmtXml, nextNumFmtId);
  }

  const count = countChildren(xml, "numFmts", "numFmt");
  const close = findElementCloseStart(xml, numFmts);
  const withFormat = `${xml.slice(0, close)}${numFmtXml}${xml.slice(close)}`;
  return {
    created: true,
    numFmtId: nextNumFmtId,
    xml: replaceStartTag(withFormat, numFmts, upsertCount(numFmts.raw, count + 1))
  };
}

function parseCellFormats(xml: string, containerName: string): WorkbookCellFormat[] {
  const container = findFirstStartTag(xml, containerName);
  if (container === undefined) {
    return [];
  }

  const body = xml.slice(container.end, findElementCloseStart(xml, container));
  return findStartTags(body, "xf").map((tag) => ({
    ...(tag.attributes.applyAlignment === undefined
      ? {}
      : { applyAlignment: tag.attributes.applyAlignment }),
    ...(tag.attributes.applyBorder === undefined
      ? {}
      : { applyBorder: tag.attributes.applyBorder }),
    ...(tag.attributes.applyFill === undefined ? {} : { applyFill: tag.attributes.applyFill }),
    ...(tag.attributes.applyFont === undefined ? {} : { applyFont: tag.attributes.applyFont }),
    ...(tag.attributes.applyNumberFormat === undefined
      ? {}
      : { applyNumberFormat: tag.attributes.applyNumberFormat }),
    ...(tag.attributes.borderId === undefined ? {} : { borderId: tag.attributes.borderId }),
    ...(tag.attributes.fillId === undefined ? {} : { fillId: tag.attributes.fillId }),
    ...(tag.attributes.fontId === undefined ? {} : { fontId: tag.attributes.fontId }),
    ...(tag.attributes.numFmtId === undefined ? {} : { numFmtId: tag.attributes.numFmtId }),
    ...(tag.attributes.xfId === undefined ? {} : { xfId: tag.attributes.xfId })
  }));
}

function parseNumberFormats(xml: string): WorkbookNumberFormat[] {
  const container = findFirstStartTag(xml, "numFmts");
  if (container === undefined) {
    return [];
  }

  const body = xml.slice(container.end, findElementCloseStart(xml, container));
  return findStartTags(body, "numFmt")
    .map((tag) => {
      const numFmtId = tag.attributes.numFmtId;
      const formatCode = tag.attributes.formatCode;
      if (numFmtId === undefined || formatCode === undefined) {
        return undefined;
      }

      return { numFmtId, formatCode };
    })
    .filter((format): format is WorkbookNumberFormat => format !== undefined);
}

function countChildren(xml: string, containerName: string, childName: string): number {
  const container = findFirstStartTag(xml, containerName);
  if (container === undefined) {
    return 0;
  }

  return findStartTags(xml.slice(container.end, findElementCloseStart(xml, container)), childName)
    .length;
}

function normalizeCellFormat(format: WorkbookCellFormat): WorkbookCellFormat {
  return {
    ...(format.applyAlignment === undefined ? {} : { applyAlignment: format.applyAlignment }),
    ...(format.applyBorder === undefined ? {} : { applyBorder: format.applyBorder }),
    ...(format.applyFill === undefined ? {} : { applyFill: format.applyFill }),
    ...(format.applyFont === undefined ? {} : { applyFont: format.applyFont }),
    ...(format.applyNumberFormat === undefined
      ? {}
      : { applyNumberFormat: format.applyNumberFormat }),
    ...(format.borderId === undefined ? {} : { borderId: format.borderId }),
    ...(format.fillId === undefined ? {} : { fillId: format.fillId }),
    ...(format.fontId === undefined ? {} : { fontId: format.fontId }),
    ...(format.numFmtId === undefined ? {} : { numFmtId: format.numFmtId }),
    ...(format.xfId === undefined ? {} : { xfId: format.xfId })
  };
}

function cellFormatKey(format: WorkbookCellFormat): string {
  return Object.entries(normalizeCellFormat(format))
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([name, value]) => `${name}=${value}`)
    .join(";");
}

function cellFormatXml(format: WorkbookCellFormat): string {
  const attributes = Object.entries(normalizeCellFormat(format))
    .map(([name, value]) => `${name}="${escapeXmlAttribute(value)}"`)
    .join(" ");

  return attributes.length === 0 ? "<xf/>" : `<xf ${attributes}/>`;
}

function insertNewCellXfs(xml: string, format: WorkbookCellFormat): EnsureWorkbookCellFormatResult {
  const cellXfsXml = `<cellXfs count="1">${cellFormatXml(format)}</cellXfs>`;
  const cellStyleXfs = findFirstStartTag(xml, "cellStyleXfs");
  if (cellStyleXfs !== undefined) {
    const insertionPoint = findElementEnd(xml, cellStyleXfs);
    return {
      created: true,
      styleId: "0",
      xml: `${xml.slice(0, insertionPoint)}${cellXfsXml}${xml.slice(insertionPoint)}`
    };
  }

  return {
    created: true,
    styleId: "0",
    xml: insertAfterStyleSheetOpen(xml, cellXfsXml)
  };
}

function insertNewNumberFormats(
  xml: string,
  numFmtXml: string,
  numFmtId: string
): EnsureWorkbookNumberFormatResult {
  return {
    created: true,
    numFmtId,
    xml: insertAfterStyleSheetOpen(xml, `<numFmts count="1">${numFmtXml}</numFmts>`)
  };
}

function insertAfterStyleSheetOpen(xml: string, insertion: string): string {
  const styleSheet = findFirstStartTag(xml, "styleSheet");
  if (styleSheet === undefined) {
    return `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${insertion}</styleSheet>`;
  }

  return `${xml.slice(0, styleSheet.end)}${insertion}${xml.slice(styleSheet.end)}`;
}

function nextCustomNumberFormatId(xml: string): number {
  return (
    Math.max(
      163,
      ...parseNumberFormats(xml).map((format) => Number.parseInt(format.numFmtId, 10))
    ) + 1
  );
}

function replaceStartTag(xml: string, originalTag: XmlTag, replacement: string): string {
  return `${xml.slice(0, originalTag.start)}${replacement}${xml.slice(originalTag.end)}`;
}

function upsertCount(rawTag: string, count: number): string {
  if (/\scount=(["']).*?\1/.test(rawTag)) {
    return rawTag.replace(/\scount=(["']).*?\1/, ` count="${count}"`);
  }

  const closing = rawTag.endsWith("/>") ? "/>" : ">";
  return `${rawTag.slice(0, -closing.length)} count="${count}"${closing}`;
}
