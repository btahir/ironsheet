import { WorkbookError } from "./errors.ts";
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
  alignment?: WorkbookCellAlignment;
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

export type WorkbookCellAlignment = {
  horizontal?: string;
  indent?: string;
  shrinkToFit?: string;
  textRotation?: string;
  vertical?: string;
  wrapText?: string;
};

export type WorkbookFontInput = {
  bold?: boolean;
  color?: string;
  italic?: boolean;
  name?: string;
  size?: number;
  strike?: boolean;
  underline?: boolean | "double";
};

export type WorkbookFillInput =
  | string
  | {
      backgroundColor?: string;
      color?: string;
      pattern?: string;
    };

export type WorkbookBorderStyleInput =
  | "dashDot"
  | "dashDotDot"
  | "dashed"
  | "dotted"
  | "double"
  | "hair"
  | "medium"
  | "mediumDashDot"
  | "mediumDashDotDot"
  | "mediumDashed"
  | "slantDashDot"
  | "thick"
  | "thin";

export type WorkbookBorderEdgeInput = {
  color?: string;
  style: WorkbookBorderStyleInput;
};

export type WorkbookBorderInput = {
  all?: WorkbookBorderEdgeInput;
  bottom?: WorkbookBorderEdgeInput;
  diagonal?: WorkbookBorderEdgeInput & { down?: boolean; up?: boolean };
  left?: WorkbookBorderEdgeInput;
  right?: WorkbookBorderEdgeInput;
  top?: WorkbookBorderEdgeInput;
};

export type WorkbookAlignmentInput = {
  horizontal?:
    | "center"
    | "centerContinuous"
    | "distributed"
    | "fill"
    | "general"
    | "justify"
    | "left"
    | "right";
  indent?: number | string;
  shrinkToFit?: boolean | string;
  textRotation?: number | string;
  vertical?: "bottom" | "center" | "distributed" | "justify" | "top";
  wrapText?: boolean | string;
};

export type WorkbookCellStyleInput = Partial<Omit<WorkbookCellFormat, "alignment">> & {
  alignment?: WorkbookAlignmentInput | WorkbookCellAlignment;
  border?: WorkbookBorderInput;
  fill?: WorkbookFillInput;
  font?: WorkbookFontInput;
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

export type EnsureWorkbookStyleComponentResult = {
  created: boolean;
  id: string;
  xml: string;
};

export type EnsureWorkbookStyleInputResult = {
  format: WorkbookCellFormat;
  xml: string;
};

export function ensureWorkbookFont(
  xml: string,
  font: WorkbookFontInput
): EnsureWorkbookStyleComponentResult {
  return ensureStyleComponent(xml, "fonts", "font", fontXml(font));
}

export function ensureWorkbookFill(
  xml: string,
  fill: WorkbookFillInput
): EnsureWorkbookStyleComponentResult {
  return ensureStyleComponent(xml, "fills", "fill", fillXml(fill));
}

export function ensureWorkbookBorder(
  xml: string,
  border: WorkbookBorderInput
): EnsureWorkbookStyleComponentResult {
  return ensureStyleComponent(xml, "borders", "border", borderXml(border));
}

export function ensureWorkbookStyleComponents(
  xml: string,
  input: WorkbookCellStyleInput
): EnsureWorkbookStyleInputResult {
  const { alignment, border, fill, font, numberFormat, ...attributes } = input;
  let nextXml = xml;
  const format: WorkbookCellFormat = { ...attributes };

  if (numberFormat !== undefined) {
    const result = ensureWorkbookNumberFormat(nextXml, numberFormat);
    nextXml = result.xml;
    format.numFmtId = result.numFmtId;
    format.applyNumberFormat = "1";
  }

  if (font !== undefined) {
    const result = ensureWorkbookFont(nextXml, font);
    nextXml = result.xml;
    format.fontId = result.id;
    format.applyFont = "1";
  }

  if (fill !== undefined) {
    const result = ensureWorkbookFill(nextXml, fill);
    nextXml = result.xml;
    format.fillId = result.id;
    format.applyFill = "1";
  }

  if (border !== undefined) {
    const result = ensureWorkbookBorder(nextXml, border);
    nextXml = result.xml;
    format.borderId = result.id;
    format.applyBorder = "1";
  }

  if (alignment !== undefined) {
    format.alignment = normalizeAlignmentInput(alignment);
    format.applyAlignment = "1";
  }

  return { format, xml: nextXml };
}

export function normalizeStyleColor(color: string): string {
  const trimmed = (color.startsWith("#") ? color.slice(1) : color).toUpperCase();
  if (!/^[0-9A-F]{6}([0-9A-F]{2})?$/.test(trimmed)) {
    throw new WorkbookError(
      `Invalid style color ${color}; expected RGB or ARGB hex like 1F4E79 or FF1F4E79`
    );
  }

  return trimmed.length === 6 ? `FF${trimmed}` : trimmed;
}

function fontXml(font: WorkbookFontInput): string {
  const parts: string[] = [];
  if (font.bold === true) {
    parts.push("<b/>");
  }
  if (font.italic === true) {
    parts.push("<i/>");
  }
  if (font.strike === true) {
    parts.push("<strike/>");
  }
  if (font.underline !== undefined && font.underline !== false) {
    parts.push(font.underline === "double" ? '<u val="double"/>' : "<u/>");
  }
  if (font.size !== undefined) {
    if (!Number.isFinite(font.size) || font.size <= 0) {
      throw new WorkbookError(`Invalid font size ${font.size}`);
    }
    parts.push(`<sz val="${escapeXmlAttribute(String(font.size))}"/>`);
  }
  if (font.color !== undefined) {
    parts.push(`<color rgb="${normalizeStyleColor(font.color)}"/>`);
  }
  if (font.name !== undefined) {
    parts.push(`<name val="${escapeXmlAttribute(font.name)}"/>`);
  }

  return parts.length === 0 ? "<font/>" : `<font>${parts.join("")}</font>`;
}

function fillXml(fill: WorkbookFillInput): string {
  const input = typeof fill === "string" ? { color: fill } : fill;
  const pattern = input.pattern ?? (input.color === undefined ? "none" : "solid");
  const foreground =
    input.color === undefined ? "" : `<fgColor rgb="${normalizeStyleColor(input.color)}"/>`;
  const background =
    input.backgroundColor === undefined
      ? input.color === undefined
        ? ""
        : '<bgColor indexed="64"/>'
      : `<bgColor rgb="${normalizeStyleColor(input.backgroundColor)}"/>`;
  const body = `${foreground}${background}`;
  const patternFill =
    body.length === 0
      ? `<patternFill patternType="${escapeXmlAttribute(pattern)}"/>`
      : `<patternFill patternType="${escapeXmlAttribute(pattern)}">${body}</patternFill>`;
  return `<fill>${patternFill}</fill>`;
}

function borderXml(border: WorkbookBorderInput): string {
  const edges = {
    bottom: border.bottom ?? border.all,
    diagonal: border.diagonal,
    left: border.left ?? border.all,
    right: border.right ?? border.all,
    top: border.top ?? border.all
  };

  const attributes = [
    border.diagonal?.up === true ? ' diagonalUp="1"' : "",
    border.diagonal?.down === true ? ' diagonalDown="1"' : ""
  ].join("");

  const edgeXml = (name: string, edge: WorkbookBorderEdgeInput | undefined): string => {
    if (edge === undefined) {
      return `<${name}/>`;
    }

    const color = `<color rgb="${normalizeStyleColor(edge.color ?? "FF000000")}"/>`;
    return `<${name} style="${escapeXmlAttribute(edge.style)}">${color}</${name}>`;
  };

  return `<border${attributes}>${edgeXml("left", edges.left)}${edgeXml("right", edges.right)}${edgeXml("top", edges.top)}${edgeXml("bottom", edges.bottom)}${edgeXml("diagonal", edges.diagonal)}</border>`;
}

function normalizeAlignmentInput(
  alignment: WorkbookAlignmentInput | WorkbookCellAlignment
): WorkbookCellAlignment {
  const flag = (value: boolean | string | undefined): string | undefined => {
    if (value === undefined) {
      return undefined;
    }
    if (typeof value === "string") {
      return value;
    }
    return value ? "1" : undefined;
  };

  const numeric = (value: number | string | undefined): string | undefined =>
    value === undefined ? undefined : String(value);

  const indent = numeric(alignment.indent);
  const shrinkToFit = flag(alignment.shrinkToFit);
  const textRotation = numeric(alignment.textRotation);
  const wrapText = flag(alignment.wrapText);

  return {
    ...(alignment.horizontal === undefined ? {} : { horizontal: alignment.horizontal }),
    ...(indent === undefined ? {} : { indent }),
    ...(shrinkToFit === undefined ? {} : { shrinkToFit }),
    ...(textRotation === undefined ? {} : { textRotation }),
    ...(alignment.vertical === undefined ? {} : { vertical: alignment.vertical }),
    ...(wrapText === undefined ? {} : { wrapText })
  };
}

const styleSheetChildOrder = [
  "numFmts",
  "fonts",
  "fills",
  "borders",
  "cellStyleXfs",
  "cellXfs",
  "cellStyles",
  "dxfs",
  "tableStyles",
  "colors",
  "extLst"
];

function insertStyleSheetContainer(
  xml: string,
  containerName: string,
  containerXml: string
): string {
  const order = styleSheetChildOrder.indexOf(containerName);
  for (let index = order - 1; index >= 0; index -= 1) {
    const name = styleSheetChildOrder[index];
    if (name === undefined) {
      continue;
    }

    const sibling = findFirstStartTag(xml, name);
    if (sibling !== undefined) {
      const insertionPoint = findElementEnd(xml, sibling);
      return `${xml.slice(0, insertionPoint)}${containerXml}${xml.slice(insertionPoint)}`;
    }
  }

  return insertAfterStyleSheetOpen(xml, containerXml);
}

function ensureStyleComponent(
  xml: string,
  containerName: string,
  childName: string,
  childXml: string
): EnsureWorkbookStyleComponentResult {
  const container = findFirstStartTag(xml, containerName);
  if (container === undefined) {
    return {
      created: true,
      id: "0",
      xml: insertStyleSheetContainer(
        xml,
        containerName,
        `<${containerName} count="1">${childXml}</${containerName}>`
      )
    };
  }

  if (container.selfClosing) {
    const openTag = upsertCount(`${container.raw.slice(0, -2)}>`, 1);
    return {
      created: true,
      id: "0",
      xml: replaceStartTag(xml, container, `${openTag}${childXml}</${container.name}>`)
    };
  }

  const close = findElementCloseStart(xml, container);
  const body = xml.slice(container.end, close);
  const children = findStartTags(body, childName).map((tag) =>
    body.slice(tag.start, findElementEnd(body, tag))
  );
  const targetKey = normalizeComponentXml(childXml);
  const existingIndex = children.findIndex((child) => normalizeComponentXml(child) === targetKey);
  if (existingIndex !== -1) {
    return {
      created: false,
      id: String(existingIndex),
      xml
    };
  }

  const withChild = `${xml.slice(0, close)}${childXml}${xml.slice(close)}`;
  return {
    created: true,
    id: String(children.length),
    xml: replaceStartTag(withChild, container, upsertCount(container.raw, children.length + 1))
  };
}

function normalizeComponentXml(xml: string): string {
  return xml
    .replace(/>\s+</g, "><")
    .replace(/\s+/g, " ")
    .replace(/\s*\/>/g, "/>")
    .trim();
}

function parseCellFormats(xml: string, containerName: string): WorkbookCellFormat[] {
  const container = findFirstStartTag(xml, containerName);
  if (container === undefined) {
    return [];
  }

  const body = xml.slice(container.end, findElementCloseStart(xml, container));
  return findStartTags(body, "xf").map((tag) => ({
    ...parseCellFormatAlignment(body, tag),
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

function parseCellFormatAlignment(
  body: string,
  tag: XmlTag
): { alignment: WorkbookCellAlignment } | Record<string, never> {
  if (tag.selfClosing) {
    return {};
  }

  const inner = body.slice(tag.end, findElementCloseStart(body, tag));
  const alignmentTag = findFirstStartTag(inner, "alignment");
  if (alignmentTag === undefined) {
    return {};
  }

  const alignment: WorkbookCellAlignment = {
    ...(alignmentTag.attributes.horizontal === undefined
      ? {}
      : { horizontal: alignmentTag.attributes.horizontal }),
    ...(alignmentTag.attributes.indent === undefined
      ? {}
      : { indent: alignmentTag.attributes.indent }),
    ...(alignmentTag.attributes.shrinkToFit === undefined
      ? {}
      : { shrinkToFit: alignmentTag.attributes.shrinkToFit }),
    ...(alignmentTag.attributes.textRotation === undefined
      ? {}
      : { textRotation: alignmentTag.attributes.textRotation }),
    ...(alignmentTag.attributes.vertical === undefined
      ? {}
      : { vertical: alignmentTag.attributes.vertical }),
    ...(alignmentTag.attributes.wrapText === undefined
      ? {}
      : { wrapText: alignmentTag.attributes.wrapText })
  };

  return Object.keys(alignment).length === 0 ? {} : { alignment };
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
    ...(format.alignment === undefined || Object.keys(format.alignment).length === 0
      ? {}
      : { alignment: format.alignment }),
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
  const { alignment, ...attributes } = normalizeCellFormat(format);
  const attributeKey = Object.entries(attributes)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([name, value]) => `${name}=${value}`)
    .join(";");
  if (alignment === undefined) {
    return attributeKey;
  }

  const alignmentKey = Object.entries(alignment)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([name, value]) => `${name}=${value}`)
    .join(",");
  return `${attributeKey};alignment(${alignmentKey})`;
}

function cellFormatXml(format: WorkbookCellFormat): string {
  const { alignment, ...normalized } = normalizeCellFormat(format);
  const attributes = Object.entries(normalized)
    .map(([name, value]) => `${name}="${escapeXmlAttribute(value)}"`)
    .join(" ");
  const openTag = attributes.length === 0 ? "<xf" : `<xf ${attributes}`;
  if (alignment === undefined) {
    return `${openTag}/>`;
  }

  const alignmentAttributes = Object.entries(alignment)
    .map(([name, value]) => `${name}="${escapeXmlAttribute(value)}"`)
    .join(" ");
  return `${openTag}><alignment ${alignmentAttributes}/></xf>`;
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
