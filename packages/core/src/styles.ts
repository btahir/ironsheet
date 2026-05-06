import { findElementCloseStart, findFirstStartTag, findStartTags } from "./xml.ts";

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
    }
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

function countChildren(xml: string, containerName: string, childName: string): number {
  const container = findFirstStartTag(xml, containerName);
  if (container === undefined) {
    return 0;
  }

  return findStartTags(xml.slice(container.end, findElementCloseStart(xml, container)), childName)
    .length;
}
