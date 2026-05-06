import { WorkbookError } from "./errors.ts";
import {
  escapeXmlAttribute,
  findElementCloseStart,
  findFirstStartTag,
  findStartTags,
  type XmlTag
} from "./xml.ts";

export const drawingRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

export const imageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

export const drawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";

export type WorkbookImageExtension =
  | "png"
  | "jpg"
  | "jpeg"
  | "gif"
  | "bmp"
  | "tif"
  | "tiff"
  | "webp";

export type WorkbookImageAnchorMarker = {
  column: number;
  row: number;
  columnOffset?: number;
  rowOffset?: number;
};

export type WorkbookImageExtent = {
  cx: number;
  cy: number;
};

export type WorkbookImageAnchor =
  | {
      kind: "oneCell";
      from: WorkbookImageAnchorMarker;
      ext: WorkbookImageExtent;
    }
  | {
      kind: "twoCell";
      from: WorkbookImageAnchorMarker;
      to: WorkbookImageAnchorMarker;
      editAs?: "twoCell" | "oneCell" | "absolute";
    };

export type WorkbookInsertImageOptions = {
  name?: string;
  description?: string;
  extension?: WorkbookImageExtension;
  anchor?: WorkbookImageAnchor;
};

export type WorkbookImage = {
  sheetName: string;
  sheetPartName: string;
  drawingPartName: string;
  drawingRelationshipId: string;
  imageRelationshipId: string;
  target: string;
  imagePartName?: string;
  targetMode?: string;
};

const defaultImageExtent: WorkbookImageExtent = { cx: 914400, cy: 914400 };
const spreadsheetDrawingNamespace =
  "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const drawingMainNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
const officeDocumentRelationshipsNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export function normalizeImageExtension(extension: string): WorkbookImageExtension {
  const normalized = extension.replace(/^\./, "").toLowerCase();
  if (isWorkbookImageExtension(normalized)) {
    return normalized;
  }

  throw new WorkbookError(`Unsupported image extension ${extension}`);
}

export function imageContentTypeForExtension(extension: WorkbookImageExtension): string {
  if (extension === "jpg" || extension === "jpeg") {
    return "image/jpeg";
  }

  if (extension === "tif" || extension === "tiff") {
    return "image/tiff";
  }

  return `image/${extension}`;
}

export function imageExtensionForBytes(data: Uint8Array): WorkbookImageExtension | undefined {
  for (const extension of ["png", "jpg", "gif", "bmp", "tif", "webp"] as const) {
    const signature = imageSignatureForExtension(extension);
    if (signature?.matches(data) === true) {
      return extension;
    }
  }

  return undefined;
}

export function assertImageBytesMatchExtension(
  extension: WorkbookImageExtension,
  data: Uint8Array
): void {
  const expected = imageSignatureForExtension(extension);
  if (expected === undefined || expected.matches(data)) {
    return;
  }

  throw new WorkbookError(`Image extension ${extension} expects ${expected.label} bytes`);
}

export function assertImageBytesMatchPartName(partName: string, data: Uint8Array): void {
  const extension = partName.slice(partName.lastIndexOf(".") + 1).toLowerCase();
  if (!isWorkbookImageExtension(extension)) {
    return;
  }

  const expected = imageSignatureForExtension(extension);
  if (expected === undefined || expected.matches(data)) {
    return;
  }

  throw new WorkbookError(`Image part ${partName} expects ${expected.label} bytes`);
}

export function createDrawingXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="${spreadsheetDrawingNamespace}" xmlns:a="${drawingMainNamespace}" xmlns:r="${officeDocumentRelationshipsNamespace}">
</xdr:wsDr>`;
}

export function appendDrawingAnchorXml(drawingXml: string, anchorXml: string): string {
  const drawing = findFirstStartTag(drawingXml, "wsDr");
  if (drawing === undefined) {
    throw new WorkbookError("Drawing part is missing wsDr root");
  }

  const namespaced = ensureDrawingNamespaces(drawingXml, drawing);
  const nextDrawing = findFirstStartTag(namespaced, "wsDr");
  if (nextDrawing === undefined) {
    throw new WorkbookError("Drawing part is missing wsDr root");
  }

  if (nextDrawing.selfClosing) {
    const opening = nextDrawing.raw.replace(/\/>$/, ">");
    return `${namespaced.slice(0, nextDrawing.start)}${opening}
  ${anchorXml}
</${nextDrawing.name}>${namespaced.slice(nextDrawing.end)}`;
  }

  const insertOffset = findElementCloseStart(namespaced, nextDrawing);
  return `${namespaced.slice(0, insertOffset)}  ${anchorXml}
${namespaced.slice(insertOffset)}`;
}

export function createPictureAnchorXml(
  relationshipId: string,
  pictureId: number,
  options: WorkbookInsertImageOptions
): string {
  const anchor = options.anchor ?? {
    kind: "oneCell" as const,
    from: { column: 0, row: 0 },
    ext: defaultImageExtent
  };
  const name = options.name ?? `Picture ${pictureId}`;
  const description = options.description ?? name;
  const pictureXml = createPictureXml(relationshipId, pictureId, name, description, anchor);

  if (anchor.kind === "oneCell") {
    return `<xdr:oneCellAnchor>
    ${markerXml("from", anchor.from)}
    <xdr:ext cx="${anchor.ext.cx}" cy="${anchor.ext.cy}"/>
    ${pictureXml}
    <xdr:clientData/>
  </xdr:oneCellAnchor>`;
  }

  const editAs =
    anchor.editAs === undefined ? "" : ` editAs="${escapeXmlAttribute(anchor.editAs)}"`;
  return `<xdr:twoCellAnchor${editAs}>
    ${markerXml("from", anchor.from)}
    ${markerXml("to", anchor.to)}
    ${pictureXml}
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;
}

export function nextDrawingPictureId(drawingXml: string): number {
  const ids = findStartTags(drawingXml, "cNvPr")
    .map((tag) => Number.parseInt(tag.attributes.id ?? "0", 10))
    .filter((id) => Number.isSafeInteger(id) && id > 0);
  return Math.max(0, ...ids) + 1;
}

function createPictureXml(
  relationshipId: string,
  pictureId: number,
  name: string,
  description: string,
  anchor: WorkbookImageAnchor
): string {
  const ext = anchor.kind === "oneCell" ? anchor.ext : defaultImageExtent;
  return `<xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="${pictureId}" name="${escapeXmlAttribute(name)}" descr="${escapeXmlAttribute(description)}"/>
        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="${escapeXmlAttribute(relationshipId)}"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${ext.cx}" cy="${ext.cy}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>`;
}

function markerXml(name: "from" | "to", marker: WorkbookImageAnchorMarker): string {
  return `<xdr:${name}><xdr:col>${marker.column}</xdr:col><xdr:colOff>${marker.columnOffset ?? 0}</xdr:colOff><xdr:row>${marker.row}</xdr:row><xdr:rowOff>${marker.rowOffset ?? 0}</xdr:rowOff></xdr:${name}>`;
}

function ensureDrawingNamespaces(xml: string, drawing: XmlTag): string {
  const attributes = [
    ["xmlns:xdr", spreadsheetDrawingNamespace],
    ["xmlns:a", drawingMainNamespace],
    ["xmlns:r", officeDocumentRelationshipsNamespace]
  ] as const;
  let raw = drawing.raw;

  for (const [name, value] of attributes) {
    if (drawing.attributes[name] !== undefined) {
      continue;
    }

    raw = upsertTagAttribute(raw, name, value);
  }

  return `${xml.slice(0, drawing.start)}${raw}${xml.slice(drawing.end)}`;
}

function upsertTagAttribute(rawTag: string, name: string, value: string): string {
  const escaped = escapeXmlAttribute(value);
  const attribute = `${name}="${escaped}"`;
  const existing = new RegExp(`\\s${escapeRegExp(name)}="[^"]*"`);
  if (existing.test(rawTag)) {
    return rawTag.replace(existing, ` ${attribute}`);
  }

  return rawTag.replace(/\/?>$/, (ending) => ` ${attribute}${ending}`);
}

function imageSignatureForExtension(
  extension: WorkbookImageExtension
): { label: string; matches: (data: Uint8Array) => boolean } | undefined {
  if (extension === "png") {
    return {
      label: "PNG",
      matches: (data) => startsWithBytes(data, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    };
  }

  if (extension === "jpg" || extension === "jpeg") {
    return {
      label: "JPEG",
      matches: (data) => data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff
    };
  }

  if (extension === "gif") {
    return {
      label: "GIF",
      matches: (data) => startsWithAscii(data, "GIF87a") || startsWithAscii(data, "GIF89a")
    };
  }

  if (extension === "bmp") {
    return {
      label: "BMP",
      matches: (data) => startsWithAscii(data, "BM")
    };
  }

  if (extension === "tif" || extension === "tiff") {
    return {
      label: "TIFF",
      matches: (data) =>
        startsWithBytes(data, [0x49, 0x49, 0x2a, 0x00]) ||
        startsWithBytes(data, [0x4d, 0x4d, 0x00, 0x2a])
    };
  }

  if (extension === "webp") {
    return {
      label: "WEBP",
      matches: (data) => startsWithAscii(data, "RIFF") && startsWithAscii(data.subarray(8), "WEBP")
    };
  }

  return undefined;
}

function isWorkbookImageExtension(value: string): value is WorkbookImageExtension {
  return (
    value === "png" ||
    value === "jpg" ||
    value === "jpeg" ||
    value === "gif" ||
    value === "bmp" ||
    value === "tif" ||
    value === "tiff" ||
    value === "webp"
  );
}

function startsWithAscii(data: Uint8Array, value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    if (data[index] !== value.charCodeAt(index)) {
      return false;
    }
  }

  return true;
}

function startsWithBytes(data: Uint8Array, bytes: number[]): boolean {
  if (data.byteLength < bytes.length) {
    return false;
  }

  return bytes.every((byte, index) => data[index] === byte);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
