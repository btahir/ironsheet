export const drawingRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

export const imageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

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
