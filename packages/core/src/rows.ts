import { formatCellAddress, parseCellAddress, parseCellRange } from "./address.ts";
import { WorksheetError } from "./errors.ts";
import { excelMaxRow, type FormulaRowEdit } from "./formula.ts";
import { findElementEnd, findStartTags, type XmlTag } from "./xml.ts";

export type WorksheetRowEditResult = {
  removedFormulaCells: number;
  removedHyperlinkRelationshipIds: string[];
  xml: string;
};

export function insertWorksheetRows(
  xml: string,
  beforeRow: number,
  count: number
): WorksheetRowEditResult {
  assertRowEditArguments(beforeRow, count);
  if (beforeRow + count - 1 > excelMaxRow) {
    throw new WorksheetError(
      `Inserting ${count} row(s) before row ${beforeRow} exceeds the Excel row limit`
    );
  }

  const edit: FormulaRowEdit = { count, mode: "insert", sheetName: "", startRow: beforeRow };
  return applyWorksheetRowEdit(xml, edit);
}

export function deleteWorksheetRows(
  xml: string,
  startRow: number,
  count: number
): WorksheetRowEditResult {
  assertRowEditArguments(startRow, count);
  if (startRow + count - 1 > excelMaxRow) {
    throw new WorksheetError(
      `Deleting ${count} row(s) from row ${startRow} exceeds the Excel row limit`
    );
  }

  const edit: FormulaRowEdit = { count, mode: "delete", sheetName: "", startRow };
  return applyWorksheetRowEdit(xml, edit);
}

export function mapRowThroughEdit(
  row: number,
  edit: FormulaRowEdit,
  position: "cell" | "rangeStart" | "rangeEnd"
): number | undefined {
  if (edit.mode === "insert") {
    return row >= edit.startRow ? row + edit.count : row;
  }

  const deletedEnd = edit.startRow + edit.count - 1;
  if (row < edit.startRow) {
    return row;
  }

  if (row > deletedEnd) {
    return row - edit.count;
  }

  if (position === "cell") {
    return undefined;
  }

  return position === "rangeStart" ? edit.startRow : edit.startRow - 1;
}

export function mapRangeRefThroughEdit(ref: string, edit: FormulaRowEdit): string | undefined {
  const range = parseCellRange(ref);
  const start = mapRowThroughEdit(range.start.row, edit, "rangeStart");
  const end = mapRowThroughEdit(range.end.row, edit, "rangeEnd");
  if (start === undefined || end === undefined || end < start) {
    return undefined;
  }

  if (start === range.start.row && end === range.end.row) {
    return ref;
  }

  const first = formatCellAddress(range.start.column, start);
  if (!ref.includes(":")) {
    return first;
  }

  return `${first}:${formatCellAddress(range.end.column, end)}`;
}

function applyWorksheetRowEdit(xml: string, edit: FormulaRowEdit): WorksheetRowEditResult {
  const rowResult = transformSheetRows(xml, edit);
  let nextXml = rowResult.xml;
  nextXml = transformMergedCells(nextXml, edit);
  const hyperlinkResult = transformHyperlinks(nextXml, edit);
  nextXml = hyperlinkResult.xml;
  nextXml = transformSqrefElements(nextXml, edit, "dataValidation", "sqref");
  nextXml = transformSqrefElements(nextXml, edit, "conditionalFormatting", "sqref");
  nextXml = removeEmptyContainer(nextXml, "dataValidations", "dataValidation");
  nextXml = transformSingleRefElement(nextXml, edit, "autoFilter");
  nextXml = transformDimension(nextXml, edit);

  return {
    removedFormulaCells: rowResult.removedFormulaCells,
    removedHyperlinkRelationshipIds: hyperlinkResult.removedRelationshipIds,
    xml: nextXml
  };
}

function transformSheetRows(
  xml: string,
  edit: FormulaRowEdit
): { removedFormulaCells: number; xml: string } {
  const sheetData = findStartTags(xml, "sheetData").at(0);
  if (sheetData === undefined || sheetData.selfClosing) {
    return { removedFormulaCells: 0, xml };
  }

  const replacements: TextReplacement[] = [];
  let removedFormulaCells = 0;

  for (const tag of findStartTags(xml, "row")) {
    const rowNumber = Number.parseInt(tag.attributes.r ?? "", 10);
    if (!Number.isInteger(rowNumber) || rowNumber < 1) {
      continue;
    }

    const elementEnd = findElementEnd(xml, tag);
    const mapped = mapRowThroughEdit(rowNumber, edit, "cell");
    if (mapped === undefined) {
      const removed = xml.slice(tag.start, elementEnd);
      for (const formula of findStartTags(removed, "f")) {
        const ref = formula.attributes.ref;
        if (
          ref !== undefined &&
          (formula.attributes.si !== undefined || formula.attributes.t === "array") &&
          mapRangeRefThroughEdit(ref, edit) !== undefined
        ) {
          throw new WorksheetError(
            `Deleting row ${rowNumber} would orphan the shared or array formula anchored at ${ref}; clear or rewrite the formula range first`
          );
        }
      }

      removedFormulaCells += findStartTags(removed, "f").length;
      replacements.push({ start: tag.start, end: elementEnd, text: "" });
      continue;
    }

    if (mapped === rowNumber) {
      continue;
    }

    const slice = xml.slice(tag.start, elementEnd);
    replacements.push({
      start: tag.start,
      end: elementEnd,
      text: rewriteRowSlice(slice, tag, mapped, edit)
    });
  }

  return { removedFormulaCells, xml: applyTextReplacements(xml, replacements) };
}

function rewriteRowSlice(
  slice: string,
  rowTag: XmlTag,
  newRow: number,
  edit: FormulaRowEdit
): string {
  const replacements: TextReplacement[] = [
    {
      start: 0,
      end: rowTag.raw.length,
      text: upsertTagAttribute(rowTag.raw, "r", String(newRow))
    }
  ];

  for (const cell of findStartTags(slice, "c")) {
    const address = cell.attributes.r;
    if (address === undefined) {
      continue;
    }

    const parsed = parseCellAddress(address);
    replacements.push({
      start: cell.start,
      end: cell.end,
      text: upsertTagAttribute(cell.raw, "r", formatCellAddress(parsed.column, newRow))
    });
  }

  for (const formula of findStartTags(slice, "f")) {
    const ref = formula.attributes.ref;
    if (ref === undefined) {
      continue;
    }

    const mappedRef = mapRangeRefThroughEdit(ref, edit);
    if (mappedRef === undefined || mappedRef === ref) {
      continue;
    }

    replacements.push({
      start: formula.start,
      end: formula.end,
      text: upsertTagAttribute(formula.raw, "ref", mappedRef)
    });
  }

  return applyTextReplacements(slice, replacements);
}

function transformMergedCells(xml: string, edit: FormulaRowEdit): string {
  const container = findStartTags(xml, "mergeCells").at(0);
  if (container === undefined) {
    return xml;
  }

  const replacements: TextReplacement[] = [];
  let remaining = 0;

  for (const tag of findStartTags(xml, "mergeCell")) {
    const ref = tag.attributes.ref;
    if (ref === undefined) {
      continue;
    }

    const elementEnd = findElementEnd(xml, tag);
    const mapped = mapRangeRefThroughEdit(ref, edit);
    if (mapped === undefined || !mapped.includes(":")) {
      replacements.push({ start: tag.start, end: elementEnd, text: "" });
      continue;
    }

    remaining += 1;
    if (mapped !== ref) {
      replacements.push({
        start: tag.start,
        end: tag.end,
        text: upsertTagAttribute(tag.raw, "ref", mapped)
      });
    }
  }

  if (remaining === 0) {
    replacements.push({
      start: container.start,
      end: findElementEnd(xml, container),
      text: ""
    });
    return applyTextReplacements(xml, replacements);
  }

  replacements.push({
    start: container.start,
    end: container.end,
    text: upsertTagAttribute(container.raw, "count", String(remaining))
  });
  return applyTextReplacements(xml, replacements);
}

function transformHyperlinks(
  xml: string,
  edit: FormulaRowEdit
): { removedRelationshipIds: string[]; xml: string } {
  const container = findStartTags(xml, "hyperlinks").at(0);
  if (container === undefined) {
    return { removedRelationshipIds: [], xml };
  }

  const replacements: TextReplacement[] = [];
  const removedRelationshipIds: string[] = [];
  let remaining = 0;

  for (const tag of findStartTags(xml, "hyperlink")) {
    const ref = tag.attributes.ref;
    if (ref === undefined) {
      continue;
    }

    const elementEnd = findElementEnd(xml, tag);
    const mapped = mapRangeRefThroughEdit(ref, edit);
    if (mapped === undefined) {
      const relationshipId = tag.attributes["r:id"];
      if (relationshipId !== undefined) {
        removedRelationshipIds.push(relationshipId);
      }

      replacements.push({ start: tag.start, end: elementEnd, text: "" });
      continue;
    }

    remaining += 1;
    if (mapped !== ref) {
      replacements.push({
        start: tag.start,
        end: tag.end,
        text: upsertTagAttribute(tag.raw, "ref", mapped)
      });
    }
  }

  if (remaining === 0) {
    replacements.push({
      start: container.start,
      end: findElementEnd(xml, container),
      text: ""
    });
  }

  return { removedRelationshipIds, xml: applyTextReplacements(xml, replacements) };
}

function transformSqrefElements(
  xml: string,
  edit: FormulaRowEdit,
  elementName: string,
  attributeName: string
): string {
  const replacements: TextReplacement[] = [];

  for (const tag of findStartTags(xml, elementName)) {
    const sqref = tag.attributes[attributeName];
    if (sqref === undefined) {
      continue;
    }

    const mappedRefs = sqref
      .split(/\s+/)
      .filter((ref) => ref.length > 0)
      .map((ref) => mapRangeRefThroughEdit(ref, edit))
      .filter((ref): ref is string => ref !== undefined);

    if (mappedRefs.length === 0) {
      replacements.push({ start: tag.start, end: findElementEnd(xml, tag), text: "" });
      continue;
    }

    const mapped = mappedRefs.join(" ");
    if (mapped !== sqref) {
      replacements.push({
        start: tag.start,
        end: tag.end,
        text: upsertTagAttribute(tag.raw, attributeName, mapped)
      });
    }
  }

  return applyTextReplacements(xml, replacements);
}

function transformSingleRefElement(xml: string, edit: FormulaRowEdit, elementName: string): string {
  const replacements: TextReplacement[] = [];

  for (const tag of findStartTags(xml, elementName)) {
    const ref = tag.attributes.ref;
    if (ref === undefined) {
      continue;
    }

    const mapped = mapRangeRefThroughEdit(ref, edit);
    if (mapped === undefined) {
      replacements.push({ start: tag.start, end: findElementEnd(xml, tag), text: "" });
      continue;
    }

    if (mapped !== ref) {
      replacements.push({
        start: tag.start,
        end: tag.end,
        text: upsertTagAttribute(tag.raw, "ref", mapped)
      });
    }
  }

  return applyTextReplacements(xml, replacements);
}

function transformDimension(xml: string, edit: FormulaRowEdit): string {
  const tag = findStartTags(xml, "dimension").at(0);
  if (tag === undefined) {
    return xml;
  }

  const ref = tag.attributes.ref;
  if (ref === undefined) {
    return xml;
  }

  const mapped = mapRangeRefThroughEdit(ref, edit) ?? "A1";
  if (mapped === ref) {
    return xml;
  }

  return applyTextReplacements(xml, [
    { start: tag.start, end: tag.end, text: upsertTagAttribute(tag.raw, "ref", mapped) }
  ]);
}

function removeEmptyContainer(xml: string, containerName: string, childName: string): string {
  const container = findStartTags(xml, containerName).at(0);
  if (container === undefined || container.selfClosing) {
    return xml;
  }

  const elementEnd = findElementEnd(xml, container);
  const body = xml.slice(container.end, elementEnd);
  const childCount = findStartTags(body, childName).length;
  if (childCount > 0) {
    const updated = upsertTagAttribute(container.raw, "count", String(childCount));
    if (updated === container.raw) {
      return xml;
    }

    return applyTextReplacements(xml, [
      { start: container.start, end: container.end, text: updated }
    ]);
  }

  return applyTextReplacements(xml, [{ start: container.start, end: elementEnd, text: "" }]);
}

function assertRowEditArguments(startRow: number, count: number): void {
  if (!Number.isInteger(startRow) || startRow < 1 || startRow > excelMaxRow) {
    throw new WorksheetError(`Invalid row ${startRow}; expected 1..${excelMaxRow}`);
  }

  if (!Number.isInteger(count) || count < 1) {
    throw new WorksheetError(`Invalid row count ${count}; expected a positive integer`);
  }
}

type TextReplacement = {
  start: number;
  end: number;
  text: string;
};

function applyTextReplacements(source: string, replacements: TextReplacement[]): string {
  if (replacements.length === 0) {
    return source;
  }

  let result = "";
  let offset = 0;

  for (const replacement of [...replacements].sort((left, right) => left.start - right.start)) {
    if (replacement.start < offset) {
      continue;
    }

    result += source.slice(offset, replacement.start);
    result += replacement.text;
    offset = replacement.end;
  }

  return result + source.slice(offset);
}

function upsertTagAttribute(rawTag: string, name: string, value: string): string {
  const pattern = new RegExp(`(\\s${name}=)(["']).*?\\2`);
  if (pattern.test(rawTag)) {
    return rawTag.replace(pattern, `$1"${value}"`);
  }

  const closing = rawTag.endsWith("/>") ? "/>" : ">";
  return `${rawTag.slice(0, -closing.length)} ${name}="${value}"${closing}`;
}
