import { parseCellAddress, parseCellRange, type CellAddress, type CellRange } from "./address.ts";

export const excelMaxColumn = 16_384;
export const excelMaxRow = 1_048_576;

export type FormulaSheetReference = {
  sheetName: string;
};

export type FormulaCellReference = {
  kind: "cell";
  ref: string;
  start: CellAddress;
  sheetName?: string;
};

export type FormulaRangeReference = {
  kind: "range";
  ref: string;
  start: CellAddress;
  end: CellAddress;
  range: CellRange;
  sheetName?: string;
};

export type FormulaReference = FormulaCellReference | FormulaRangeReference;

export type FormulaStructuredReference = {
  tableName: string;
  raw: string;
};

export function parseFormulaSheetReferences(formula: string): FormulaSheetReference[] {
  const references = new Set<string>();
  const scrubbed = stripDoubleQuotedStrings(formula);
  const pattern = /(?:^|[,( +\-*/^&=<>])((?:'(?:(?:'')|[^'])+'|[A-Za-z_][A-Za-z0-9_ .]*))!/g;

  for (const match of scrubbed.matchAll(pattern)) {
    const rawName = match[1];
    if (rawName === undefined || rawName.includes("[")) {
      continue;
    }

    references.add(unquoteSheetName(rawName));
  }

  return [...references].map((sheetName) => ({ sheetName }));
}

export function parseFormulaStructuredReferences(formula: string): FormulaStructuredReference[] {
  const references = new Map<string, FormulaStructuredReference>();
  const scrubbed = stripDoubleQuotedStrings(formula);
  let offset = 0;

  while (offset < scrubbed.length) {
    const bracket = scrubbed.indexOf("[", offset);
    if (bracket === -1) {
      break;
    }

    const tableStart = findTableNameStart(scrubbed, bracket);
    if (tableStart === undefined) {
      offset = bracket + 1;
      continue;
    }

    if (!isReferenceBoundaryBefore(scrubbed, tableStart)) {
      offset = bracket + 1;
      continue;
    }

    const tableName = scrubbed.slice(tableStart, bracket);
    const end = findStructuredReferenceEnd(scrubbed, bracket);
    if (end === undefined) {
      offset = bracket + 1;
      continue;
    }

    const raw = formula.slice(tableStart, end);
    if (!references.has(raw)) {
      references.set(raw, { tableName, raw });
    }

    offset = end;
  }

  return [...references.values()];
}

export function parseFormulaReferences(formula: string): FormulaReference[] {
  const references = new Map<string, FormulaReference>();
  const scrubbed = stripDoubleQuotedStrings(formula);
  const pattern =
    /(?:(?<sheet>'(?:(?:'')|[^'])+'|[A-Za-z_][A-Za-z0-9_ .]*)!)?(?<first>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})(?::(?<second>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}))?/g;

  for (const match of scrubbed.matchAll(pattern)) {
    const raw = match[0];
    const startIndex = match.index ?? 0;
    const endIndex = startIndex + raw.length;
    if (
      !isReferenceBoundaryBefore(scrubbed, startIndex) ||
      !isReferenceBoundaryAfter(scrubbed, endIndex)
    ) {
      continue;
    }

    const rawSheetToken = match.groups?.sheet;
    if (rawSheetToken !== undefined && startIndex > 0 && scrubbed[startIndex - 1] === "]") {
      continue;
    }

    const firstAddress = match.groups?.first;
    if (firstAddress === undefined) {
      continue;
    }

    const secondAddress = match.groups?.second;
    if (rawSheetToken === undefined && secondAddress === undefined && scrubbed[endIndex] === "(") {
      continue;
    }

    const sheetName = rawSheetToken === undefined ? undefined : unquoteSheetName(rawSheetToken);
    const ref = normalizedReference(firstAddress, secondAddress);
    const key = `${sheetName ?? ""}:${ref}`;
    if (references.has(key)) {
      continue;
    }

    if (secondAddress === undefined) {
      const start = parseCellAddress(stripAbsoluteMarkers(firstAddress));
      references.set(key, {
        kind: "cell",
        ref,
        start,
        ...(sheetName === undefined ? {} : { sheetName })
      });
      continue;
    }

    const range = parseCellRange(ref);
    references.set(key, {
      kind: "range",
      ref: range.ref,
      start: range.start,
      end: range.end,
      range,
      ...(sheetName === undefined ? {} : { sheetName })
    });
  }

  return [...references.values()];
}

export function formulaReferenceWithinExcelBounds(reference: FormulaReference): boolean {
  return (
    reference.start.column <= excelMaxColumn &&
    reference.start.row <= excelMaxRow &&
    (reference.kind === "cell" ||
      (reference.end.column <= excelMaxColumn && reference.end.row <= excelMaxRow))
  );
}

function findTableNameStart(source: string, bracket: number): number | undefined {
  let offset = bracket - 1;
  while (offset >= 0 && /[A-Za-z0-9_.]/.test(source[offset] ?? "")) {
    offset -= 1;
  }

  const start = offset + 1;
  if (start === bracket || !/^[A-Za-z_]/.test(source[start] ?? "")) {
    return undefined;
  }

  return start;
}

function findStructuredReferenceEnd(source: string, bracket: number): number | undefined {
  let depth = 0;
  for (let offset = bracket; offset < source.length; offset += 1) {
    const char = source[offset];
    if (char === "[") {
      depth += 1;
    } else if (char === "]") {
      depth -= 1;
      if (depth === 0) {
        return offset + 1;
      }
    }
  }

  return undefined;
}

function stripDoubleQuotedStrings(formula: string): string {
  let result = "";
  let offset = 0;

  while (offset < formula.length) {
    const char = formula[offset];
    if (char !== '"') {
      result += char;
      offset += 1;
      continue;
    }

    result += " ";
    offset += 1;
    while (offset < formula.length) {
      if (formula[offset] === '"') {
        if (formula[offset + 1] === '"') {
          result += "  ";
          offset += 2;
          continue;
        }

        result += " ";
        offset += 1;
        break;
      }

      result += " ";
      offset += 1;
    }
  }

  return result;
}

function normalizedReference(firstAddress: string, secondAddress: string | undefined): string {
  const first = stripAbsoluteMarkers(firstAddress);
  if (secondAddress === undefined) {
    return parseCellAddress(first).address;
  }

  return parseCellRange(`${first}:${stripAbsoluteMarkers(secondAddress)}`).ref;
}

function stripAbsoluteMarkers(address: string): string {
  return address.replaceAll("$", "");
}

function isReferenceBoundaryBefore(source: string, index: number): boolean {
  if (index === 0) {
    return true;
  }

  return !/[A-Za-z0-9_.$\]]/.test(source[index - 1] ?? "");
}

function isReferenceBoundaryAfter(source: string, index: number): boolean {
  if (index >= source.length) {
    return true;
  }

  return !/[A-Za-z0-9_.$]/.test(source[index] ?? "");
}

function unquoteSheetName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replaceAll("''", "'");
  }

  return trimmed;
}
