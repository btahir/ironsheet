import { WorksheetError } from "./errors.ts";

export type CellAddress = {
  address: string;
  column: number;
  row: number;
};

export type CellRange = {
  ref: string;
  start: CellAddress;
  end: CellAddress;
};

export function parseCellAddress(address: string): CellAddress {
  const match = /^([A-Za-z]+)([1-9][0-9]*)$/.exec(address);
  if (match === null) {
    throw new WorksheetError(`Invalid cell address ${address}`);
  }

  const [, rawColumnLabel, rawRowLabel] = match;
  if (rawColumnLabel === undefined || rawRowLabel === undefined) {
    throw new WorksheetError(`Invalid cell address ${address}`);
  }

  const columnLabel = rawColumnLabel.toUpperCase();
  const row = Number.parseInt(rawRowLabel, 10);
  let column = 0;

  for (const char of columnLabel) {
    column = column * 26 + (char.charCodeAt(0) - 64);
  }

  return {
    address: `${columnLabel}${row}`,
    column,
    row
  };
}

export function parseCellRange(ref: string): CellRange {
  const parts = ref.split(":");
  if (parts.length > 2) {
    throw new WorksheetError(`Invalid cell range ${ref}`);
  }

  const start = parseCellAddress(parts[0] ?? "");
  const end = parseCellAddress(parts[1] ?? parts[0] ?? "");
  const normalizedStart = {
    address: `${numberToColumnLabel(Math.min(start.column, end.column))}${Math.min(start.row, end.row)}`,
    column: Math.min(start.column, end.column),
    row: Math.min(start.row, end.row)
  };
  const normalizedEnd = {
    address: `${numberToColumnLabel(Math.max(start.column, end.column))}${Math.max(start.row, end.row)}`,
    column: Math.max(start.column, end.column),
    row: Math.max(start.row, end.row)
  };

  return {
    ref:
      normalizedStart.address === normalizedEnd.address
        ? normalizedStart.address
        : `${normalizedStart.address}:${normalizedEnd.address}`,
    start: normalizedStart,
    end: normalizedEnd
  };
}

export function formatCellAddress(column: number, row: number): string {
  return `${numberToColumnLabel(column)}${row}`;
}

export function columnLabelToNumber(label: string): number {
  return parseCellAddress(`${label}1`).column;
}

export function splitCellAddress(address: string): { columnLabel: string; rowLabel: string } {
  const parsed = parseCellAddress(address);
  return {
    columnLabel: numberToColumnLabel(parsed.column),
    rowLabel: String(parsed.row)
  };
}

export function numberToColumnLabel(column: number): string {
  if (!Number.isInteger(column) || column < 1) {
    throw new WorksheetError(`Invalid column number ${column}`);
  }

  let remaining = column;
  let label = "";

  while (remaining > 0) {
    remaining -= 1;
    label = String.fromCharCode(65 + (remaining % 26)) + label;
    remaining = Math.floor(remaining / 26);
  }

  return label;
}

export function compareCellAddresses(a: string, b: string): number {
  const left = parseCellAddress(a);
  const right = parseCellAddress(b);

  if (left.row !== right.row) {
    return left.row - right.row;
  }

  return left.column - right.column;
}
