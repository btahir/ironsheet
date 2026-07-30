import type { WorkbookArchiveInspection } from '@ironsheet/browser';
import type { WorkbookFeatureSummary } from '@ironsheet/core';

export type WorkbookTableSummary = {
  columns: string[];
  displayName: string;
  name: string;
  preview: string[][];
  ref: string;
  rowCount: number;
  sheetName: string;
};

export type WorkbookInspectionResult = {
  archive: WorkbookArchiveInspection;
  features: WorkbookFeatureSummary;
  fileName: string;
  tables: WorkbookTableSummary[];
  validation: { errors: number; infos: number; warnings: number };
};

export type CsvInspectionResult = {
  delimiter: string;
  headers: string[];
  preview: string[][];
  rowCount: number;
  warnings: string[];
};

export type RefreshResult = {
  blob: Blob;
  diagnostics: Array<{ code: string; message: string; severity: string }>;
  fileName: string;
  newRowCount: number;
  oldRowCount: number;
  tableName: string;
  validation: { errors: number; infos: number; warnings: number };
};

export type WorkerRequest =
  | { type: 'inspect-workbook'; bytes: ArrayBuffer; fileName: string }
  | { type: 'parse-csv'; text: string }
  | { type: 'refresh'; mapping: number[]; tableName: string };

export type WorkerResponse =
  | { type: 'progress'; label: string }
  | { type: 'workbook-inspected'; result: WorkbookInspectionResult }
  | { type: 'csv-parsed'; result: CsvInspectionResult }
  | { type: 'refreshed'; result: RefreshResult }
  | { type: 'error'; code: string; message: string };
