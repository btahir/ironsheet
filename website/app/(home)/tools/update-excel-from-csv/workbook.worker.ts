import {
  inspectWorkbookArchiveFromBytes,
  openWorkbookFromBytes,
  writeWorkbookToBlobSafely,
} from '@ironsheet/browser';
import {
  diffZipPackages,
  type Workbook,
  type WorkbookFeatureSummary,
  type WorkbookTable,
} from '@ironsheet/core';
import { csvValueToCellInput, parseCsv, type ParsedCsv } from './csv';
import type {
  CsvInspectionResult,
  RefreshResult,
  WorkbookInspectionResult,
  WorkbookTableSummary,
  WorkerRequest,
  WorkerResponse,
} from './worker-protocol';

type WorkerScope = {
  onmessage: ((event: MessageEvent<WorkerRequest>) => void) | null;
  postMessage(message: WorkerResponse): void;
};

const workerScope = globalThis as unknown as WorkerScope;
let workbookBytes: Uint8Array | undefined;
let workbookFileName = '';
let workbookInspection: WorkbookInspectionResult | undefined;
let parsedCsv: ParsedCsv | undefined;

workerScope.onmessage = (event) => {
  void handleRequest(event.data).catch((error: unknown) => {
    workerScope.postMessage({
      type: 'error',
      code: errorCode(error),
      message: error instanceof Error ? error.message : 'The workbook could not be processed.',
    });
  });
};

async function handleRequest(request: WorkerRequest): Promise<void> {
  if (request.type === 'inspect-workbook') {
    postProgress('Checking workbook size and structure');
    workbookBytes = new Uint8Array(request.bytes);
    workbookFileName = request.fileName;
    const archive = inspectWorkbookArchiveFromBytes(workbookBytes);
    if (!archive.accepted) {
      throw new Error(archive.issues[0]?.message ?? 'The workbook exceeds the safe browser limits.');
    }

    postProgress('Inspecting workbook features');
    const workbook = await openWorkbookFromBytes(workbookBytes);
    const [inspection, tables, validation] = await Promise.all([
      workbook.inspect(),
      workbook.tables(),
      workbook.validate(),
    ]);
    const tableSummaries = await Promise.all(
      tables.map((table) => summarizeTable(workbook, table)),
    );
    workbookInspection = {
      archive,
      features: inspection.features,
      fileName: request.fileName,
      tables: tableSummaries,
      validation: validation.summary,
    };
    workerScope.postMessage({ type: 'workbook-inspected', result: workbookInspection });
    return;
  }

  if (request.type === 'parse-csv') {
    postProgress('Reading CSV headers and rows');
    parsedCsv = parseCsv(request.text);
    const result: CsvInspectionResult = {
      delimiter: delimiterName(parsedCsv.delimiter),
      headers: parsedCsv.headers,
      preview: parsedCsv.rows.slice(0, 5),
      rowCount: parsedCsv.rows.length,
      warnings: parsedCsv.warnings,
    };
    workerScope.postMessage({ type: 'csv-parsed', result });
    return;
  }

  if (workbookBytes === undefined || workbookInspection === undefined || parsedCsv === undefined) {
    throw new Error('Choose an Excel workbook and CSV before creating the update.');
  }
  const table = workbookInspection.tables.find((candidate) => candidate.name === request.tableName);
  if (table === undefined) {
    throw new Error(`The selected table ${request.tableName} is no longer available.`);
  }
  if (request.mapping.length !== table.columns.length || request.mapping.some((index) => index < 0)) {
    throw new Error('Match every Excel table column to a CSV column before continuing.');
  }

  postProgress('Opening a fresh copy of the workbook');
  const workbook = await openWorkbookFromBytes(workbookBytes);
  const replacementRows = parsedCsv.rows.map((row) =>
    request.mapping.map((sourceIndex) => csvValueToCellInput(row[sourceIndex] ?? '')),
  );

  postProgress(`Replacing ${replacementRows.length.toLocaleString()} table rows`);
  await workbook.replaceTableRows(request.tableName, replacementRows);
  postProgress('Validating the updated workbook');
  const contentType = workbookFileName.toLocaleLowerCase().endsWith('.xlsm')
    ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const written = await writeWorkbookToBlobSafely(workbook, { type: contentType });
  if (!written.wrote || written.blob === undefined) {
    throw new Error(
      `Ironsheet refused to create an invalid workbook (${written.validation.summary.errors} validation errors).`,
    );
  }

  const outputBytes = new Uint8Array(await written.blob.arrayBuffer());
  const packageDiff = diffZipPackages(workbookBytes, outputBytes);
  const updatedWorkbook = await openWorkbookFromBytes(outputBytes);
  const updatedInspection = await updatedWorkbook.inspect();

  const result: RefreshResult = {
    blob: written.blob,
    diagnostics: workbook.diagnostics().map((diagnostic) => ({
      code: diagnostic.code,
      message: diagnostic.message,
      severity: diagnostic.severity,
    })),
    fileName: updatedFileName(workbookFileName),
    newRowCount: replacementRows.length,
    oldRowCount: table.rowCount,
    packageDiff: {
      changedParts: packageDiff.entries
        .filter((entry) => entry.contentChanged)
        .map((entry) => entry.name),
      contentChanges:
        packageDiff.summary.added + packageDiff.summary.changed + packageDiff.summary.removed,
      repacked: packageDiff.summary.repacked,
      unchanged: packageDiff.summary.unchanged,
    },
    preservedFeatures: preservedFeatures(
      workbookInspection.features,
      updatedInspection.features,
    ),
    tableName: table.displayName,
    validation: written.validation.summary,
  };
  workerScope.postMessage({ type: 'refreshed', result });
}

function preservedFeatures(
  before: WorkbookFeatureSummary,
  after: WorkbookFeatureSummary,
): Array<{ count: number; label: string }> {
  const labels: Array<[keyof WorkbookFeatureSummary, string, string]> = [
    ['charts', 'chart', 'charts'],
    ['media', 'embedded image', 'embedded images'],
    ['formulaCells', 'formula', 'formulas'],
    ['hiddenSheets', 'hidden sheet', 'hidden sheets'],
    ['conditionalFormats', 'conditional format', 'conditional formats'],
    ['dataValidations', 'validation rule', 'validation rules'],
    ['definedNames', 'defined name', 'defined names'],
    ['pivotTables', 'pivot table', 'pivot tables'],
    ['macros', 'macro project', 'macro projects'],
  ];

  return labels.flatMap(([key, singular, plural]) => {
    const count = before[key];
    return count > 0 && after[key] === count
      ? [{ count, label: count === 1 ? singular : plural }]
      : [];
  });
}

async function summarizeTable(workbook: Workbook, table: WorkbookTable): Promise<WorkbookTableSummary> {
  const range = parseRange(table.ref);
  const bodyStart = range.startRow + 1;
  const bodyEnd = range.endRow - table.totalsRowCount;
  const previewEnd = Math.min(bodyEnd, bodyStart + 4);
  const preview =
    previewEnd < bodyStart
      ? []
      : (
          await workbook.readRange(
            sheetNameForPart(workbook, table.worksheetPartName),
            `${columnLabel(range.startColumn)}${bodyStart}:${columnLabel(range.endColumn)}${previewEnd}`,
          )
        ).cells.map((row) => row.map((cell) => displayCellValue(cell?.value ?? null)));

  return {
    columns: table.columns.map((column, index) => column.name ?? `Column ${index + 1}`),
    displayName: table.displayName,
    name: table.name,
    preview,
    ref: table.ref,
    rowCount: Math.max(0, bodyEnd - bodyStart + 1),
    sheetName: sheetNameForPart(workbook, table.worksheetPartName),
  };
}

function sheetNameForPart(workbook: Workbook, partName: string): string {
  return workbook.sheets().find((sheet) => sheet.partName === partName)?.name ?? partName;
}

function parseRange(ref: string): {
  endColumn: number;
  endRow: number;
  startColumn: number;
  startRow: number;
} {
  const [start = '', end = start] = ref.split(':');
  const parsedStart = parseAddress(start);
  const parsedEnd = parseAddress(end);
  return {
    endColumn: parsedEnd.column,
    endRow: parsedEnd.row,
    startColumn: parsedStart.column,
    startRow: parsedStart.row,
  };
}

function parseAddress(address: string): { column: number; row: number } {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(address);
  if (match === null) {
    throw new Error(`Unsupported table range: ${address}`);
  }
  let column = 0;
  for (const char of (match[1] ?? '').toLocaleUpperCase()) {
    column = column * 26 + char.charCodeAt(0) - 64;
  }
  return { column, row: Number.parseInt(match[2] ?? '0', 10) };
}

function columnLabel(column: number): string {
  let current = column;
  let label = '';
  while (current > 0) {
    current -= 1;
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26);
  }
  return label;
}

function displayCellValue(value: string | number | boolean | Date | null): string {
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  return value === null ? '' : String(value);
}

function delimiterName(delimiter: string): string {
  if (delimiter === '\t') return 'Tab';
  if (delimiter === ';') return 'Semicolon';
  if (delimiter === '|') return 'Pipe';
  return 'Comma';
}

function updatedFileName(fileName: string): string {
  const match = /^(.*?)(\.(?:xlsx|xlsm))$/i.exec(fileName);
  return match === null ? `${fileName}-updated.xlsx` : `${match[1]}-updated${match[2]}`;
}

function errorCode(error: unknown): string {
  if (typeof error === 'object' && error !== null && 'code' in error) {
    return String(error.code);
  }
  const message = error instanceof Error ? error.message : '';
  if (/password|encrypted/i.test(message)) return 'WORKBOOK_ENCRYPTED';
  if (/compression|ZIP|central directory/i.test(message)) return 'WORKBOOK_ARCHIVE';
  if (/occupied worksheet row/i.test(message)) return 'TABLE_EXPANSION_BLOCKED';
  return 'PROCESSING_FAILED';
}

function postProgress(label: string): void {
  workerScope.postMessage({ type: 'progress', label });
}
