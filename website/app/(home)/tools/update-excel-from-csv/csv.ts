export const csvLimits = {
  maxCells: 250_000,
  maxRows: 50_000
} as const;

export type CsvDelimiter = ',' | ';' | '\t' | '|';

export type ParsedCsv = {
  delimiter: CsvDelimiter;
  headers: string[];
  rows: string[][];
  warnings: string[];
};

export class CsvParseError extends Error {
  constructor(
    readonly code: 'CSV_EMPTY' | 'CSV_HEADER_EMPTY' | 'CSV_LIMIT' | 'CSV_UNCLOSED_QUOTE',
    message: string,
  ) {
    super(message);
    this.name = 'CsvParseError';
  }
}

export function parseCsv(text: string): ParsedCsv {
  const normalizedText = text.replace(/^\uFEFF/, '');
  if (normalizedText.trim().length === 0) {
    throw new CsvParseError('CSV_EMPTY', 'The CSV file is empty.');
  }

  const delimiter = detectDelimiter(normalizedText);
  const records = parseRecords(normalizedText, delimiter).filter((record) =>
    record.some((value) => value.length > 0),
  );
  const rawHeaders = records.shift();
  if (rawHeaders === undefined || rawHeaders.length === 0) {
    throw new CsvParseError('CSV_EMPTY', 'The CSV file does not contain a header row.');
  }

  const headers = rawHeaders.map((header) => header.trim());
  const emptyHeader = headers.findIndex((header) => header.length === 0);
  if (emptyHeader !== -1) {
    throw new CsvParseError(
      'CSV_HEADER_EMPTY',
      `Column ${emptyHeader + 1} has an empty header. Add a name and try again.`,
    );
  }

  if (records.length > csvLimits.maxRows || records.length * headers.length > csvLimits.maxCells) {
    throw new CsvParseError(
      'CSV_LIMIT',
      `This first release supports up to ${csvLimits.maxRows.toLocaleString()} rows and ${csvLimits.maxCells.toLocaleString()} cells per refresh.`,
    );
  }

  const rows = records.map((record) =>
    Array.from({ length: headers.length }, (_, index) => record[index] ?? ''),
  );
  const normalizedHeaders = headers.map(normalizeHeader);
  const duplicateHeaders = headers.filter(
    (_, index) => normalizedHeaders.indexOf(normalizedHeaders[index] ?? '') !== index,
  );
  const warnings = duplicateHeaders.length
    ? [`Duplicate CSV headers found: ${[...new Set(duplicateHeaders)].join(', ')}`]
    : [];

  return { delimiter, headers, rows, warnings };
}

export function normalizeHeader(header: string): string {
  return header.trim().toLocaleLowerCase().replace(/[\s_-]+/g, '');
}

export function csvValueToCellInput(value: string): string | number | boolean | null {
  if (value.length === 0) {
    return null;
  }
  if (/^(?:true|false)$/i.test(value)) {
    return value.toLocaleLowerCase() === 'true';
  }
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(value)) {
    const parsed = Number(value);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }
  return value;
}

function detectDelimiter(text: string): CsvDelimiter {
  const candidates: CsvDelimiter[] = [',', '\t', ';', '|'];
  const scores = new Map(candidates.map((candidate) => [candidate, 0]));
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === '"') {
      if (quoted && text[index + 1] === '"') {
        index += 1;
      } else {
        quoted = !quoted;
      }
      continue;
    }
    if (!quoted && (char === '\n' || char === '\r')) {
      break;
    }
    if (!quoted && candidates.includes(char as CsvDelimiter)) {
      const delimiter = char as CsvDelimiter;
      scores.set(delimiter, (scores.get(delimiter) ?? 0) + 1);
    }
  }

  return candidates.reduce((best, candidate) =>
    (scores.get(candidate) ?? 0) > (scores.get(best) ?? 0) ? candidate : best,
  );
}

function parseRecords(text: string, delimiter: CsvDelimiter): string[][] {
  const records: string[][] = [];
  let record: string[] = [];
  let field = '';
  let quoted = false;

  const finishField = () => {
    record.push(field);
    field = '';
  };
  const finishRecord = () => {
    finishField();
    records.push(record);
    record = [];
  };

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index] ?? '';
    if (char === '"') {
      if (quoted && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
      continue;
    }
    if (!quoted && char === delimiter) {
      finishField();
      continue;
    }
    if (!quoted && (char === '\n' || char === '\r')) {
      if (char === '\r' && text[index + 1] === '\n') {
        index += 1;
      }
      finishRecord();
      continue;
    }
    field += char;
  }

  if (quoted) {
    throw new CsvParseError('CSV_UNCLOSED_QUOTE', 'The CSV contains an unclosed quoted field.');
  }
  if (field.length > 0 || record.length > 0) {
    finishRecord();
  }
  return records;
}
