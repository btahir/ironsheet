import {
  OoxmlPackage,
  Workbook,
  parseZip,
  type CompressionAdapter,
  type ValidationReport
} from "@ironsheet/core";

const defaultArchiveLimits: Required<WorkbookArchiveLimits> = {
  maxCompressedBytes: 25 * 1024 * 1024,
  maxCompressionRatio: 200,
  maxEntries: 10_000,
  maxEntryBytes: 100 * 1024 * 1024,
  maxUncompressedBytes: 250 * 1024 * 1024,
  maxWorksheetBytes: 75 * 1024 * 1024
};

export type WorkbookArchiveLimits = {
  maxCompressedBytes?: number;
  maxCompressionRatio?: number;
  maxEntries?: number;
  maxEntryBytes?: number;
  maxUncompressedBytes?: number;
  maxWorksheetBytes?: number;
};

export type WorkbookArchiveIssueCode =
  | "ARCHIVE_COMPRESSED_SIZE_LIMIT"
  | "ARCHIVE_COMPRESSION_RATIO_LIMIT"
  | "ARCHIVE_ENTRY_COUNT_LIMIT"
  | "ARCHIVE_ENTRY_SIZE_LIMIT"
  | "ARCHIVE_UNCOMPRESSED_SIZE_LIMIT"
  | "ARCHIVE_UNSAFE_PATH"
  | "ARCHIVE_WORKSHEET_SIZE_LIMIT";

export type WorkbookArchiveIssue = {
  code: WorkbookArchiveIssueCode;
  message: string;
  entry?: string;
};

export type WorkbookArchiveInspection = {
  accepted: boolean;
  compressedBytes: number;
  compressionRatio: number;
  entryCount: number;
  issues: WorkbookArchiveIssue[];
  largestEntry?: { name: string; uncompressedBytes: number };
  largestWorksheet?: { name: string; uncompressedBytes: number };
  uncompressedBytes: number;
};

export type SafeWorkbookBlobWriteResult = {
  blob?: Blob;
  validation: ValidationReport;
  wrote: boolean;
};

export const browserCompressionAdapter: CompressionAdapter = {
  async inflateRaw(data) {
    return pipeBytesThroughCompressionStream(data, "deflate-raw", "decompress");
  },
  async deflateRaw(data) {
    return pipeBytesThroughCompressionStream(data, "deflate-raw", "compress");
  }
};

export async function openWorkbookFromBlob(blob: Blob): Promise<Workbook> {
  return openWorkbookFromArrayBuffer(await blob.arrayBuffer());
}

export async function openWorkbookFromArrayBuffer(data: ArrayBuffer): Promise<Workbook> {
  return openWorkbookFromBytes(new Uint8Array(data));
}

export async function openWorkbookFromBytes(data: Uint8Array): Promise<Workbook> {
  return Workbook.fromPackage(OoxmlPackage.open(data, browserCompressionAdapter));
}

export async function inspectWorkbookArchiveFromBlob(
  blob: Blob,
  limits: WorkbookArchiveLimits = {}
): Promise<WorkbookArchiveInspection> {
  return inspectWorkbookArchiveFromBytes(new Uint8Array(await blob.arrayBuffer()), limits);
}

export function inspectWorkbookArchiveFromBytes(
  data: Uint8Array,
  limits: WorkbookArchiveLimits = {}
): WorkbookArchiveInspection {
  const resolvedLimits = { ...defaultArchiveLimits, ...limits };
  const entries = parseZip(data).entries;
  const uncompressedBytes = entries.reduce((total, entry) => total + entry.uncompressedSize, 0);
  const compressionRatio = uncompressedBytes / Math.max(1, data.byteLength);
  const largestEntry = entries.reduce<(typeof entries)[number] | undefined>(
    (largest, entry) =>
      largest === undefined || entry.uncompressedSize > largest.uncompressedSize ? entry : largest,
    undefined
  );
  const largestWorksheet = entries
    .filter((entry) => /^xl\/worksheets\/.+\.xml$/i.test(entry.name))
    .reduce<(typeof entries)[number] | undefined>(
      (largest, entry) =>
        largest === undefined || entry.uncompressedSize > largest.uncompressedSize
          ? entry
          : largest,
      undefined
    );
  const issues: WorkbookArchiveIssue[] = [];

  if (data.byteLength > resolvedLimits.maxCompressedBytes) {
    issues.push({
      code: "ARCHIVE_COMPRESSED_SIZE_LIMIT",
      message: `Workbook is ${data.byteLength} bytes compressed; the configured limit is ${resolvedLimits.maxCompressedBytes} bytes`
    });
  }
  if (entries.length > resolvedLimits.maxEntries) {
    issues.push({
      code: "ARCHIVE_ENTRY_COUNT_LIMIT",
      message: `Workbook contains ${entries.length} ZIP entries; the configured limit is ${resolvedLimits.maxEntries}`
    });
  }
  if (uncompressedBytes > resolvedLimits.maxUncompressedBytes) {
    issues.push({
      code: "ARCHIVE_UNCOMPRESSED_SIZE_LIMIT",
      message: `Workbook expands to ${uncompressedBytes} bytes; the configured limit is ${resolvedLimits.maxUncompressedBytes} bytes`
    });
  }
  if (compressionRatio > resolvedLimits.maxCompressionRatio) {
    issues.push({
      code: "ARCHIVE_COMPRESSION_RATIO_LIMIT",
      message: `Workbook compression ratio is ${compressionRatio.toFixed(1)}; the configured limit is ${resolvedLimits.maxCompressionRatio}`
    });
  }
  if (largestEntry !== undefined && largestEntry.uncompressedSize > resolvedLimits.maxEntryBytes) {
    issues.push({
      code: "ARCHIVE_ENTRY_SIZE_LIMIT",
      entry: largestEntry.name,
      message: `${largestEntry.name} expands to ${largestEntry.uncompressedSize} bytes; the configured per-entry limit is ${resolvedLimits.maxEntryBytes} bytes`
    });
  }
  if (
    largestWorksheet !== undefined &&
    largestWorksheet.uncompressedSize > resolvedLimits.maxWorksheetBytes
  ) {
    issues.push({
      code: "ARCHIVE_WORKSHEET_SIZE_LIMIT",
      entry: largestWorksheet.name,
      message: `${largestWorksheet.name} expands to ${largestWorksheet.uncompressedSize} bytes; the configured worksheet limit is ${resolvedLimits.maxWorksheetBytes} bytes`
    });
  }

  for (const entry of entries) {
    if (entry.name.startsWith("/") || entry.name.split("/").includes("..")) {
      issues.push({
        code: "ARCHIVE_UNSAFE_PATH",
        entry: entry.name,
        message: `Workbook contains an unsafe ZIP entry path: ${entry.name}`
      });
    }
  }

  return {
    accepted: issues.length === 0,
    compressedBytes: data.byteLength,
    compressionRatio,
    entryCount: entries.length,
    issues,
    ...(largestEntry === undefined
      ? {}
      : {
          largestEntry: {
            name: largestEntry.name,
            uncompressedBytes: largestEntry.uncompressedSize
          }
        }),
    ...(largestWorksheet === undefined
      ? {}
      : {
          largestWorksheet: {
            name: largestWorksheet.name,
            uncompressedBytes: largestWorksheet.uncompressedSize
          }
        }),
    uncompressedBytes
  };
}

export async function writeWorkbookToBlob(
  workbook: Workbook,
  options: { type?: string } = {}
): Promise<Blob> {
  const data = await workbook.write();
  return new Blob([bytesToArrayBuffer(data)], {
    type: options.type ?? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

export async function writeWorkbookToBlobSafely(
  workbook: Workbook,
  options: { type?: string } = {}
): Promise<SafeWorkbookBlobWriteResult> {
  const validation = await workbook.validate();
  if (validation.summary.errors > 0) {
    return { validation, wrote: false };
  }

  return {
    blob: await writeWorkbookToBlob(workbook, options),
    validation,
    wrote: true
  };
}

async function pipeBytesThroughCompressionStream(
  data: Uint8Array,
  format: CompressionFormat,
  mode: "compress" | "decompress"
): Promise<Uint8Array> {
  const stream =
    mode === "compress" ? new CompressionStream(format) : new DecompressionStream(format);
  const writer = stream.writable.getWriter();
  const outputPromise = readStreamBytes(stream.readable);
  await writer.write(new Uint8Array(bytesToArrayBuffer(data)));
  await writer.close();

  return outputPromise;
}

async function readStreamBytes(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let totalBytes = 0;

  while (true) {
    const result = await reader.read();
    if (result.done) {
      break;
    }

    chunks.push(result.value);
    totalBytes += result.value.byteLength;
  }

  const output = new Uint8Array(totalBytes);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }

  return output;
}

function bytesToArrayBuffer(data: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(data.byteLength);
  copy.set(data);
  return copy.buffer;
}
