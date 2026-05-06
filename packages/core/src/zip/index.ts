import { ZipError } from "../errors.ts";
import { concatBytes, readUInt16LE, readUInt32LE, writeUInt16LE, writeUInt32LE } from "./binary.ts";
import { crc32 } from "./crc32.ts";

const localFileHeaderSignature = 0x04034b50;
const centralDirectoryHeaderSignature = 0x02014b50;
const endOfCentralDirectorySignature = 0x06054b50;
const utf8FileNameFlag = 1 << 11;
const stored = 0;
const deflated = 8;
const maxCommentLength = 0xffff;

const textDecoder = new TextDecoder();
const textEncoder = new TextEncoder();

export type CompressionMethod = 0 | 8;

export type CompressionAdapter = {
  inflateRaw(data: Uint8Array): Uint8Array | Promise<Uint8Array>;
  deflateRaw(data: Uint8Array): Uint8Array | Promise<Uint8Array>;
};

export type ZipEntry = {
  name: string;
  compressionMethod: CompressionMethod;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  compressedData: Uint8Array;
  lastModTime: number;
  lastModDate: number;
  externalAttributes: number;
};

export type ZipFile = {
  entries: ZipEntry[];
};

export type ZipWriteEntry = {
  name: string;
  data?: Uint8Array;
  compressedData?: Uint8Array;
  compressionMethod?: CompressionMethod;
  crc32?: number;
  uncompressedSize?: number;
  lastModTime?: number;
  lastModDate?: number;
  externalAttributes?: number;
};

export { crc32 };

export function parseZip(data: Uint8Array): ZipFile {
  const eocdOffset = findEndOfCentralDirectory(data);
  const entryCount = readUInt16LE(data, eocdOffset + 10);
  const centralDirectorySize = readUInt32LE(data, eocdOffset + 12);
  const centralDirectoryOffset = readUInt32LE(data, eocdOffset + 16);

  if (entryCount === 0xffff || centralDirectorySize === 0xffffffff) {
    throw new ZipError("ZIP64 archives are not supported yet");
  }

  const entries: ZipEntry[] = [];
  let offset = centralDirectoryOffset;
  const endOffset = centralDirectoryOffset + centralDirectorySize;

  while (offset < endOffset) {
    if (readUInt32LE(data, offset) !== centralDirectoryHeaderSignature) {
      throw new ZipError(`Invalid central directory header at byte ${offset}`);
    }

    const flags = readUInt16LE(data, offset + 8);
    const compressionMethod = readUInt16LE(data, offset + 10);
    if (compressionMethod !== stored && compressionMethod !== deflated) {
      throw new ZipError(`Unsupported compression method ${compressionMethod}`);
    }

    const lastModTime = readUInt16LE(data, offset + 12);
    const lastModDate = readUInt16LE(data, offset + 14);
    const entryCrc32 = readUInt32LE(data, offset + 16);
    const compressedSize = readUInt32LE(data, offset + 20);
    const uncompressedSize = readUInt32LE(data, offset + 24);
    const fileNameLength = readUInt16LE(data, offset + 28);
    const extraLength = readUInt16LE(data, offset + 30);
    const commentLength = readUInt16LE(data, offset + 32);
    const externalAttributes = readUInt32LE(data, offset + 38);
    const localHeaderOffset = readUInt32LE(data, offset + 42);
    const fileNameStart = offset + 46;
    const fileNameEnd = fileNameStart + fileNameLength;
    const name = decodeFileName(data.subarray(fileNameStart, fileNameEnd), flags);

    const localHeader = parseLocalHeader(data, localHeaderOffset);
    const compressedData = data.subarray(
      localHeader.dataOffset,
      localHeader.dataOffset + compressedSize
    );

    entries.push({
      name,
      compressionMethod,
      crc32: entryCrc32,
      compressedSize,
      uncompressedSize,
      compressedData,
      lastModTime,
      lastModDate,
      externalAttributes
    });

    offset = fileNameEnd + extraLength + commentLength;
  }

  if (entries.length !== entryCount) {
    throw new ZipError(`Expected ${entryCount} entries but parsed ${entries.length}`);
  }

  assertUniqueEntryNames(entries);

  return { entries };
}

export async function readEntryData(
  entry: ZipEntry,
  compression: CompressionAdapter
): Promise<Uint8Array> {
  if (entry.compressionMethod === stored) {
    return entry.compressedData;
  }

  const inflated = await compression.inflateRaw(entry.compressedData);
  if (inflated.byteLength !== entry.uncompressedSize) {
    throw new ZipError(
      `Entry ${entry.name} inflated to ${inflated.byteLength} bytes, expected ${entry.uncompressedSize}`
    );
  }

  return inflated;
}

export async function writeZip(
  entries: ZipWriteEntry[],
  compression?: CompressionAdapter
): Promise<Uint8Array> {
  assertSafeEntryNames(entries.map((entry) => entry.name));

  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
    const prepared = await prepareEntry(entry, compression);
    const nameBytes = textEncoder.encode(prepared.name);
    const localHeader = new Uint8Array(30 + nameBytes.byteLength);

    writeUInt32LE(localHeader, 0, localFileHeaderSignature);
    writeUInt16LE(localHeader, 4, 20);
    writeUInt16LE(localHeader, 6, utf8FileNameFlag);
    writeUInt16LE(localHeader, 8, prepared.compressionMethod);
    writeUInt16LE(localHeader, 10, prepared.lastModTime);
    writeUInt16LE(localHeader, 12, prepared.lastModDate);
    writeUInt32LE(localHeader, 14, prepared.crc32);
    writeUInt32LE(localHeader, 18, prepared.compressedData.byteLength);
    writeUInt32LE(localHeader, 22, prepared.uncompressedSize);
    writeUInt16LE(localHeader, 26, nameBytes.byteLength);
    writeUInt16LE(localHeader, 28, 0);
    localHeader.set(nameBytes, 30);

    localParts.push(localHeader, prepared.compressedData);

    const centralHeader = new Uint8Array(46 + nameBytes.byteLength);
    writeUInt32LE(centralHeader, 0, centralDirectoryHeaderSignature);
    writeUInt16LE(centralHeader, 4, 20);
    writeUInt16LE(centralHeader, 6, 20);
    writeUInt16LE(centralHeader, 8, utf8FileNameFlag);
    writeUInt16LE(centralHeader, 10, prepared.compressionMethod);
    writeUInt16LE(centralHeader, 12, prepared.lastModTime);
    writeUInt16LE(centralHeader, 14, prepared.lastModDate);
    writeUInt32LE(centralHeader, 16, prepared.crc32);
    writeUInt32LE(centralHeader, 20, prepared.compressedData.byteLength);
    writeUInt32LE(centralHeader, 24, prepared.uncompressedSize);
    writeUInt16LE(centralHeader, 28, nameBytes.byteLength);
    writeUInt16LE(centralHeader, 30, 0);
    writeUInt16LE(centralHeader, 32, 0);
    writeUInt16LE(centralHeader, 34, 0);
    writeUInt16LE(centralHeader, 36, 0);
    writeUInt32LE(centralHeader, 38, prepared.externalAttributes);
    writeUInt32LE(centralHeader, 42, offset);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);

    offset += localHeader.byteLength + prepared.compressedData.byteLength;
  }

  const centralDirectoryOffset = offset;
  const centralDirectory = concatBytes(centralParts);
  const eocd = new Uint8Array(22);

  writeUInt32LE(eocd, 0, endOfCentralDirectorySignature);
  writeUInt16LE(eocd, 4, 0);
  writeUInt16LE(eocd, 6, 0);
  writeUInt16LE(eocd, 8, entries.length);
  writeUInt16LE(eocd, 10, entries.length);
  writeUInt32LE(eocd, 12, centralDirectory.byteLength);
  writeUInt32LE(eocd, 16, centralDirectoryOffset);
  writeUInt16LE(eocd, 20, 0);

  return concatBytes([...localParts, centralDirectory, eocd]);
}

function findEndOfCentralDirectory(data: Uint8Array): number {
  const minOffset = Math.max(0, data.byteLength - 22 - maxCommentLength);

  for (let offset = data.byteLength - 22; offset >= minOffset; offset -= 1) {
    if (readUInt32LE(data, offset) === endOfCentralDirectorySignature) {
      return offset;
    }
  }

  throw new ZipError("End of central directory not found");
}

function parseLocalHeader(data: Uint8Array, offset: number): { dataOffset: number } {
  if (readUInt32LE(data, offset) !== localFileHeaderSignature) {
    throw new ZipError(`Invalid local file header at byte ${offset}`);
  }

  const fileNameLength = readUInt16LE(data, offset + 26);
  const extraLength = readUInt16LE(data, offset + 28);
  return { dataOffset: offset + 30 + fileNameLength + extraLength };
}

function decodeFileName(data: Uint8Array, flags: number): string {
  if ((flags & utf8FileNameFlag) === 0) {
    return textDecoder.decode(data);
  }

  return textDecoder.decode(data);
}

function assertUniqueEntryNames(entries: ZipEntry[]): void {
  assertSafeEntryNames(entries.map((entry) => entry.name));
}

function assertSafeEntryNames(names: string[]): void {
  const seen = new Set<string>();

  for (const name of names) {
    if (isUnsafeEntryName(name)) {
      throw new ZipError(`Unsafe ZIP entry path: ${name}`);
    }

    if (seen.has(name)) {
      throw new ZipError(`Duplicate ZIP entry: ${name}`);
    }

    seen.add(name);
  }
}

function isUnsafeEntryName(name: string): boolean {
  return (
    name.startsWith("/") ||
    name.includes("\\") ||
    name.split("/").some((segment) => segment === "..")
  );
}

async function prepareEntry(
  entry: ZipWriteEntry,
  compression?: CompressionAdapter
): Promise<Required<ZipWriteEntry>> {
  if (entry.compressedData !== undefined) {
    return {
      name: entry.name,
      data: entry.data ?? new Uint8Array(),
      compressedData: entry.compressedData,
      compressionMethod: entry.compressionMethod ?? stored,
      crc32: entry.crc32 ?? crc32(entry.compressedData),
      uncompressedSize: entry.uncompressedSize ?? entry.compressedData.byteLength,
      lastModTime: entry.lastModTime ?? 0,
      lastModDate: entry.lastModDate ?? 0,
      externalAttributes: entry.externalAttributes ?? 0
    };
  }

  if (entry.data === undefined) {
    throw new ZipError(`Entry ${entry.name} has neither data nor compressedData`);
  }

  const method = entry.compressionMethod ?? (compression === undefined ? stored : deflated);
  const compressedData =
    method === stored ? entry.data : await requireCompression(compression).deflateRaw(entry.data);

  return {
    name: entry.name,
    data: entry.data,
    compressedData,
    compressionMethod: method,
    crc32: entry.crc32 ?? crc32(entry.data),
    uncompressedSize: entry.uncompressedSize ?? entry.data.byteLength,
    lastModTime: entry.lastModTime ?? 0,
    lastModDate: entry.lastModDate ?? 0,
    externalAttributes: entry.externalAttributes ?? 0
  };
}

function requireCompression(compression: CompressionAdapter | undefined): CompressionAdapter {
  if (compression === undefined) {
    throw new ZipError("Compression adapter is required for deflated entries");
  }

  return compression;
}
