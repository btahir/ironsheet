import { ZipError } from "../errors.ts";
import {
  concatBytes,
  readUInt16LE,
  readUInt32LE,
  readUInt64LE,
  writeUInt16LE,
  writeUInt32LE
} from "./binary.ts";
import { crc32 } from "./crc32.ts";

const localFileHeaderSignature = 0x04034b50;
const centralDirectoryHeaderSignature = 0x02014b50;
const endOfCentralDirectorySignature = 0x06054b50;
const zip64EndOfCentralDirectorySignature = 0x06064b50;
const zip64EndOfCentralDirectoryLocatorSignature = 0x07064b50;
const utf8FileNameFlag = 1 << 11;
const stored = 0;
const deflated = 8;
const maxCommentLength = 0xffff;
const maxUInt16 = 0xffff;
const maxUInt32 = 0xffffffff;

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
  assertAvailable(data, eocdOffset, 22, "end of central directory");
  const diskNumber = readUInt16LE(data, eocdOffset + 4);
  const centralDirectoryDisk = readUInt16LE(data, eocdOffset + 6);
  const diskEntryCount = readUInt16LE(data, eocdOffset + 8);
  const totalEntryCount = readUInt16LE(data, eocdOffset + 10);
  const centralDirectorySize32 = readUInt32LE(data, eocdOffset + 12);
  const centralDirectoryOffset32 = readUInt32LE(data, eocdOffset + 16);

  if (diskNumber !== 0 || centralDirectoryDisk !== 0) {
    throw new ZipError("Multi-disk ZIP archives are not supported");
  }

  let entryCount = totalEntryCount;
  let centralDirectorySize = centralDirectorySize32;
  let centralDirectoryOffset = centralDirectoryOffset32;

  if (
    diskEntryCount === maxUInt16 ||
    totalEntryCount === maxUInt16 ||
    centralDirectorySize32 === maxUInt32 ||
    centralDirectoryOffset32 === maxUInt32
  ) {
    const zip64 = parseZip64EndOfCentralDirectory(data, eocdOffset);
    entryCount = zip64.entryCount;
    centralDirectorySize = zip64.centralDirectorySize;
    centralDirectoryOffset = zip64.centralDirectoryOffset;
  }

  const entries: ZipEntry[] = [];
  let offset = centralDirectoryOffset;
  const endOffset = centralDirectoryOffset + centralDirectorySize;
  assertAvailable(data, centralDirectoryOffset, centralDirectorySize, "central directory");

  while (offset < endOffset) {
    assertAvailable(data, offset, 46, "central directory header");
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
    const compressedSize32 = readUInt32LE(data, offset + 20);
    const uncompressedSize32 = readUInt32LE(data, offset + 24);
    const fileNameLength = readUInt16LE(data, offset + 28);
    const extraLength = readUInt16LE(data, offset + 30);
    const commentLength = readUInt16LE(data, offset + 32);
    const diskStart32 = readUInt16LE(data, offset + 34);
    const externalAttributes = readUInt32LE(data, offset + 38);
    const localHeaderOffset32 = readUInt32LE(data, offset + 42);
    const fileNameStart = offset + 46;
    const fileNameEnd = fileNameStart + fileNameLength;
    const extraStart = fileNameEnd;
    const extraEnd = extraStart + extraLength;
    const entryEnd = extraEnd + commentLength;
    assertAvailable(
      data,
      fileNameStart,
      fileNameLength + extraLength + commentLength,
      "central directory entry variable data"
    );
    const name = decodeFileName(data.subarray(fileNameStart, fileNameEnd), flags);
    const zip64 = parseZip64ExtendedInformation(data.subarray(extraStart, extraEnd), {
      compressedSize: compressedSize32 === maxUInt32,
      uncompressedSize: uncompressedSize32 === maxUInt32,
      localHeaderOffset: localHeaderOffset32 === maxUInt32,
      diskStart: diskStart32 === maxUInt16
    });
    const compressedSize = compressedSize32 === maxUInt32 ? zip64.compressedSize : compressedSize32;
    const uncompressedSize =
      uncompressedSize32 === maxUInt32 ? zip64.uncompressedSize : uncompressedSize32;
    const localHeaderOffset =
      localHeaderOffset32 === maxUInt32 ? zip64.localHeaderOffset : localHeaderOffset32;
    const diskStart = diskStart32 === maxUInt16 ? zip64.diskStart : diskStart32;

    if (diskStart !== 0) {
      throw new ZipError("Multi-disk ZIP archives are not supported");
    }

    const localHeader = parseLocalHeader(data, localHeaderOffset);
    assertAvailable(data, localHeader.dataOffset, compressedSize, `compressed data for ${name}`);
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

    offset = entryEnd;
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
  if (entries.length > maxUInt16) {
    throw new ZipError("ZIP64 writing is not supported for archives with more than 65535 entries");
  }

  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
    const prepared = await prepareEntry(entry, compression);
    const nameBytes = textEncoder.encode(prepared.name);
    assertWriteUInt16(nameBytes.byteLength, `entry name length for ${prepared.name}`);
    assertWriteUInt32(prepared.compressedData.byteLength, `compressed size for ${prepared.name}`);
    assertWriteUInt32(prepared.uncompressedSize, `uncompressed size for ${prepared.name}`);
    assertWriteUInt32(offset, `local header offset for ${prepared.name}`);
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
  assertWriteUInt32(centralDirectory.byteLength, "central directory size");
  assertWriteUInt32(centralDirectoryOffset, "central directory offset");
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

function parseZip64EndOfCentralDirectory(
  data: Uint8Array,
  eocdOffset: number
): {
  entryCount: number;
  centralDirectorySize: number;
  centralDirectoryOffset: number;
} {
  const locatorOffset = eocdOffset - 20;
  if (locatorOffset < 0) {
    throw new ZipError("ZIP64 end of central directory locator not found");
  }

  assertAvailable(data, locatorOffset, 20, "ZIP64 end of central directory locator");
  if (readUInt32LE(data, locatorOffset) !== zip64EndOfCentralDirectoryLocatorSignature) {
    throw new ZipError("ZIP64 end of central directory locator not found");
  }

  const zip64EocdDisk = readUInt32LE(data, locatorOffset + 4);
  const zip64EocdOffset = readSafeUInt64LE(data, locatorOffset + 8, "ZIP64 EOCD offset");
  const totalDisks = readUInt32LE(data, locatorOffset + 16);
  if (zip64EocdDisk !== 0 || totalDisks !== 1) {
    throw new ZipError("Multi-disk ZIP64 archives are not supported");
  }

  assertAvailable(data, zip64EocdOffset, 56, "ZIP64 end of central directory record");
  if (readUInt32LE(data, zip64EocdOffset) !== zip64EndOfCentralDirectorySignature) {
    throw new ZipError(`Invalid ZIP64 end of central directory record at byte ${zip64EocdOffset}`);
  }

  const recordSize = readSafeUInt64LE(data, zip64EocdOffset + 4, "ZIP64 EOCD record size");
  if (recordSize < 44) {
    throw new ZipError("ZIP64 end of central directory record is too small");
  }
  assertAvailable(data, zip64EocdOffset, 12 + recordSize, "ZIP64 end of central directory record");

  const diskNumber = readUInt32LE(data, zip64EocdOffset + 16);
  const centralDirectoryDisk = readUInt32LE(data, zip64EocdOffset + 20);
  const diskEntryCount = readSafeUInt64LE(data, zip64EocdOffset + 24, "ZIP64 entry count");
  const entryCount = readSafeUInt64LE(data, zip64EocdOffset + 32, "ZIP64 entry count");
  if (diskNumber !== 0 || centralDirectoryDisk !== 0 || diskEntryCount !== entryCount) {
    throw new ZipError("Multi-disk ZIP64 archives are not supported");
  }

  return {
    entryCount,
    centralDirectorySize: readSafeUInt64LE(
      data,
      zip64EocdOffset + 40,
      "ZIP64 central directory size"
    ),
    centralDirectoryOffset: readSafeUInt64LE(
      data,
      zip64EocdOffset + 48,
      "ZIP64 central directory offset"
    )
  };
}

function parseZip64ExtendedInformation(
  extra: Uint8Array,
  required: {
    compressedSize: boolean;
    uncompressedSize: boolean;
    localHeaderOffset: boolean;
    diskStart: boolean;
  }
): {
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
  diskStart: number;
} {
  const needsZip64 =
    required.compressedSize ||
    required.uncompressedSize ||
    required.localHeaderOffset ||
    required.diskStart;
  if (!needsZip64) {
    return { compressedSize: 0, uncompressedSize: 0, localHeaderOffset: 0, diskStart: 0 };
  }

  let offset = 0;
  while (offset < extra.byteLength) {
    if (offset + 4 > extra.byteLength) {
      throw new ZipError("Malformed ZIP extra field");
    }

    const headerId = readUInt16LE(extra, offset);
    const dataSize = readUInt16LE(extra, offset + 2);
    const dataStart = offset + 4;
    const dataEnd = dataStart + dataSize;
    if (dataEnd > extra.byteLength) {
      throw new ZipError("Malformed ZIP extra field");
    }

    if (headerId === 0x0001) {
      let cursor = dataStart;
      const readRequiredUInt64 = (label: string): number => {
        if (cursor + 8 > dataEnd) {
          throw new ZipError(`ZIP64 extended information is missing ${label}`);
        }

        const value = readSafeUInt64LE(extra, cursor, label);
        cursor += 8;
        return value;
      };
      const readRequiredUInt32 = (label: string): number => {
        if (cursor + 4 > dataEnd) {
          throw new ZipError(`ZIP64 extended information is missing ${label}`);
        }

        const value = readUInt32LE(extra, cursor);
        cursor += 4;
        return value;
      };

      const uncompressedSize = required.uncompressedSize
        ? readRequiredUInt64("uncompressed size")
        : 0;
      const compressedSize = required.compressedSize ? readRequiredUInt64("compressed size") : 0;
      const localHeaderOffset = required.localHeaderOffset
        ? readRequiredUInt64("local header offset")
        : 0;
      const diskStart = required.diskStart ? readRequiredUInt32("disk start") : 0;
      return { compressedSize, uncompressedSize, localHeaderOffset, diskStart };
    }

    offset = dataEnd;
  }

  throw new ZipError("ZIP64 extended information extra field not found");
}

function parseLocalHeader(data: Uint8Array, offset: number): { dataOffset: number } {
  assertAvailable(data, offset, 30, "local file header");
  if (readUInt32LE(data, offset) !== localFileHeaderSignature) {
    throw new ZipError(`Invalid local file header at byte ${offset}`);
  }

  const fileNameLength = readUInt16LE(data, offset + 26);
  const extraLength = readUInt16LE(data, offset + 28);
  assertAvailable(
    data,
    offset + 30,
    fileNameLength + extraLength,
    "local file header variable data"
  );
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

function assertAvailable(data: Uint8Array, offset: number, length: number, context: string): void {
  if (
    !Number.isSafeInteger(offset) ||
    !Number.isSafeInteger(length) ||
    offset < 0 ||
    length < 0 ||
    offset > data.byteLength ||
    length > data.byteLength - offset
  ) {
    throw new ZipError(`Unexpected end of data while reading ${context}`);
  }
}

function readSafeUInt64LE(data: Uint8Array, offset: number, label: string): number {
  const value = readUInt64LE(data, offset);
  if (value > BigInt(Number.MAX_SAFE_INTEGER)) {
    throw new ZipError(`${label} exceeds JavaScript's safe integer range`);
  }

  return Number(value);
}

function assertWriteUInt16(value: number, label: string): void {
  if (!Number.isSafeInteger(value) || value < 0 || value > maxUInt16) {
    throw new ZipError(`ZIP64 writing is not supported: ${label} exceeds 16-bit ZIP limit`);
  }
}

function assertWriteUInt32(value: number, label: string): void {
  if (!Number.isSafeInteger(value) || value < 0 || value > maxUInt32) {
    throw new ZipError(`ZIP64 writing is not supported: ${label} exceeds 32-bit ZIP limit`);
  }
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
