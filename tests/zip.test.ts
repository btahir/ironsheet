import assert from "node:assert/strict";
import test from "node:test";
import { crc32, parseZip, readEntryData, writeZip, ZipError } from "../packages/core/src/index.ts";

const textEncoder = new TextEncoder();
const storedEntryCompression = {
  inflateRaw(): Uint8Array {
    throw new Error("Stored entries should not inflate");
  },
  deflateRaw(): Uint8Array {
    throw new Error("Stored entries should not deflate");
  }
};

test("ZIP writer rejects duplicate entry names", async () => {
  await assert.rejects(
    () =>
      writeZip([
        { name: "xl/workbook.xml", data: new Uint8Array([1]) },
        { name: "xl/workbook.xml", data: new Uint8Array([2]) }
      ]),
    ZipError
  );
});

test("ZIP writer rejects unsafe entry paths", async () => {
  await assert.rejects(
    () => writeZip([{ name: "../xl/workbook.xml", data: new Uint8Array([1]) }]),
    ZipError
  );
  await assert.rejects(
    () => writeZip([{ name: "xl\\workbook.xml", data: new Uint8Array([1]) }]),
    ZipError
  );
});

test("ZIP reader rejects unsafe entry paths emitted by an external archive", async () => {
  const archive = await writeZip([{ name: "xl/workbook.xml", data: new Uint8Array([1]) }]);
  const unsafeArchive = new Uint8Array(archive);
  const safeName = new TextEncoder().encode("xl/workbook.xml");
  const unsafeName = new TextEncoder().encode("../workbook.xml");
  const replacements = replaceAllBytes(unsafeArchive, safeName, unsafeName);
  assert.equal(replacements, 2);

  assert.throws(() => parseZip(unsafeArchive), ZipError);
});

test("ZIP reader parses ZIP64 end of central directory metadata", async () => {
  const archive = zip64Archive();
  const parsed = parseZip(archive);

  assert.deepEqual(
    parsed.entries.map((entry) => entry.name),
    ["xl/workbook.xml"]
  );
  const entry = parsed.entries[0];
  assert.ok(entry);
  assert.deepEqual(Array.from(await readEntryData(entry, storedEntryCompression)), [1, 2, 3]);
});

test("ZIP reader parses ZIP64 central directory entry sizes and offsets", async () => {
  const archive = zip64Archive({ zip64Entry: true });
  const parsed = parseZip(archive);
  const entry = parsed.entries[0];
  assert.ok(entry);

  assert.equal(entry.compressedSize, 3);
  assert.equal(entry.uncompressedSize, 3);
  assert.deepEqual(Array.from(await readEntryData(entry, storedEntryCompression)), [1, 2, 3]);
});

test("ZIP reader rejects ZIP64 archives missing the locator", async () => {
  const archive = await writeZip([{ name: "xl/workbook.xml", data: new Uint8Array([1]) }]);
  const mutated = new Uint8Array(archive);
  const eocdOffset = findSignature(mutated, 0x06054b50);
  writeUInt16LE(mutated, eocdOffset + 8, 0xffff);
  writeUInt16LE(mutated, eocdOffset + 10, 0xffff);
  writeUInt32LE(mutated, eocdOffset + 12, 0xffffffff);
  writeUInt32LE(mutated, eocdOffset + 16, 0xffffffff);

  assert.throws(() => parseZip(mutated), ZipError);
});

test("ZIP reader rejects multi-disk ZIP64 archives", () => {
  assert.throws(() => parseZip(zip64Archive({ totalDisks: 2 })), ZipError);
});

test("ZIP reader rejects ZIP64 values beyond JavaScript safe integers", () => {
  assert.throws(
    () => parseZip(zip64Archive({ centralDirectoryOffset: BigInt(Number.MAX_SAFE_INTEGER) + 1n })),
    ZipError
  );
});

test("ZIP writer rejects entries that would require ZIP64 output", async () => {
  await assert.rejects(
    () =>
      writeZip([
        {
          name: "xl/workbook.xml",
          compressedData: new Uint8Array([1]),
          compressionMethod: 0,
          crc32: 1,
          uncompressedSize: 0x1_0000_0000
        }
      ]),
    ZipError
  );
});

function replaceAllBytes(
  haystack: Uint8Array,
  needle: Uint8Array,
  replacement: Uint8Array
): number {
  assert.equal(needle.byteLength, replacement.byteLength);
  let replacements = 0;
  for (let offset = 0; offset <= haystack.byteLength - needle.byteLength; offset += 1) {
    if (needle.every((byte, index) => haystack[offset + index] === byte)) {
      haystack.set(replacement, offset);
      replacements += 1;
      offset += needle.byteLength - 1;
    }
  }

  return replacements;
}

function zip64Archive(
  options: {
    zip64Entry?: boolean;
    totalDisks?: number;
    centralDirectoryOffset?: bigint;
  } = {}
): Uint8Array {
  const name = textEncoder.encode("xl/workbook.xml");
  const data = new Uint8Array([1, 2, 3]);
  const entryCrc = crc32(data);

  const localHeader = new Uint8Array(30 + name.byteLength);
  writeUInt32LE(localHeader, 0, 0x04034b50);
  writeUInt16LE(localHeader, 4, 20);
  writeUInt16LE(localHeader, 6, 1 << 11);
  writeUInt16LE(localHeader, 8, 0);
  writeUInt32LE(localHeader, 14, entryCrc);
  writeUInt32LE(localHeader, 18, data.byteLength);
  writeUInt32LE(localHeader, 22, data.byteLength);
  writeUInt16LE(localHeader, 26, name.byteLength);
  localHeader.set(name, 30);

  const localHeaderOffset = 0;
  const centralDirectoryOffset = localHeader.byteLength + data.byteLength;
  const zip64Extra =
    options.zip64Entry === true
      ? zip64EntryExtra(data.byteLength, data.byteLength, localHeaderOffset)
      : new Uint8Array();
  const centralHeader = new Uint8Array(46 + name.byteLength + zip64Extra.byteLength);
  writeUInt32LE(centralHeader, 0, 0x02014b50);
  writeUInt16LE(centralHeader, 4, 20);
  writeUInt16LE(centralHeader, 6, 20);
  writeUInt16LE(centralHeader, 8, 1 << 11);
  writeUInt16LE(centralHeader, 10, 0);
  writeUInt32LE(centralHeader, 16, entryCrc);
  writeUInt32LE(centralHeader, 20, options.zip64Entry === true ? 0xffffffff : data.byteLength);
  writeUInt32LE(centralHeader, 24, options.zip64Entry === true ? 0xffffffff : data.byteLength);
  writeUInt16LE(centralHeader, 28, name.byteLength);
  writeUInt16LE(centralHeader, 30, zip64Extra.byteLength);
  writeUInt32LE(centralHeader, 42, options.zip64Entry === true ? 0xffffffff : localHeaderOffset);
  centralHeader.set(name, 46);
  centralHeader.set(zip64Extra, 46 + name.byteLength);

  const zip64RecordOffset = centralDirectoryOffset + centralHeader.byteLength;
  const zip64Record = new Uint8Array(56);
  writeUInt32LE(zip64Record, 0, 0x06064b50);
  writeUInt64LE(zip64Record, 4, 44n);
  writeUInt16LE(zip64Record, 12, 45);
  writeUInt16LE(zip64Record, 14, 45);
  writeUInt64LE(zip64Record, 24, 1n);
  writeUInt64LE(zip64Record, 32, 1n);
  writeUInt64LE(zip64Record, 40, BigInt(centralHeader.byteLength));
  writeUInt64LE(zip64Record, 48, options.centralDirectoryOffset ?? BigInt(centralDirectoryOffset));

  const locator = new Uint8Array(20);
  writeUInt32LE(locator, 0, 0x07064b50);
  writeUInt64LE(locator, 8, BigInt(zip64RecordOffset));
  writeUInt32LE(locator, 16, options.totalDisks ?? 1);

  const eocd = new Uint8Array(22);
  writeUInt32LE(eocd, 0, 0x06054b50);
  writeUInt16LE(eocd, 8, 0xffff);
  writeUInt16LE(eocd, 10, 0xffff);
  writeUInt32LE(eocd, 12, 0xffffffff);
  writeUInt32LE(eocd, 16, 0xffffffff);

  return concatBytes([localHeader, data, centralHeader, zip64Record, locator, eocd]);
}

function zip64EntryExtra(
  uncompressedSize: number,
  compressedSize: number,
  localHeaderOffset: number
): Uint8Array {
  const extra = new Uint8Array(4 + 24);
  writeUInt16LE(extra, 0, 0x0001);
  writeUInt16LE(extra, 2, 24);
  writeUInt64LE(extra, 4, BigInt(uncompressedSize));
  writeUInt64LE(extra, 12, BigInt(compressedSize));
  writeUInt64LE(extra, 20, BigInt(localHeaderOffset));
  return extra;
}

function findSignature(data: Uint8Array, signature: number): number {
  for (let offset = data.byteLength - 4; offset >= 0; offset -= 1) {
    if (readUInt32LE(data, offset) === signature) {
      return offset;
    }
  }

  throw new Error(`Signature ${signature.toString(16)} not found`);
}

function concatBytes(chunks: Uint8Array[]): Uint8Array {
  const size = chunks.reduce((total, chunk) => total + chunk.byteLength, 0);
  const output = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
}

function readUInt32LE(data: Uint8Array, offset: number): number {
  return (
    (byteAt(data, offset) |
      (byteAt(data, offset + 1) << 8) |
      (byteAt(data, offset + 2) << 16) |
      (byteAt(data, offset + 3) << 24)) >>>
    0
  );
}

function byteAt(data: Uint8Array, offset: number): number {
  const byte = data[offset];
  if (byte === undefined) {
    throw new Error(`Unexpected end of test data at byte ${offset}`);
  }
  return byte;
}

function writeUInt16LE(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
}

function writeUInt32LE(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
  target[offset + 2] = (value >>> 16) & 0xff;
  target[offset + 3] = (value >>> 24) & 0xff;
}

function writeUInt64LE(target: Uint8Array, offset: number, value: bigint): void {
  for (let index = 0; index < 8; index += 1) {
    target[offset + index] = Number((value >> BigInt(index * 8)) & 0xffn);
  }
}
