import assert from "node:assert/strict";
import test from "node:test";
import { parseZip, readEntryData, writeZip } from "../packages/core/src/index.ts";
import {
  browserCompressionAdapter,
  inspectWorkbookArchiveFromBytes,
  openWorkbookFromBytes,
  writeWorkbookToBlobSafely
} from "../packages/browser/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

test(
  "browser compression adapter deflates and inflates ZIP entries",
  { skip: !supportsRawDeflateCompressionStreams() },
  async () => {
    const zip = await writeZip(
      [
        {
          name: "xl/workbook.xml",
          data: textEncoder.encode("browser-ready"),
          compressionMethod: 8
        }
      ],
      browserCompressionAdapter
    );
    const entry = parseZip(zip).entries[0];
    assert.ok(entry);

    const data = await readEntryData(entry, browserCompressionAdapter);
    assert.equal(textDecoder.decode(data), "browser-ready");
  }
);

test(
  "browser compression adapter drains large entries without blocking on backpressure",
  { skip: !supportsRawDeflateCompressionStreams(), timeout: 5_000 },
  async () => {
    const input = createDeterministicBytes(256 * 1024);

    const compressed = await browserCompressionAdapter.deflateRaw(input);
    const output = await browserCompressionAdapter.inflateRaw(compressed);

    assert.deepEqual(output, input);
  }
);

test("browser archive inspection reports workbook resource usage and limits", async () => {
  const bytes = await createMinimalWorkbook({ includeDrawing: true, includeTable: true });
  const inspection = inspectWorkbookArchiveFromBytes(bytes);

  assert.equal(inspection.accepted, true);
  assert.equal(inspection.compressedBytes, bytes.byteLength);
  assert.ok(inspection.uncompressedBytes > 0);
  assert.ok(inspection.entryCount > 0);
  assert.equal(inspection.largestWorksheet?.name, "xl/worksheets/sheet1.xml");

  const limited = inspectWorkbookArchiveFromBytes(bytes, { maxCompressedBytes: 1 });
  assert.equal(limited.accepted, false);
  assert.equal(limited.issues[0]?.code, "ARCHIVE_COMPRESSED_SIZE_LIMIT");
});

test(
  "safe browser writes refuse invalid workbooks and return validated blobs",
  { skip: !supportsRawDeflateCompressionStreams() },
  async () => {
    const bytes = await createMinimalWorkbook({ includeTable: true });
    const workbook = await openWorkbookFromBytes(bytes);
    await workbook.replaceTableRows("RevenueTable", [["Fresh", 42]]);

    const written = await writeWorkbookToBlobSafely(workbook);
    assert.equal(written.wrote, true);
    assert.equal(written.validation.summary.errors, 0);
    assert.ok(written.blob);

    const invalidWorkbook = await openWorkbookFromBytes(bytes);
    invalidWorkbook.pkg.deletePart("xl/worksheets/sheet1.xml");
    const refused = await writeWorkbookToBlobSafely(invalidWorkbook);
    assert.equal(refused.wrote, false);
    assert.ok(refused.validation.summary.errors > 0);
    assert.equal(refused.blob, undefined);
  }
);

function supportsRawDeflateCompressionStreams(): boolean {
  try {
    new CompressionStream("deflate-raw");
    new DecompressionStream("deflate-raw");
    return true;
  } catch {
    return false;
  }
}

function createDeterministicBytes(length: number): Uint8Array {
  const data = new Uint8Array(length);
  let state = 0x12345678;

  for (let index = 0; index < data.byteLength; index += 1) {
    state = (Math.imul(state, 1_664_525) + 1_013_904_223) >>> 0;
    data[index] = state >>> 24;
  }

  return data;
}
