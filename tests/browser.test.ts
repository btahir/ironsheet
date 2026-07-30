import assert from "node:assert/strict";
import test from "node:test";
import { parseZip, readEntryData, writeZip } from "../packages/core/src/index.ts";
import { browserCompressionAdapter } from "../packages/browser/src/index.ts";

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
