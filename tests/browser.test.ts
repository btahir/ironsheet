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

function supportsRawDeflateCompressionStreams(): boolean {
  try {
    new CompressionStream("deflate-raw");
    new DecompressionStream("deflate-raw");
    return true;
  } catch {
    return false;
  }
}
