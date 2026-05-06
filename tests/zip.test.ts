import assert from "node:assert/strict";
import test from "node:test";
import { parseZip, writeZip, ZipError } from "../packages/core/src/index.ts";

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
