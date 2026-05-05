import assert from "node:assert/strict";
import test from "node:test";
import { PackageError } from "../packages/core/src/index.ts";
import { findStartTags } from "../packages/core/src/xml.ts";

test("XML scanner ignores tags inside comments and CDATA", () => {
  const tags = findStartTags(
    `<worksheet>
      <!-- <c r="A1"><v>not a cell</v></c> -->
      <![CDATA[<c r="B2"><v>also not a cell</v></c>]]>
      <sheetData><row r="1"><c r="C3"><v>1</v></c></row></sheetData>
    </worksheet>`,
    "c"
  );

  assert.deepEqual(
    tags.map((tag) => tag.attributes.r),
    ["C3"]
  );
});

test("XML scanner skips processing instructions and closing tags", () => {
  const tags = findStartTags(
    `<?xml version="1.0"?>
    <root><item id="1"></item><item id="2"/></root>`,
    "item"
  );

  assert.deepEqual(
    tags.map((tag) => tag.attributes.id),
    ["1", "2"]
  );
});

test("XML scanner reports unterminated comments", () => {
  assert.throws(() => findStartTags("<root><!-- nope", "item"), PackageError);
});
