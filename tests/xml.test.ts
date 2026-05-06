import assert from "node:assert/strict";
import test from "node:test";
import {
  findElementCloseStart,
  findFirstStartTag,
  findStartTags,
  PackageError,
  tokenizeXml
} from "../packages/core/src/index.ts";

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

test("XML tokenizer emits structural tokens without parsing ignored markup", () => {
  const tokens = [
    ...tokenizeXml(
      '<?xml version="1.0"?><root a="1 > 0"><item/>text<![CDATA[<item/>]]><!-- <item/> --></root>'
    )
  ];

  assert.deepEqual(
    tokens.map((token) => token.kind),
    ["processingInstruction", "start", "start", "text", "cdata", "comment", "end"]
  );
  assert.equal(tokens[1]?.kind, "start");
  assert.equal(tokens[1]?.kind === "start" ? tokens[1].tag.attributes.a : undefined, "1 > 0");
  assert.equal(tokens[4]?.kind === "cdata" ? tokens[4].text : undefined, "<item/>");
});

test("XML tokenizer reports unterminated quoted attributes", () => {
  assert.throws(() => [...tokenizeXml('<root name="nope></root>')], PackageError);
});

test("XML element close matching handles nested same-name elements", () => {
  const xml = "<root><item><item><v>nested</v></item><v>outer</v></item></root>";
  const item = findFirstStartTag(xml, "item");

  assert.ok(item);
  assert.equal(xml.slice(findElementCloseStart(xml, item)), "</item></root>");
});
