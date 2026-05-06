import assert from "node:assert/strict";
import test from "node:test";
import {
  findElementCloseStart,
  findFirstStartTag,
  findStartTags,
  PackageError,
  streamXmlElements,
  tokenizeXml,
  tokenizeXmlChunks,
  transformXmlChunks
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

test("XML chunk tokenizer emits tokens split across chunk boundaries", async () => {
  const tokens = [];
  for await (const token of tokenizeXmlChunks([
    "<?xml",
    ' version="1.0"?><ro',
    'ot><item a="1 > 0"/>te',
    "xt<![CDATA[<item/>]]><!-- <item/> --></root>"
  ])) {
    tokens.push(token);
  }

  assert.deepEqual(
    tokens.map((token) => token.kind),
    ["processingInstruction", "start", "start", "text", "text", "cdata", "comment", "end"]
  );
  assert.equal(tokens[1]?.kind, "start");
  assert.equal(tokens[1]?.kind === "start" ? tokens[1].tag.start : undefined, 21);
  assert.equal(tokens[2]?.kind === "start" ? tokens[2].tag.attributes.a : undefined, "1 > 0");
  assert.equal(
    tokens
      .filter((token) => token.kind === "text")
      .map((token) => token.text)
      .join(""),
    "text"
  );
  assert.equal(tokens[5]?.kind === "cdata" ? tokens[5].text : undefined, "<item/>");
});

test("XML chunk transform passes through untouched tokens and replaces selected tokens", async () => {
  const chunks: string[] = [];
  for await (const chunk of transformXmlChunks(
    ["<root><v>old</v><!-- keep --><v>later</v></root>"],
    (token) => (token.kind === "text" && token.text === "old" ? "new" : undefined)
  )) {
    chunks.push(chunk);
  }

  assert.equal(chunks.join(""), "<root><v>new</v><!-- keep --><v>later</v></root>");
});

test("XML element stream extracts matching elements across chunks", async () => {
  const rows = [];
  for await (const row of streamXmlElements(
    [
      "<worksheet><sheetData><row r=",
      '"1"><c r="A1"><v>1</v></c></row><row r="2"/>',
      "</sheetData></worksheet>"
    ],
    "row"
  )) {
    rows.push(row);
  }

  assert.deepEqual(
    rows.map((row) => ({
      raw: row.raw,
      start: row.start,
      end: row.end,
      ref: row.tag.attributes.r
    })),
    [
      {
        raw: '<row r="1"><c r="A1"><v>1</v></c></row>',
        start: 22,
        end: 61,
        ref: "1"
      },
      {
        raw: '<row r="2"/>',
        start: 61,
        end: 73,
        ref: "2"
      }
    ]
  );
});

test("XML element stream reports an unclosed selected element", async () => {
  await assert.rejects(async () => {
    for await (const _row of streamXmlElements(['<worksheet><row r="1">'], "row")) {
      // drain stream
    }
  }, PackageError);
});

test("XML tokenizer reports unterminated quoted attributes", () => {
  assert.throws(() => [...tokenizeXml('<root name="nope></root>')], PackageError);
});

test("XML chunk tokenizer reports unterminated structures at final chunk", async () => {
  await assert.rejects(async () => {
    for await (const _token of tokenizeXmlChunks(["<root><!-- nope"])) {
      // drain stream
    }
  }, PackageError);
});

test("XML element close matching handles nested same-name elements", () => {
  const xml = "<root><item><item><v>nested</v></item><v>outer</v></item></root>";
  const item = findFirstStartTag(xml, "item");

  assert.ok(item);
  assert.equal(xml.slice(findElementCloseStart(xml, item)), "</item></root>");
});
