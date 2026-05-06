import assert from "node:assert/strict";
import test from "node:test";
import { streamRowsXml, streamWorksheetRowsXml } from "../packages/core/src/index.ts";

test("streams row XML one row at a time from async data", async () => {
  async function* rows(): AsyncGenerator<Array<string | number | boolean>> {
    yield ["North", 10];
    yield ["South", true];
  }

  const chunks: string[] = [];
  for await (const chunk of streamRowsXml(rows(), { startColumn: 2, startRow: 5 })) {
    chunks.push(chunk);
  }

  assert.deepEqual(chunks, [
    '<row r="5"><c r="B5" t="inlineStr"><is><t>North</t></is></c><c r="C5"><v>10</v></c></row>',
    '<row r="6"><c r="B6" t="inlineStr"><is><t>South</t></is></c><c r="C6" t="b"><v>1</v></c></row>'
  ]);
});

test("streams row XML from sync data", async () => {
  const chunks: string[] = [];
  for await (const chunk of streamRowsXml([[{ formula: "=SUM(A1:A2)", result: 3 }]], {
    startRow: 3
  })) {
    chunks.push(chunk);
  }

  assert.deepEqual(chunks, ['<row r="3"><c r="A3"><f>SUM(A1:A2)</f><v>3</v></c></row>']);
});

test("streams existing worksheet rows from chunked worksheet XML", async () => {
  const rows = [];
  for await (const row of streamWorksheetRowsXml([
    '<worksheet><sheetData><row r="1"><c r="A1"><v>1',
    '</v></c></row><!-- <row r="99"/> --><row r="2"><c r="A2"/>',
    "</row></sheetData></worksheet>"
  ])) {
    rows.push(row);
  }

  assert.deepEqual(
    rows.map((row) => ({
      attributes: row.attributes,
      raw: row.raw,
      rowNumber: row.rowNumber
    })),
    [
      {
        attributes: { r: "1" },
        raw: '<row r="1"><c r="A1"><v>1</v></c></row>',
        rowNumber: 1
      },
      {
        attributes: { r: "2" },
        raw: '<row r="2"><c r="A2"/></row>',
        rowNumber: 2
      }
    ]
  );
});
