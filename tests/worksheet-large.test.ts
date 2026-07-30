import assert from "node:assert/strict";
import test from "node:test";
import { replaceRowsInRange } from "../packages/core/src/worksheet.ts";

test(
  "large table replacement recalculates dimensions without argument overflows",
  { timeout: 10_000 },
  () => {
    const rowCount = 70_000;
    const rows = Array.from({ length: rowCount }, (_, index) => [`Row ${index + 1}`, index + 1]);
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Old</t></is></c><c r="B2"><v>1</v></c></row>
  </sheetData>
</worksheet>`;

    const updated = replaceRowsInRange(
      xml,
      { startRow: 2, endRow: 2, startColumn: 1, endColumn: 2 },
      rows
    );

    assert.match(updated, /<dimension ref="A1:B70001"\/>/);
    assert.match(updated, /<row r="70001">/);
  }
);
