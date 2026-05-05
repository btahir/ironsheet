import assert from "node:assert/strict";
import test from "node:test";
import {
  appendRows,
  parseDefinedNames,
  parseSharedStrings,
  patchCell,
  readCell
} from "../packages/core/src/index.ts";

const prefixedWorksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:dimension ref="A1:A1"/>
  <x:sheetData>
    <x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Prefixed</x:t></x:is></x:c></x:row>
  </x:sheetData>
</x:worksheet>`;

test("worksheet reader handles namespace-prefixed cell XML", () => {
  assert.deepEqual(readCell(prefixedWorksheet, "A1"), {
    address: "A1",
    value: "Prefixed"
  });
});

test("worksheet patching preserves the worksheet namespace prefix for generated cells", () => {
  const result = patchCell(prefixedWorksheet, "A1", "Changed");

  assert.match(result.xml, /<x:c r="A1" t="inlineStr"><x:is><x:t>Changed<\/x:t><\/x:is><\/x:c>/);
});

test("worksheet patching preserves the worksheet namespace prefix for generated rows", () => {
  const result = patchCell(prefixedWorksheet, "B2", "Inserted");

  assert.match(
    result.xml,
    /<x:row r="2"><x:c r="B2" t="inlineStr"><x:is><x:t>Inserted<\/x:t><\/x:is><\/x:c><\/x:row>/
  );
});

test("worksheet patching inserts new rows in row order", () => {
  const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A3:A3"/>
  <sheetData>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Later</t></is></c></row>
  </sheetData>
</worksheet>`;

  const result = patchCell(worksheet, "A2", "Earlier");

  assert.match(
    result.xml,
    /<row r="2"><c r="A2" t="inlineStr"><is><t>Earlier<\/t><\/is><\/c><\/row>\s*<row r="3">/
  );
});

test("worksheet append preserves the worksheet namespace prefix for generated rows", () => {
  const result = appendRows(prefixedWorksheet, [["Next", 2]]);

  assert.match(
    result.xml,
    /<x:row r="2"><x:c r="A2" t="inlineStr"><x:is><x:t>Next<\/x:t><\/x:is><\/x:c><x:c r="B2"><x:v>2<\/x:v><\/x:c><\/x:row>/
  );
});

test("worksheet formulas preserve cached boolean and string result types", () => {
  const boolResult = patchCell(prefixedWorksheet, "B2", { formula: "=TRUE()", result: true });
  assert.match(boolResult.xml, /<x:c r="B2" t="b"><x:f>TRUE\(\)<\/x:f><x:v>1<\/x:v><\/x:c>/);
  assert.deepEqual(readCell(boolResult.xml, "B2"), {
    address: "B2",
    value: true,
    formula: "TRUE()"
  });

  const stringResult = patchCell(prefixedWorksheet, "C2", {
    formula: '=CONCAT("I","ron")',
    result: "Iron"
  });
  assert.match(
    stringResult.xml,
    /<x:c r="C2" t="str"><x:f>CONCAT\("I","ron"\)<\/x:f><x:v>Iron<\/x:v><\/x:c>/
  );
  assert.deepEqual(readCell(stringResult.xml, "C2"), {
    address: "C2",
    value: "Iron",
    formula: 'CONCAT("I","ron")'
  });
});

test("shared strings and defined names parse namespace-prefixed XML", () => {
  assert.deepEqual(
    parseSharedStrings(
      '<x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:si><x:r><x:t>Rich </x:t></x:r><x:r><x:t>Shared</x:t></x:r></x:si></x:sst>'
    ),
    ["Rich Shared"]
  );
  assert.deepEqual(
    parseDefinedNames(
      '<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:definedNames><x:definedName name="Revenue">Sheet1!$A$1:$B$2</x:definedName></x:definedNames></x:workbook>'
    ),
    [{ name: "Revenue", text: "Sheet1!$A$1:$B$2" }]
  );
});
