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

test("worksheet append preserves the worksheet namespace prefix for generated rows", () => {
  const result = appendRows(prefixedWorksheet, [["Next", 2]]);

  assert.match(
    result.xml,
    /<x:row r="2"><x:c r="A2" t="inlineStr"><x:is><x:t>Next<\/x:t><\/x:is><\/x:c><x:c r="B2"><x:v>2<\/x:v><\/x:c><\/x:row>/
  );
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
