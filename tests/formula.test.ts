import assert from "node:assert/strict";
import test from "node:test";
import { parseFormulaSheetReferences } from "../packages/core/src/index.ts";

test("formula sheet reference parser handles quoted and unquoted sheet names", () => {
  assert.deepEqual(parseFormulaSheetReferences("SUM(Sheet1!A1,'Revenue Data'!B2)"), [
    { sheetName: "Sheet1" },
    { sheetName: "Revenue Data" }
  ]);
});

test("formula sheet reference parser ignores string literals and external workbook refs", () => {
  assert.deepEqual(
    parseFormulaSheetReferences('IF(A1="Missing!A1",[Book1.xlsx]Sheet1!A1,Sheet2!A1)'),
    [{ sheetName: "Sheet2" }]
  );
});

test("formula sheet reference parser unescapes quoted sheet names", () => {
  assert.deepEqual(parseFormulaSheetReferences("'Bob''s Sheet'!A1+'Bob''s Sheet'!A2"), [
    { sheetName: "Bob's Sheet" }
  ]);
});
