import assert from "node:assert/strict";
import test from "node:test";
import {
  formulaReferenceWithinExcelBounds,
  parseFormulaReferences,
  parseFormulaSheetReferences,
  parseFormulaStructuredReferences,
  renameFormulaStructuredReferenceTable
} from "../packages/core/src/index.ts";

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

test("formula reference parser returns local cells and ranges", () => {
  assert.deepEqual(parseFormulaReferences("SUM($A$1:B2,C3)"), [
    {
      kind: "range",
      ref: "A1:B2",
      start: { address: "A1", column: 1, row: 1 },
      end: { address: "B2", column: 2, row: 2 },
      range: {
        ref: "A1:B2",
        start: { address: "A1", column: 1, row: 1 },
        end: { address: "B2", column: 2, row: 2 }
      }
    },
    {
      kind: "cell",
      ref: "C3",
      start: { address: "C3", column: 3, row: 3 }
    }
  ]);
});

test("formula reference parser returns sheet-qualified references", () => {
  assert.deepEqual(parseFormulaReferences("'Revenue Data'!$B$2:Sheet1!C3"), [
    {
      kind: "cell",
      ref: "B2",
      sheetName: "Revenue Data",
      start: { address: "B2", column: 2, row: 2 }
    },
    {
      kind: "cell",
      ref: "C3",
      sheetName: "Sheet1",
      start: { address: "C3", column: 3, row: 3 }
    }
  ]);
});

test("formula reference parser ignores strings, external refs, and identifier-like text", () => {
  assert.deepEqual(
    parseFormulaReferences('IF(A1="B2",[Book1.xlsx]Sheet1!C3,LOG10(D4)+R2D2+Table1[Amount])'),
    [
      {
        kind: "cell",
        ref: "A1",
        start: { address: "A1", column: 1, row: 1 }
      },
      {
        kind: "cell",
        ref: "D4",
        start: { address: "D4", column: 4, row: 4 }
      }
    ]
  );
});

test("formula reference bounds checker uses Excel worksheet limits", () => {
  const [lastCell, outOfBoundsColumn, outOfBoundsRow] = parseFormulaReferences(
    "XFD1048576+XFE1+A1048577"
  );

  assert.equal(lastCell === undefined ? false : formulaReferenceWithinExcelBounds(lastCell), true);
  assert.equal(
    outOfBoundsColumn === undefined ? true : formulaReferenceWithinExcelBounds(outOfBoundsColumn),
    false
  );
  assert.equal(
    outOfBoundsRow === undefined ? true : formulaReferenceWithinExcelBounds(outOfBoundsRow),
    false
  );
});

test("formula structured reference parser extracts table references", () => {
  assert.deepEqual(
    parseFormulaStructuredReferences(
      'SUM(RevenueTable[Amount],RevenueTable[[#Totals],[Amount]],"Missing[Nope]")'
    ),
    [
      { tableName: "RevenueTable", raw: "RevenueTable[Amount]" },
      { tableName: "RevenueTable", raw: "RevenueTable[[#Totals],[Amount]]" }
    ]
  );
});

test("formula structured reference parser ignores bracketed external references", () => {
  assert.deepEqual(
    parseFormulaStructuredReferences("[Book1.xlsx]Sheet1!A1+Expense_Table[Amount]"),
    [{ tableName: "Expense_Table", raw: "Expense_Table[Amount]" }]
  );
});

test("formula structured reference rename retargets table tokens only", () => {
  assert.equal(
    renameFormulaStructuredReferenceTable(
      'SUM(RevenueTable[Amount],RevenueTable[[#Totals],[Amount]],OtherTable[Amount],"RevenueTable[Amount]")',
      ["RevenueTable"],
      "SalesData"
    ),
    'SUM(SalesData[Amount],SalesData[[#Totals],[Amount]],OtherTable[Amount],"RevenueTable[Amount]")'
  );
});
