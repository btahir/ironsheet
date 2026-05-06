import assert from "node:assert/strict";
import test from "node:test";
import { parseZip, readEntryData, Workbook } from "../packages/core/src/index.ts";
import { openPackage, openWorkbook, nodeCompressionAdapter } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

const textDecoder = new TextDecoder();

test("inspects workbook sheets from OPC relationships", async () => {
  const data = await createMinimalWorkbook();
  const workbook = await openWorkbook(data);

  assert.deepEqual(workbook.sheets(), [
    {
      name: "Sheet1",
      id: "1",
      relationshipId: "rId1",
      partName: "xl/worksheets/sheet1.xml"
    }
  ]);

  assert.deepEqual((await workbook.inspect()).features, {
    calcChains: 0,
    charts: 0,
    comments: 0,
    conditionalFormats: 0,
    dataValidations: 0,
    definedNames: 0,
    drawings: 0,
    externalRelationships: 0,
    formulaCells: 0,
    hiddenSheets: 0,
    hyperlinks: 0,
    macros: 0,
    media: 0,
    merges: 0,
    pivotTables: 0,
    sharedStrings: 0,
    tables: 0
  });
});

test("inspect reports workbook feature signals", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({
      includeCalcChain: true,
      includeComment: true,
      includeConditionalFormatting: true,
      includeDataValidation: true,
      includeDefinedName: true,
      includeDrawing: true,
      includeHiddenSheet: true,
      includeHyperlink: true,
      includeMacro: true,
      includeMerge: true,
      includeTable: true,
      useSharedStrings: true
    })
  );

  assert.deepEqual(workbook.sheets(), [
    {
      name: "Sheet1",
      id: "1",
      relationshipId: "rId1",
      partName: "xl/worksheets/sheet1.xml"
    },
    {
      name: "HiddenData",
      id: "2",
      relationshipId: "rIdHidden",
      partName: "xl/worksheets/sheet2.xml",
      state: "hidden"
    }
  ]);

  assert.deepEqual((await workbook.inspect()).features, {
    calcChains: 1,
    charts: 1,
    comments: 1,
    conditionalFormats: 1,
    dataValidations: 1,
    definedNames: 2,
    drawings: 1,
    externalRelationships: 1,
    formulaCells: 0,
    hiddenSheets: 1,
    hyperlinks: 1,
    macros: 1,
    media: 1,
    merges: 1,
    pivotTables: 0,
    sharedStrings: 1,
    tables: 1
  });
});

test("inspect counts worksheet formula cells", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeTable: true, includeTableTotals: true })
  );

  assert.equal((await workbook.inspect()).features.formulaCells, 1);
});

test("lists workbook hyperlinks with external targets", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeHyperlink: true }));

  assert.deepEqual(await workbook.hyperlinks(), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      ref: "A1",
      relationshipId: "rIdHyperlink1",
      target: "https://example.com",
      targetMode: "External"
    }
  ]);
});

test("lists workbook comments with authors", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeComment: true }));

  assert.deepEqual(await workbook.comments(), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      commentPartName: "xl/comments1.xml",
      relationshipId: "rIdComments1",
      ref: "A1",
      text: "Check this value",
      author: "Ironsheet",
      authorId: "0",
      rawXml: '<comment ref="A1" authorId="0"><text><r><t>Check this value</t></r></text></comment>'
    }
  ]);
});

test("lists workbook auto filters", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeAutoFilter: true }));

  assert.deepEqual(await workbook.autoFilters(), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      ref: "A1:B10",
      rawXml:
        '<autoFilter ref="A1:B10"><filterColumn colId="0"><filters><filter val="Original"/></filters></filterColumn></autoFilter>'
    }
  ]);
});

test("lists workbook images and replaces existing image bytes", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeDrawing: true }));
  const image = {
    sheetName: "Sheet1",
    sheetPartName: "xl/worksheets/sheet1.xml",
    drawingPartName: "xl/drawings/drawing1.xml",
    drawingRelationshipId: "rIdDrawing1",
    imageRelationshipId: "rIdImage1",
    target: "../media/image1.png",
    imagePartName: "xl/media/image1.png"
  };

  assert.deepEqual(await workbook.images(), [image]);

  const replacement = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
  assert.deepEqual(await workbook.replaceImage("xl/media/image1.png", replacement), image);
  assert.deepEqual(Array.from(await workbook.pkg.readPart("xl/media/image1.png")), [
    ...replacement
  ]);

  const reopened = await openWorkbook(await workbook.write());
  assert.deepEqual(Array.from(await reopened.pkg.readPart("xl/media/image1.png")), [
    ...replacement
  ]);
  assert.deepEqual(await reopened.images(), [image]);
  assert.deepEqual((await reopened.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("lists workbook merged cells", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeMerge: true }));

  assert.deepEqual(await workbook.mergedCells(), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      ref: "A1:B1"
    }
  ]);
});

test("lists workbook data validations", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeDataValidation: true }));

  assert.deepEqual(await workbook.dataValidations(), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      sqref: "B2:B10",
      allowBlank: true,
      formula1: "0",
      formula2: "100",
      operator: "between",
      type: "whole"
    }
  ]);
});

test("lists workbook conditional formats", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeConditionalFormatting: true })
  );
  const conditionalFormats = await workbook.conditionalFormats();

  assert.equal(conditionalFormats.length, 1);
  assert.equal(conditionalFormats[0]?.sheetName, "Sheet1");
  assert.equal(conditionalFormats[0]?.sheetPartName, "xl/worksheets/sheet1.xml");
  assert.equal(conditionalFormats[0]?.sqref, "A1:A10");
  assert.deepEqual(conditionalFormats[0]?.rules[0], {
    attributes: {
      operator: "greaterThan",
      priority: "1",
      type: "cellIs"
    },
    formulas: ["10"],
    operator: "greaterThan",
    priority: "1",
    rawXml:
      '<cfRule type="cellIs" priority="1" operator="greaterThan"><formula>10</formula></cfRule>',
    type: "cellIs"
  });
});

test("lists workbook table metadata", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeSecondTable: true, includeTable: true })
  );

  assert.deepEqual(await workbook.tables(), [
    {
      name: "RevenueTable",
      displayName: "RevenueTable",
      columns: [
        { id: "1", name: "Name" },
        { id: "2", name: "Amount" }
      ],
      partName: "xl/tables/table1.xml",
      worksheetPartName: "xl/worksheets/sheet1.xml",
      ref: "A1:B2",
      totalsRowCount: 0
    },
    {
      name: "ExpenseTable",
      displayName: "ExpenseTable",
      columns: [
        { id: "1", name: "Category" },
        { id: "2", name: "Amount" }
      ],
      partName: "xl/tables/table2.xml",
      worksheetPartName: "xl/worksheets/sheet1.xml",
      ref: "C1:D2",
      totalsRowCount: 0
    }
  ]);
});

test("renames tables and retargets structured references", async () => {
  const pkg = await openPackage(
    await createMinimalWorkbook({
      includeDefinedName: true,
      includeDrawing: true,
      includeTable: true
    })
  );
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      "</sheetData>",
      '<row r="3"><c r="C3"><f>SUM(RevenueTable[Amount])</f><v>1</v></c></row></sheetData>'
    )
  );
  const workbookXml = await pkg.readText("xl/workbook.xml");
  pkg.setText(
    "xl/workbook.xml",
    workbookXml.replace(
      "</definedNames>",
      '<definedName name="RevenueFormula">RevenueTable[Amount]</definedName></definedNames>'
    )
  );
  const chartXml = await pkg.readText("xl/charts/chart1.xml");
  pkg.setText(
    "xl/charts/chart1.xml",
    chartXml.replace(
      "<c:chart>",
      "<c:chart><c:plotArea><c:ser><c:val><c:numRef><c:f>RevenueTable[Amount]</c:f></c:numRef></c:val></c:ser></c:plotArea>"
    )
  );
  const workbook = await Workbook.fromPackage(pkg);

  const renamed = await workbook.renameTable("RevenueTable", "SalesData");

  assert.equal(renamed.name, "SalesData");
  assert.equal(renamed.displayName, "SalesData");
  assert.match(
    await pkg.readText("xl/tables/table1.xml"),
    /name="SalesData" displayName="SalesData"/
  );
  assert.equal(
    (await workbook.formulas()).some((formula) => formula.formula === "SUM(SalesData[Amount])"),
    true
  );
  assert.deepEqual(
    (await workbook.definedNames()).find((definedName) => definedName.name === "RevenueFormula")
      ?.text,
    "SalesData[Amount]"
  );
  assert.equal((await pkg.readText("xl/charts/chart1.xml")).includes("SalesData[Amount]"), true);
  assert.match(await pkg.readText("xl/workbook.xml"), /forceFullCalc="1"/);
});

test("renames sheets and retargets workbook references", async () => {
  const pkg = await openPackage(
    await createMinimalWorkbook({
      includeDefinedName: true,
      includeDrawing: true,
      includePivotTable: true,
      includeTable: true
    })
  );
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace('<c r="B2"><v>1</v></c>', '<c r="B2"><f>SUM(Sheet1!$B$2)</f><v>1</v></c>')
  );
  const workbook = await Workbook.fromPackage(pkg);

  const renamed = await workbook.renameSheet("Sheet1", "Revenue 2026");

  assert.equal(renamed.name, "Revenue 2026");
  assert.equal(workbook.sheet("Revenue 2026").partName, "xl/worksheets/sheet1.xml");
  assert.throws(() => workbook.sheet("Sheet1"), /Unknown worksheet/);
  assert.match(await pkg.readText("xl/workbook.xml"), /name="Revenue 2026"/);
  assert.deepEqual(
    (await workbook.definedNames()).find((definedName) => definedName.name === "RevenueRange")
      ?.text,
    "'Revenue 2026'!$A$1:$B$2"
  );
  assert.match(await pkg.readText("xl/worksheets/sheet1.xml"), /SUM\('Revenue 2026'!\$B\$2\)/);
  assert.match(await pkg.readText("xl/charts/chart1.xml"), /'Revenue 2026'!\$A\$2:\$A\$3/);
  assert.match(
    await pkg.readText("xl/pivotCache/pivotCacheDefinition1.xml"),
    /sheet="Revenue 2026"/
  );
  assert.match(await pkg.readText("xl/workbook.xml"), /forceFullCalc="1"/);
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("renaming sheets rejects duplicates and invalid names", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeHiddenSheet: true }));

  await assert.rejects(() => workbook.renameSheet("Sheet1", "HiddenData"), /already used/);
  await assert.rejects(() => workbook.renameSheet("Sheet1", "Bad/Name"), /invalid Excel/);
  await assert.rejects(() => workbook.renameSheet("Sheet1", "A".repeat(32)), /31 character/);
});

test("hides and shows worksheets without hiding the whole workbook", async () => {
  const singleSheetWorkbook = await openWorkbook(await createMinimalWorkbook());
  await assert.rejects(
    () => singleSheetWorkbook.hideSheet("Sheet1"),
    /at least one visible worksheet/
  );

  const workbook = await openWorkbook(await createMinimalWorkbook({ includeHiddenSheet: true }));

  assert.deepEqual(await workbook.showSheet("HiddenData"), {
    name: "HiddenData",
    id: "2",
    relationshipId: "rIdHidden",
    partName: "xl/worksheets/sheet2.xml"
  });
  assert.deepEqual(await workbook.hideSheet("Sheet1", "veryHidden"), {
    name: "Sheet1",
    id: "1",
    relationshipId: "rId1",
    partName: "xl/worksheets/sheet1.xml",
    state: "veryHidden"
  });
  assert.deepEqual(workbook.sheets(), [
    {
      name: "Sheet1",
      id: "1",
      relationshipId: "rId1",
      partName: "xl/worksheets/sheet1.xml",
      state: "veryHidden"
    },
    {
      name: "HiddenData",
      id: "2",
      relationshipId: "rIdHidden",
      partName: "xl/worksheets/sheet2.xml"
    }
  ]);

  const workbookXml = await workbook.pkg.readText("xl/workbook.xml");
  assert.match(workbookXml, /<sheet name="Sheet1" sheetId="1" r:id="rId1" state="veryHidden"\/>/);
  assert.match(workbookXml, /<sheet name="HiddenData" sheetId="2" r:id="rIdHidden"\/>/);
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("renaming tables rejects duplicate and invalid names", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeSecondTable: true, includeTable: true })
  );

  await assert.rejects(() => workbook.renameTable("RevenueTable", "ExpenseTable"), /already used/);
  await assert.rejects(() => workbook.renameTable("RevenueTable", "A1"), /cell reference/);
});

test("renames table columns and retargets structured references", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeTable: true }));
  const tableXml = await pkg.readText("xl/tables/table1.xml");
  pkg.setText(
    "xl/tables/table1.xml",
    tableXml.replace(
      '<tableColumn id="2" name="Amount"/>',
      '<tableColumn id="2" name="Amount"><calculatedColumnFormula>[@Amount]*2</calculatedColumnFormula></tableColumn>'
    )
  );
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      "</sheetData>",
      '<row r="3"><c r="C3"><f>SUM(RevenueTable[Amount])</f><v>1</v></c></row></sheetData>'
    )
  );
  const workbook = await Workbook.fromPackage(pkg);

  const renamed = await workbook.renameTableColumn("RevenueTable", "Amount", "NetAmount");

  assert.deepEqual(renamed.columns, [
    { id: "1", name: "Name" },
    { id: "2", name: "NetAmount" }
  ]);
  assert.deepEqual(await workbook.readCell("Sheet1", "B1"), {
    address: "B1",
    value: "NetAmount"
  });
  assert.match(await pkg.readText("xl/tables/table1.xml"), /name="NetAmount"/);
  assert.match(await pkg.readText("xl/tables/table1.xml"), /\[@NetAmount\]\*2/);
  assert.equal(
    (await workbook.formulas()).some(
      (formula) => formula.formula === "SUM(RevenueTable[NetAmount])"
    ),
    true
  );
  assert.match(await pkg.readText("xl/workbook.xml"), /forceFullCalc="1"/);
});

test("renaming table columns rejects duplicate and empty names", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeTable: true }));

  await assert.rejects(
    () => workbook.renameTableColumn("RevenueTable", "Amount", "Name"),
    /already used/
  );
  await assert.rejects(
    () => workbook.renameTableColumn("RevenueTable", "Amount", ""),
    /cannot be empty/
  );
});

test("appends and removes rightmost table columns", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeTable: true }));

  const appended = await workbook.appendTableColumn("RevenueTable", "Margin", [0.5]);

  assert.equal(appended.ref, "A1:C2");
  assert.deepEqual(appended.columns, [
    { id: "1", name: "Name" },
    { id: "2", name: "Amount" },
    { id: "3", name: "Margin" }
  ]);
  assert.deepEqual(await workbook.readRange("Sheet1", "A1:C2"), {
    range: "A1:C2",
    cells: [
      [
        { address: "A1", value: "Name" },
        { address: "B1", value: "Amount" },
        { address: "C1", value: "Margin" }
      ],
      [
        { address: "A2", value: "Old" },
        { address: "B2", value: 1 },
        { address: "C2", value: 0.5 }
      ]
    ]
  });
  assert.match(await workbook.pkg.readText("xl/tables/table1.xml"), /ref="A1:C2"/);
  assert.match(await workbook.pkg.readText("xl/tables/table1.xml"), /<tableColumns count="3">/);

  const removed = await workbook.removeRightmostTableColumn("RevenueTable", "Margin");

  assert.equal(removed.ref, "A1:B2");
  assert.deepEqual(removed.columns, [
    { id: "1", name: "Name" },
    { id: "2", name: "Amount" }
  ]);
  assert.deepEqual(await workbook.readCell("Sheet1", "C1"), undefined);
  assert.deepEqual(await workbook.readCell("Sheet1", "C2"), undefined);
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("removing table columns only allows the rightmost column", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeTable: true }));

  await assert.rejects(
    () => workbook.removeRightmostTableColumn("RevenueTable", "Name"),
    /rightmost/
  );
});

test("retargets chart formulas exactly", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeDrawing: true }));
  const chartXml = await pkg.readText("xl/charts/chart1.xml");
  pkg.setText(
    "xl/charts/chart1.xml",
    chartXml.replace(
      "<c:chart>",
      "<c:chart><c:plotArea><c:ser><c:val><c:numRef><c:f>Sheet1!$A$1:$B$2</c:f></c:numRef></c:val></c:ser></c:plotArea>"
    )
  );
  const workbook = await Workbook.fromPackage(pkg);

  assert.equal(
    await workbook.retargetChartFormulas([{ from: "Sheet1!$A$1:$B$2", to: "Sheet1!$A$1:$C$2" }]),
    1
  );
  assert.match(await pkg.readText("xl/charts/chart1.xml"), /Sheet1!\$A\$1:\$C\$2/);
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("retargets pivot cache worksheet sources", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includePivotTable: true }));
  const workbook = await Workbook.fromPackage(pkg);

  assert.equal(
    await workbook.retargetPivotCacheSources([
      {
        from: { sheet: "Sheet1", ref: "A1:B2" },
        to: { sheet: "Sheet1", ref: "A1:C2" }
      }
    ]),
    1
  );
  assert.match(
    await pkg.readText("xl/pivotCache/pivotCacheDefinition1.xml"),
    /worksheetSource ref="A1:C2" sheet="Sheet1"/
  );
  assert.equal((await workbook.validate()).summary.errors, 0);
});

test("reads workbook formula inventory with parsed dependencies", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeTable: true }));
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<c r="A2" t="inlineStr"><is><t>Old</t></is></c>',
      '<c r="A2"><f>SUM(A1:B1,RevenueTable[Amount])</f><v>1</v></c>'
    )
  );
  const workbook = await Workbook.fromPackage(pkg);
  const [formula] = await workbook.formulas();

  assert.equal(formula?.sheetName, "Sheet1");
  assert.equal(formula?.sheetPartName, "xl/worksheets/sheet1.xml");
  assert.equal(formula?.address, "A2");
  assert.equal(formula?.formula, "SUM(A1:B1,RevenueTable[Amount])");
  assert.deepEqual(
    formula?.references.map((reference) => reference.ref),
    ["A1:B1"]
  );
  assert.deepEqual(formula?.structuredReferences, [
    { tableName: "RevenueTable", raw: "RevenueTable[Amount]" }
  ]);
});

test("reads shared formula metadata from workbook formula inventory", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c></row>',
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" si="0" ref="B1:B2">A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>'
    )
  );
  const workbook = await Workbook.fromPackage(pkg);
  const formulas = await workbook.formulas();

  assert.deepEqual(
    formulas.map((formula) => ({
      address: formula.address,
      formula: formula.formula,
      formulaRef: formula.formulaRef,
      formulaType: formula.formulaType,
      sharedIndex: formula.sharedIndex
    })),
    [
      {
        address: "B1",
        formula: "A1*2",
        formulaRef: "B1:B2",
        formulaType: "shared",
        sharedIndex: "0"
      },
      {
        address: "B2",
        formula: "",
        formulaRef: undefined,
        formulaType: "shared",
        sharedIndex: "0"
      }
    ]
  );
});

test("reads workbook defined names", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeDefinedName: true }));

  assert.deepEqual(await workbook.definedNames(), [
    {
      name: "RevenueRange",
      text: "Sheet1!$A$1:$B$2",
      comment: "Template output range"
    },
    {
      name: "_xlnm.Print_Titles",
      text: "Sheet1!$1:$1",
      hidden: true,
      localSheetId: "0"
    }
  ]);
});

test("sets, replaces, and deletes workbook defined names", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await workbook.setDefinedName("ReportRange", "Sheet1!$A$1:$B$10", {
    comment: "Export range",
    hidden: true
  });
  await workbook.setDefinedName("ScopedRange", "Sheet1!$A$1", { sheetName: "Sheet1" });
  await workbook.setDefinedName("ReportRange", "Sheet1!$A$1:$C$10");

  assert.deepEqual(await workbook.definedNames(), [
    {
      name: "ReportRange",
      text: "Sheet1!$A$1:$C$10"
    },
    {
      name: "ScopedRange",
      text: "Sheet1!$A$1",
      localSheetId: "0"
    }
  ]);
  assert.equal(await workbook.deleteDefinedName("ScopedRange", { sheetName: "Sheet1" }), true);
  assert.equal(await workbook.deleteDefinedName("MissingName"), false);
  assert.deepEqual(await workbook.definedNames(), [
    {
      name: "ReportRange",
      text: "Sheet1!$A$1:$C$10"
    }
  ]);
  assert.match(await workbook.pkg.readText("xl/workbook.xml"), /forceFullCalc="1"/);
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("defined name mutation rejects invalid names and unknown sheet scopes", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await assert.rejects(() => workbook.setDefinedName("A1", "Sheet1!$A$1"), /cell reference/);
  await assert.rejects(
    () => workbook.setDefinedName("ScopedRange", "Sheet1!$A$1", { sheetName: "Missing" }),
    /Unknown worksheet/
  );
});

test("sets, replaces, and deletes auto filters", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeConditionalFormatting: true })
  );

  assert.deepEqual(await workbook.setAutoFilter("Sheet1", { ref: "$A$1:$C$10" }), {
    sheetName: "Sheet1",
    sheetPartName: "xl/worksheets/sheet1.xml",
    ref: "A1:C10"
  });

  let sheetXml = await workbook.pkg.readText("xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<autoFilter ref="A1:C10"\/>/);
  assert.ok(sheetXml.indexOf("<autoFilter") < sheetXml.indexOf("<conditionalFormatting"));

  await workbook.setAutoFilter("Sheet1", {
    ref: "A1:C20",
    rawXml:
      '<autoFilter ref="A1:C10"><filterColumn colId="0"><filters><filter val="Original"/></filters></filterColumn></autoFilter>'
  });
  assert.deepEqual(await workbook.autoFilters("Sheet1"), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      ref: "A1:C20",
      rawXml:
        '<autoFilter ref="A1:C20"><filterColumn colId="0"><filters><filter val="Original"/></filters></filterColumn></autoFilter>'
    }
  ]);

  assert.equal(await workbook.deleteAutoFilter("Sheet1"), true);
  assert.equal(await workbook.deleteAutoFilter("Sheet1"), false);
  sheetXml = await workbook.pkg.readText("xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<autoFilter/);
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("sets, replaces, and deletes conditional formats", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeDataValidation: true }));

  assert.deepEqual(
    await workbook.setConditionalFormat("Sheet1", {
      sqref: "$C$2:$C$10",
      rules: [
        {
          formulas: ["100"],
          operator: "greaterThan",
          priority: "1",
          stopIfTrue: true,
          type: "cellIs"
        }
      ]
    }),
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      sqref: "C2:C10",
      rules: [
        {
          formulas: ["100"],
          operator: "greaterThan",
          priority: "1",
          stopIfTrue: true,
          type: "cellIs"
        }
      ]
    }
  );

  let sheetXml = await workbook.pkg.readText("xl/worksheets/sheet1.xml");
  assert.match(
    sheetXml,
    /<conditionalFormatting sqref="C2:C10"><cfRule type="cellIs" priority="1" operator="greaterThan" stopIfTrue="1"><formula>100<\/formula><\/cfRule><\/conditionalFormatting>/
  );
  assert.ok(sheetXml.indexOf("<conditionalFormatting") < sheetXml.indexOf("<dataValidations"));

  await workbook.setConditionalFormat("Sheet1", {
    sqref: "C2:C10",
    rules: [
      {
        rawXml: '<cfRule type="top10" priority="2" rank="5"/>'
      }
    ]
  });
  assert.deepEqual(await workbook.conditionalFormats("Sheet1"), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      sqref: "C2:C10",
      rules: [
        {
          attributes: {
            priority: "2",
            rank: "5",
            type: "top10"
          },
          priority: "2",
          rank: "5",
          rawXml: '<cfRule type="top10" priority="2" rank="5"/>',
          type: "top10"
        }
      ]
    }
  ]);

  assert.equal(await workbook.deleteConditionalFormat("Sheet1", "$C$2:$C$10"), true);
  assert.equal(await workbook.deleteConditionalFormat("Sheet1", "C2:C10"), false);
  sheetXml = await workbook.pkg.readText("xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<conditionalFormatting/);
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("sets, replaces, and deletes data validations", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  assert.deepEqual(
    await workbook.setDataValidation("Sheet1", {
      sqref: "$B$2:$B$10",
      allowBlank: true,
      error: "Use a whole number from 0 to 100.",
      errorTitle: "Invalid amount",
      formula1: "0",
      formula2: "100",
      operator: "between",
      prompt: "Enter an amount",
      promptTitle: "Amount",
      showErrorMessage: true,
      showInputMessage: true,
      type: "whole"
    }),
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      sqref: "B2:B10",
      allowBlank: true,
      error: "Use a whole number from 0 to 100.",
      errorTitle: "Invalid amount",
      formula1: "0",
      formula2: "100",
      operator: "between",
      prompt: "Enter an amount",
      promptTitle: "Amount",
      showErrorMessage: true,
      showInputMessage: true,
      type: "whole"
    }
  );

  assert.match(
    await workbook.pkg.readText("xl/worksheets/sheet1.xml"),
    /<dataValidations count="1"><dataValidation sqref="B2:B10" type="whole" operator="between" allowBlank="1" showErrorMessage="1" showInputMessage="1" errorTitle="Invalid amount" error="Use a whole number from 0 to 100\." promptTitle="Amount" prompt="Enter an amount"><formula1>0<\/formula1><formula2>100<\/formula2><\/dataValidation><\/dataValidations>/
  );

  await workbook.setDataValidation("Sheet1", {
    sqref: "B2:B10",
    formula1: '"North,South,West"',
    type: "list"
  });
  assert.deepEqual(await workbook.dataValidations("Sheet1"), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      sqref: "B2:B10",
      formula1: '"North,South,West"',
      type: "list"
    }
  ]);

  assert.equal(await workbook.deleteDataValidation("Sheet1", "$B$2:$B$10"), true);
  assert.equal(await workbook.deleteDataValidation("Sheet1", "B2:B10"), false);
  assert.deepEqual(await workbook.dataValidations("Sheet1"), []);
  assert.doesNotMatch(await workbook.pkg.readText("xl/worksheets/sheet1.xml"), /dataValidations/);
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("sets, replaces, and deletes external hyperlinks", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  assert.deepEqual(
    await workbook.setHyperlink("Sheet1", "b2", "https://docs.example.com", {
      display: "Docs",
      tooltip: "Open docs"
    }),
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      ref: "B2",
      relationshipId: "rId1",
      target: "https://docs.example.com",
      targetMode: "External",
      display: "Docs",
      tooltip: "Open docs"
    }
  );

  assert.deepEqual(await workbook.hyperlinks("Sheet1"), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      ref: "B2",
      relationshipId: "rId1",
      target: "https://docs.example.com",
      targetMode: "External",
      display: "Docs",
      tooltip: "Open docs"
    }
  ]);

  assert.match(
    await workbook.pkg.readText("xl/worksheets/sheet1.xml"),
    /xmlns:r="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships"/
  );
  assert.match(
    await workbook.pkg.readText("xl/worksheets/sheet1.xml"),
    /<hyperlinks><hyperlink ref="B2" r:id="rId1" display="Docs" tooltip="Open docs"\/><\/hyperlinks>/
  );
  assert.match(
    await workbook.pkg.readText("xl/worksheets/_rels/sheet1.xml.rels"),
    /<Relationship Id="rId1" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink" Target="https:\/\/docs\.example\.com" TargetMode="External"\/>/
  );

  await workbook.setHyperlink("Sheet1", "B2", "https://docs.example.com/v2", {
    display: "Docs v2"
  });
  assert.deepEqual(await workbook.hyperlinks("Sheet1"), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      ref: "B2",
      relationshipId: "rId1",
      target: "https://docs.example.com/v2",
      targetMode: "External",
      display: "Docs v2"
    }
  ]);

  assert.equal(await workbook.deleteHyperlink("Sheet1", "B2"), true);
  assert.equal(await workbook.deleteHyperlink("Sheet1", "B2"), false);
  assert.deepEqual(await workbook.hyperlinks("Sheet1"), []);
  assert.doesNotMatch(await workbook.pkg.readText("xl/worksheets/sheet1.xml"), /<hyperlinks>/);
  assert.doesNotMatch(
    await workbook.pkg.readText("xl/worksheets/_rels/sheet1.xml.rels"),
    /hyperlink/
  );
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("merges, dedupes, rejects overlaps, and unmerges cells", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  assert.deepEqual(await workbook.mergeCells("Sheet1", "$B$2:$C$3"), {
    sheetName: "Sheet1",
    sheetPartName: "xl/worksheets/sheet1.xml",
    ref: "B2:C3"
  });
  assert.deepEqual(await workbook.mergeCells("Sheet1", "B2:C3"), {
    sheetName: "Sheet1",
    sheetPartName: "xl/worksheets/sheet1.xml",
    ref: "B2:C3"
  });
  assert.deepEqual(await workbook.mergedCells("Sheet1"), [
    {
      sheetName: "Sheet1",
      sheetPartName: "xl/worksheets/sheet1.xml",
      ref: "B2:C3"
    }
  ]);
  assert.match(
    await workbook.pkg.readText("xl/worksheets/sheet1.xml"),
    /<mergeCells count="1"><mergeCell ref="B2:C3"\/><\/mergeCells>/
  );

  await assert.rejects(() => workbook.mergeCells("Sheet1", "C3:D4"), /overlaps existing merge/);
  await assert.rejects(() => workbook.mergeCells("Sheet1", "D4"), /single cell/);

  assert.equal(await workbook.unmergeCells("Sheet1", "B2:C3"), true);
  assert.equal(await workbook.unmergeCells("Sheet1", "B2:C3"), false);
  assert.deepEqual(await workbook.mergedCells("Sheet1"), []);
  assert.doesNotMatch(await workbook.pkg.readText("xl/worksheets/sheet1.xml"), /<mergeCells/);
  assert.deepEqual((await workbook.validate()).summary, { errors: 0, warnings: 0, infos: 0 });
});

test("reads workbook style metadata", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  assert.deepEqual(await workbook.styles(), {
    cellStyleXfs: [{ numFmtId: "0", fontId: "0", fillId: "0", borderId: "0" }],
    cellXfs: [
      { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", xfId: "0" },
      { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", xfId: "0", applyFont: "1" },
      { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", xfId: "0", applyFill: "1" },
      { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", xfId: "0", applyBorder: "1" },
      {
        numFmtId: "0",
        fontId: "0",
        fillId: "0",
        borderId: "0",
        xfId: "0",
        applyNumberFormat: "1"
      }
    ],
    counts: {
      borders: 1,
      cellStyleXfs: 1,
      cellXfs: 5,
      fills: 1,
      fonts: 1,
      numFmts: 0
    },
    numberFormats: []
  });
});

test("styles cells with deduped custom number formats", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  const firstStyleId = await workbook.styleCell("Sheet1", "A1", {
    numberFormat: "$#,##0.00"
  });
  const secondStyleId = await workbook.styleCell("Sheet1", "A1", {
    numberFormat: "$#,##0.00"
  });

  assert.equal(firstStyleId, secondStyleId);
  assert.deepEqual(await workbook.readCell("Sheet1", "A1"), {
    address: "A1",
    value: "Original",
    styleId: firstStyleId
  });
  assert.deepEqual(await workbook.styles(), {
    cellStyleXfs: [{ numFmtId: "0", fontId: "0", fillId: "0", borderId: "0" }],
    cellXfs: [
      { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", xfId: "0" },
      { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", xfId: "0", applyFont: "1" },
      { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", xfId: "0", applyFill: "1" },
      { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", xfId: "0", applyBorder: "1" },
      {
        numFmtId: "0",
        fontId: "0",
        fillId: "0",
        borderId: "0",
        xfId: "0",
        applyNumberFormat: "1"
      },
      {
        numFmtId: "164",
        fontId: "0",
        fillId: "0",
        borderId: "0",
        xfId: "0",
        applyFont: "1",
        applyNumberFormat: "1"
      }
    ],
    counts: {
      borders: 1,
      cellStyleXfs: 1,
      cellXfs: 6,
      fills: 1,
      fonts: 1,
      numFmts: 1
    },
    numberFormats: [{ numFmtId: "164", formatCode: "$#,##0.00" }]
  });
});

test("reads existing cell values and style ids", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  assert.deepEqual(await workbook.readCell("Sheet1", "A1"), {
    address: "A1",
    value: "Original",
    styleId: "1"
  });
  assert.equal(await workbook.readCell("Sheet1", "Z99"), undefined);
});

test("reads shared string cell values", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ useSharedStrings: true }));

  assert.deepEqual(await workbook.readCell("Sheet1", "A1"), {
    address: "A1",
    value: "Original",
    styleId: "1"
  });
});

test("reads rich inline string cells as combined text", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ useRichInlineString: true }));

  assert.deepEqual(await workbook.readCell("Sheet1", "A1"), {
    address: "A1",
    value: "Rich Text",
    styleId: "1"
  });
});

test("patches and reads ranges", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await workbook.patchRange("Sheet1", "B2", [
    ["Move fast", 12],
    [true, { formula: "=C2*2", result: 24 }]
  ]);

  assert.deepEqual(await workbook.readRange("Sheet1", "B2:C3"), {
    range: "B2:C3",
    cells: [
      [
        { address: "B2", value: "Move fast" },
        { address: "C2", value: 12 }
      ],
      [
        { address: "B3", value: true },
        { address: "C3", value: 24, formula: "C2*2" }
      ]
    ]
  });
});

test("appends rows after the current used range", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeCalcChain: true }));

  await workbook.appendRows("Sheet1", [
    ["North", 10],
    ["South", { formula: "=B2*2", result: 20 }]
  ]);

  assert.deepEqual(await workbook.readRange("Sheet1", "A2:B3"), {
    range: "A2:B3",
    cells: [
      [
        { address: "A2", value: "North" },
        { address: "B2", value: 10 }
      ],
      [
        { address: "A3", value: "South" },
        { address: "B3", value: 20, formula: "B2*2" }
      ]
    ]
  });

  const outputZip = parseZip(await workbook.write());
  const names = outputZip.entries.map((entry) => entry.name);
  assert.equal(names.includes("xl/calcChain.xml"), false);

  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(sheetXml, /<dimension ref="A1:B3"\/>/);
});

test("append rows inserts a contiguous row block without rewriting existing rows", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await workbook.appendRows("Sheet1", [
    ["A", 1],
    ["B", 2],
    ["C", 3]
  ]);

  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(
    sheetXml,
    /<row r="2"><c r="A2" t="inlineStr"><is><t>A<\/t><\/is><\/c><c r="B2"><v>1<\/v><\/c><\/row><row r="3"><c r="A3" t="inlineStr"><is><t>B<\/t><\/is><\/c><c r="B3"><v>2<\/v><\/c><\/row><row r="4"><c r="A4" t="inlineStr"><is><t>C<\/t><\/is><\/c><c r="B4"><v>3<\/v><\/c><\/row>/
  );
  assert.match(sheetXml, /<dimension ref="A1:B4"\/>/);
});

test("patches one cell and preserves untouched entry payloads", async () => {
  const original = await createMinimalWorkbook();
  const originalZip = parseZip(original);
  const originalStyles = originalZip.entries.find((entry) => entry.name === "xl/styles.xml");
  assert.ok(originalStyles);

  const workbook = await openWorkbook(original);
  await workbook.patchCell("Sheet1", "B2", "Move fast");
  const output = await workbook.write();
  const outputZip = parseZip(output);

  const outputStyles = outputZip.entries.find((entry) => entry.name === "xl/styles.xml");
  assert.ok(outputStyles);
  assert.deepEqual(outputStyles.compressedData, originalStyles.compressedData);

  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));

  assert.match(sheetXml, /<c r="B2" t="inlineStr"><is><t>Move fast<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<dimension ref="A1:B2"\/>/);
});

test("patching an existing styled cell preserves the style id", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await workbook.patchCell("Sheet1", "A1", "Styled replacement");
  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));

  assert.match(sheetXml, /<c r="A1" s="1" t="inlineStr"><is><t>Styled replacement<\/t><\/is><\/c>/);
});

test("patches formulas and marks workbook for recalculation", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook());

  await workbook.patchCell("Sheet1", "C3", { formula: "=SUM(A1:B2)" });
  const outputZip = parseZip(await workbook.write());

  const workbookEntry = outputZip.entries.find((entry) => entry.name === "xl/workbook.xml");
  assert.ok(workbookEntry);
  const workbookXml = textDecoder.decode(
    await readEntryData(workbookEntry, nodeCompressionAdapter)
  );

  assert.match(workbookXml, /<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"\/>/);
});

test("formula recalc insertion preserves workbook namespace prefixes", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  pkg.setText(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </x:sheets>
</x:workbook>`
  );
  const workbook = await Workbook.fromPackage(pkg);

  await workbook.patchCell("Sheet1", "C3", { formula: "=SUM(A1:B2)" });
  const outputZip = parseZip(await workbook.write());
  const workbookEntry = outputZip.entries.find((entry) => entry.name === "xl/workbook.xml");
  assert.ok(workbookEntry);
  const workbookXml = textDecoder.decode(
    await readEntryData(workbookEntry, nodeCompressionAdapter)
  );

  assert.match(workbookXml, /<x:calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"\/>/);
});

test("formula patches remove stale calculation chain parts", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeCalcChain: true }));

  await workbook.patchCell("Sheet1", "C3", { formula: "=SUM(A1:B2)" });
  const outputZip = parseZip(await workbook.write());
  const names = outputZip.entries.map((entry) => entry.name);

  assert.equal(names.includes("xl/calcChain.xml"), false);

  const relsEntry = outputZip.entries.find((entry) => entry.name === "xl/_rels/workbook.xml.rels");
  assert.ok(relsEntry);
  const relsXml = textDecoder.decode(await readEntryData(relsEntry, nodeCompressionAdapter));
  assert.doesNotMatch(relsXml, /calcChain/);

  const contentTypesEntry = outputZip.entries.find((entry) => entry.name === "[Content_Types].xml");
  assert.ok(contentTypesEntry);
  const contentTypesXml = textDecoder.decode(
    await readEntryData(contentTypesEntry, nodeCompressionAdapter)
  );
  assert.doesNotMatch(contentTypesXml, /calcChain/);
});

test("replacing an existing formula with a value removes stale calculation chain parts", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeCalcChain: true }));
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c>',
      '<c r="A1"><f>1+1</f><v>2</v></c>'
    )
  );
  const workbook = await Workbook.fromPackage(pkg);

  await workbook.patchCell("Sheet1", "A1", 7);
  const outputZip = parseZip(await workbook.write());
  const names = outputZip.entries.map((entry) => entry.name);

  assert.equal(names.includes("xl/calcChain.xml"), false);
});

test("value edits that feed formulas mark the workbook for recalculation", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const worksheetXml = await pkg.readText("xl/worksheets/sheet1.xml");
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    worksheetXml.replace(
      '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c></row>',
      '<row r="1"><c r="A1"><v>1</v></c><c r="C1"><f>A1*2</f><v>2</v></c></row>'
    )
  );
  const workbook = await Workbook.fromPackage(pkg);

  await workbook.patchCell("Sheet1", "A1", 3);
  const outputZip = parseZip(await workbook.write());
  const workbookEntry = outputZip.entries.find((entry) => entry.name === "xl/workbook.xml");
  assert.ok(workbookEntry);
  const workbookXml = textDecoder.decode(
    await readEntryData(workbookEntry, nodeCompressionAdapter)
  );

  assert.match(workbookXml, /<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"\/>/);
  assert.equal(
    workbook
      .diagnostics()
      .some((diagnostic) => diagnostic.code === "FORMULA_DEPENDENCIES_RECALCULATED"),
    true
  );
});

test("cell patches preserve macro-enabled workbook parts", async () => {
  const original = await createMinimalWorkbook({ includeMacro: true });
  const originalZip = parseZip(original);
  const originalVba = originalZip.entries.find((entry) => entry.name === "xl/vbaProject.bin");
  assert.ok(originalVba);

  const workbook = await openWorkbook(original);
  await workbook.patchCell("Sheet1", "B2", "macro safe");
  const outputZip = parseZip(await workbook.write());
  const outputVba = outputZip.entries.find((entry) => entry.name === "xl/vbaProject.bin");
  assert.ok(outputVba);

  assert.deepEqual(outputVba.compressedData, originalVba.compressedData);

  const relsEntry = outputZip.entries.find((entry) => entry.name === "xl/_rels/workbook.xml.rels");
  assert.ok(relsEntry);
  const relsXml = textDecoder.decode(await readEntryData(relsEntry, nodeCompressionAdapter));
  assert.match(relsXml, /vbaProject/);

  const contentTypesEntry = outputZip.entries.find((entry) => entry.name === "[Content_Types].xml");
  assert.ok(contentTypesEntry);
  const contentTypesXml = textDecoder.decode(
    await readEntryData(contentTypesEntry, nodeCompressionAdapter)
  );
  assert.match(contentTypesXml, /macroEnabled\.main\+xml/);
  assert.match(contentTypesXml, /vbaProject/);
});

test("cell patches preserve merge cells and hyperlink relationships", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeHyperlink: true, includeMerge: true })
  );

  assert.deepEqual(workbook.diagnostics(), [
    {
      severity: "warning",
      code: "EXTERNAL_RELATIONSHIP_PRESERVED",
      message: "Preserved external relationship rIdHyperlink1",
      part: "xl/worksheets/sheet1.xml"
    }
  ]);

  await workbook.patchCell("Sheet1", "C3", "safe");
  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(sheetXml, /<mergeCell ref="A1:B1"\/>/);
  assert.match(sheetXml, /<hyperlink ref="A1" r:id="rIdHyperlink1"\/>/);

  const rels = outputZip.entries.find(
    (entry) => entry.name === "xl/worksheets/_rels/sheet1.xml.rels"
  );
  assert.ok(rels);
  const relsXml = textDecoder.decode(await readEntryData(rels, nodeCompressionAdapter));
  assert.match(relsXml, /Target="https:\/\/example.com"/);
  assert.match(relsXml, /TargetMode="External"/);
});

test("cell patches preserve validation, formatting, drawing, chart, media, and hidden sheets", async () => {
  const original = await createMinimalWorkbook({
    includeConditionalFormatting: true,
    includeDataValidation: true,
    includeDrawing: true,
    includeHiddenSheet: true
  });
  const originalZip = parseZip(original);
  const originalImage = originalZip.entries.find((entry) => entry.name === "xl/media/image1.png");
  const originalHiddenSheet = originalZip.entries.find(
    (entry) => entry.name === "xl/worksheets/sheet2.xml"
  );
  assert.ok(originalImage);
  assert.ok(originalHiddenSheet);

  const workbook = await openWorkbook(original);
  await workbook.patchCell("Sheet1", "C3", "preserve complex parts");
  const outputZip = parseZip(await workbook.write());

  const outputImage = outputZip.entries.find((entry) => entry.name === "xl/media/image1.png");
  const outputHiddenSheet = outputZip.entries.find(
    (entry) => entry.name === "xl/worksheets/sheet2.xml"
  );
  assert.ok(outputImage);
  assert.ok(outputHiddenSheet);
  assert.deepEqual(outputImage.compressedData, originalImage.compressedData);
  assert.deepEqual(outputHiddenSheet.compressedData, originalHiddenSheet.compressedData);

  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(sheetXml, /<conditionalFormatting sqref="A1:A10">/);
  assert.match(sheetXml, /<dataValidation type="whole" operator="between"/);
  assert.match(sheetXml, /<drawing r:id="rIdDrawing1"\/>/);

  const drawing = outputZip.entries.find((entry) => entry.name === "xl/drawings/drawing1.xml");
  const chart = outputZip.entries.find((entry) => entry.name === "xl/charts/chart1.xml");
  const workbookEntry = outputZip.entries.find((entry) => entry.name === "xl/workbook.xml");
  assert.ok(drawing);
  assert.ok(chart);
  assert.ok(workbookEntry);
  assert.match(
    textDecoder.decode(await readEntryData(drawing, nodeCompressionAdapter)),
    /r:id="rIdChart1"/
  );
  assert.match(
    textDecoder.decode(await readEntryData(chart, nodeCompressionAdapter)),
    /<a:t>Revenue<\/a:t>/
  );
  assert.match(
    textDecoder.decode(await readEntryData(workbookEntry, nodeCompressionAdapter)),
    /state="hidden"/
  );
});

test("diagnostics report macro preservation and calc chain invalidation", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeCalcChain: true, includeMacro: true })
  );

  assert.equal(
    workbook.diagnostics().some((diagnostic) => diagnostic.code === "MACRO_PROJECT_PRESERVED"),
    true
  );

  await workbook.patchCell("Sheet1", "C3", { formula: "=SUM(A1:B2)" });

  assert.equal(
    workbook.diagnostics().some((diagnostic) => diagnostic.code === "FORMULA_CALC_CHAIN_REMOVED"),
    true
  );
});

test("data edits report defined name, chart, and table impact diagnostics", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({
      includeDefinedName: true,
      includeDrawing: true,
      includeTable: true
    })
  );

  await workbook.patchCell("Sheet1", "C3", "changed");
  await workbook.patchCell("Sheet1", "D4", "changed again");

  assert.deepEqual(
    workbook
      .diagnostics()
      .filter((diagnostic) =>
        [
          "DEFINED_NAMES_MAY_NEED_REVIEW",
          "CHARTS_MAY_NEED_REFRESH",
          "WORKSHEET_TABLES_NOT_RESIZED"
        ].includes(diagnostic.code)
      )
      .map((diagnostic) => diagnostic.code),
    ["DEFINED_NAMES_MAY_NEED_REVIEW", "CHARTS_MAY_NEED_REFRESH", "WORKSHEET_TABLES_NOT_RESIZED"]
  );
});

test("table replacement reports dependent workbook structures without table resize warning", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({
      includeDefinedName: true,
      includeDrawing: true,
      includeTable: true
    })
  );

  await workbook.replaceTableRows("RevenueTable", [["Updated", 42]]);
  const diagnostics = workbook.diagnostics();

  assert.equal(
    diagnostics.some((diagnostic) => diagnostic.code === "DEFINED_NAMES_MAY_NEED_REVIEW"),
    true
  );
  assert.equal(
    diagnostics.some((diagnostic) => diagnostic.code === "CHARTS_MAY_NEED_REFRESH"),
    true
  );
  assert.equal(
    diagnostics.some((diagnostic) => diagnostic.code === "WORKSHEET_TABLES_NOT_RESIZED"),
    false
  );
});

test("replaces basic table body rows and updates the table ref", async () => {
  const workbook = await openWorkbook(await createMinimalWorkbook({ includeTable: true }));

  const table = await workbook.replaceTableRows("RevenueTable", [
    ["New", 10],
    ["Growth", 20]
  ]);
  assert.equal(table.ref, "A1:B3");

  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(
    sheetXml,
    /<row r="2"><c r="A2" t="inlineStr"><is><t>New<\/t><\/is><\/c><c r="B2"><v>10<\/v><\/c><\/row>/
  );
  assert.match(
    sheetXml,
    /<row r="3"><c r="A3" t="inlineStr"><is><t>Growth<\/t><\/is><\/c><c r="B3"><v>20<\/v><\/c><\/row>/
  );

  const tableEntry = outputZip.entries.find((entry) => entry.name === "xl/tables/table1.xml");
  assert.ok(tableEntry);
  const tableXml = textDecoder.decode(await readEntryData(tableEntry, nodeCompressionAdapter));
  assert.match(tableXml, /ref="A1:B3"/);
  assert.match(tableXml, /<autoFilter ref="A1:B3"\/>/);
});

test("table row replacement preserves body styles and shrinks dimensions", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({
      includeTable: true,
      styledTableBody: true,
      tableRows: [
        ["Old", 1],
        ["Older", 2],
        ["Oldest", 3]
      ]
    })
  );

  const table = await workbook.replaceTableRows("RevenueTable", [["Fresh", 99]]);
  assert.equal(table.ref, "A1:B2");

  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(sheetXml, /<dimension ref="A1:B2"\/>/);
  assert.match(
    sheetXml,
    /<row r="2" s="2" customFormat="1"><c r="A2" s="3" t="inlineStr"><is><t>Fresh<\/t><\/is><\/c><c r="B2" s="4"><v>99<\/v><\/c><\/row>/
  );
  assert.doesNotMatch(sheetXml, /Oldest/);

  const tableEntry = outputZip.entries.find((entry) => entry.name === "xl/tables/table1.xml");
  assert.ok(tableEntry);
  const tableXml = textDecoder.decode(await readEntryData(tableEntry, nodeCompressionAdapter));
  assert.match(tableXml, /ref="A1:B2"/);
  assert.match(tableXml, /<autoFilter ref="A1:B2"\/>/);
});

test("table row replacement preserves and moves totals rows", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeTable: true, includeTableTotals: true })
  );

  const table = await workbook.replaceTableRows("RevenueTable", [
    ["Fresh", 99],
    ["Future", 101]
  ]);
  assert.equal(table.ref, "A1:B4");

  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));
  assert.match(
    sheetXml,
    /<row r="4"><c r="A4" t="inlineStr"><is><t>Total<\/t><\/is><\/c><c r="B4"><f>SUBTOTAL\(109,B2:B3\)<\/f><v>1<\/v><\/c><\/row>/
  );

  const tableEntry = outputZip.entries.find((entry) => entry.name === "xl/tables/table1.xml");
  assert.ok(tableEntry);
  const tableXml = textDecoder.decode(await readEntryData(tableEntry, nodeCompressionAdapter));
  assert.match(tableXml, /ref="A1:B4"/);
  assert.match(tableXml, /<autoFilter ref="A1:B4"\/>/);
});

test("table row replacement handles namespace-prefixed worksheets", async () => {
  const pkg = await openPackage(await createMinimalWorkbook({ includeTable: true }));
  pkg.setText(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:dimension ref="A1:B2"/>
  <x:sheetData>
    <x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Name</x:t></x:is></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>Amount</x:t></x:is></x:c></x:row>
    <x:row r="2"><x:c r="A2" t="inlineStr"><x:is><x:t>Old</x:t></x:is></x:c><x:c r="B2"><x:v>1</x:v></x:c></x:row>
  </x:sheetData>
  <x:tableParts count="1"><x:tablePart r:id="rIdTable1"/></x:tableParts>
</x:worksheet>`
  );
  const workbook = await Workbook.fromPackage(pkg);

  await workbook.replaceTableRows("RevenueTable", [["Prefixed", 7]]);
  const outputZip = parseZip(await workbook.write());
  const sheet = outputZip.entries.find((entry) => entry.name === "xl/worksheets/sheet1.xml");
  assert.ok(sheet);
  const sheetXml = textDecoder.decode(await readEntryData(sheet, nodeCompressionAdapter));

  assert.match(
    sheetXml,
    /<x:row r="2"><x:c r="A2" t="inlineStr"><x:is><x:t>Prefixed<\/x:t><\/x:is><\/x:c><x:c r="B2"><x:v>7<\/x:v><\/x:c><\/x:row>/
  );
  assert.match(sheetXml, /<x:dimension ref="A1:B2"\/>/);
});

test("table formula writes invalidate stale calculation chains", async () => {
  const workbook = await openWorkbook(
    await createMinimalWorkbook({ includeCalcChain: true, includeTable: true })
  );

  await workbook.replaceTableRows("RevenueTable", [
    ["Formula", { formula: "=SUM(1,2)", result: 3 }]
  ]);
  const outputZip = parseZip(await workbook.write());
  const names = outputZip.entries.map((entry) => entry.name);

  assert.equal(names.includes("xl/calcChain.xml"), false);
});
