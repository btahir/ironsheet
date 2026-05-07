import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { resolve } from "node:path";
import test from "node:test";
import {
  mutateWorkbookFile,
  readWorkbook,
  renderWorkbookTemplateSafely,
  validateWorkbookFile
} from "@ironsheet/node";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("documented mutateWorkbookFile flow writes a validated workbook", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-docs-mutate-"));

  try {
    const inputPath = resolve(directory, "template.xlsm");
    const outputPath = resolve(directory, "output.xlsm");
    await writeFile(
      inputPath,
      await createMinimalWorkbook({
        includeDefinedName: true,
        includeMacro: true,
        includeTable: true
      })
    );

    const report = await mutateWorkbookFile(inputPath, outputPath, async (workbook) => {
      await workbook.patchNamedRange("RevenueRange", [
        ["Region", "Amount"],
        ["North", 42000]
      ]);

      await workbook.replaceTableRows("RevenueTable", [
        ["North", 42000],
        ["South", 31500]
      ]);
    });

    assert.equal(report.wrote, true);
    assert.equal(report.validation.summary.errors, 0);
    assert.equal(
      report.diff.entries.some((entry) => entry.contentChanged),
      true
    );
    assert.deepEqual((await validateWorkbookFile(outputPath)).summary, {
      errors: 0,
      warnings: 0,
      infos: 0
    });
    const workbook = await readWorkbook(outputPath);
    assert.deepEqual(
      Array.from(await workbook.pkg.readPart("xl/vbaProject.bin")),
      [0xca, 0xfe, 0xba, 0xbe, 0, 1]
    );
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("documented safe template render flow preflights and writes a validated workbook", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-docs-template-"));

  try {
    const inputPath = resolve(directory, "template.xlsx");
    const outputPath = resolve(directory, "report.xlsx");
    const logoPngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 9]);
    await writeFile(
      inputPath,
      await createMinimalWorkbook({
        includeDefinedName: true,
        includeDrawing: true,
        includeTable: true
      })
    );

    const report = await renderWorkbookTemplateSafely(inputPath, outputPath, {
      cells: [{ sheetName: "Sheet1", address: "D1", value: "Q1" }],
      ranges: [{ sheetName: "Sheet1", startAddress: "E1", values: [["Name", "Amount"]] }],
      names: [
        {
          name: "RevenueRange",
          values: [
            ["Region", "Amount"],
            ["North", 42000]
          ]
        }
      ],
      tables: [{ tableName: "RevenueTable", rows: [["North", 42000]] }],
      images: [{ imagePartName: "xl/media/image1.png", data: logoPngBytes }]
    });

    assert.equal(report.wrote, true);
    assert.deepEqual(report.render.applied, {
      cells: 1,
      images: 1,
      names: 1,
      ranges: 1,
      tables: 1
    });
    assert.deepEqual((await validateWorkbookFile(outputPath)).summary, {
      errors: 0,
      warnings: 0,
      infos: 0
    });
    const workbook = await readWorkbook(outputPath);
    assert.deepEqual(
      Array.from(await workbook.pkg.readPart("xl/media/image1.png")),
      [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 9]
    );
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});
