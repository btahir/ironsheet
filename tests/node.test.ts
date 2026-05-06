import assert from "node:assert/strict";
import { constants } from "node:fs";
import { access, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { resolve } from "node:path";
import test from "node:test";
import {
  insertWorkbookImageFile,
  mutateWorkbookFile,
  openWorkbook,
  preflightWorkbookTemplate,
  renderWorkbookTemplateSafely
} from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("safe template render writes only after validation and reports diff", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-node-"));

  try {
    const inputPath = resolve(directory, "input.xlsx");
    const outputPath = resolve(directory, "output.xlsx");
    await writeFile(inputPath, await createMinimalWorkbook({ includeDefinedName: true }));

    const report = await renderWorkbookTemplateSafely(inputPath, outputPath, {
      names: [{ name: "RevenueRange", values: [["North", 10]] }]
    });

    assert.equal(report.wrote, true);
    assert.equal(report.validation.summary.errors, 0);
    assert.equal(report.diff.summary.changed > 0, true);
    assert.deepEqual(report.render.applied, {
      cells: 0,
      images: 0,
      names: 1,
      ranges: 0,
      tables: 0
    });

    const workbook = await openWorkbook(new Uint8Array(await readFile(outputPath)));
    assert.equal((await workbook.readCell("Sheet1", "A1"))?.value, "North");
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("safe mutation refuses to write packages with validation errors", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-node-"));

  try {
    const inputPath = resolve(directory, "input.xlsx");
    const outputPath = resolve(directory, "output.xlsx");
    await writeFile(inputPath, await createMinimalWorkbook());

    const report = await mutateWorkbookFile(inputPath, outputPath, async (workbook) => {
      const relsXml = await workbook.pkg.readText("xl/_rels/workbook.xml.rels");
      workbook.pkg.setText(
        "xl/_rels/workbook.xml.rels",
        relsXml.replace(
          "</Relationships>",
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        )
      );
    });

    assert.equal(report.wrote, false);
    assert.equal(report.validation.summary.errors > 0, true);
    await assert.rejects(() => access(outputPath, constants.F_OK));
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("node adapter preflights template patches without writing output", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-node-"));

  try {
    const inputPath = resolve(directory, "input.xlsx");
    await writeFile(inputPath, await createMinimalWorkbook({ includeDefinedName: true }));

    const result = await preflightWorkbookTemplate(inputPath, {
      names: [{ name: "RevenueRange", values: [["North", 10]] }]
    });

    assert.deepEqual(result.counts, { cells: 0, images: 0, names: 1, ranges: 0, tables: 0 });
    assert.equal(result.targets.names[0]?.ref, "A1:B2");
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("node adapter inserts image files", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-node-"));

  try {
    const inputPath = resolve(directory, "input.xlsx");
    const outputPath = resolve(directory, "output.xlsx");
    const imagePath = resolve(directory, "logo.png");
    await writeFile(inputPath, await createMinimalWorkbook());
    await writeFile(imagePath, new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 5]));

    const image = await insertWorkbookImageFile(inputPath, outputPath, "Sheet1", imagePath);
    const workbook = await openWorkbook(new Uint8Array(await readFile(outputPath)));

    assert.equal(image.imagePartName, "xl/media/image1.png");
    assert.deepEqual(
      (await workbook.images()).map((candidate) => candidate.imagePartName),
      ["xl/media/image1.png"]
    );
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});
