import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { resolve } from "node:path";
import process from "node:process";
import { spawnSync } from "node:child_process";
import test from "node:test";
import { openPackage, openWorkbook } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("CLI validate exits successfully for valid workbooks", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-cli-"));

  try {
    const workbookPath = resolve(directory, "valid.xlsx");
    await writeFile(workbookPath, await createMinimalWorkbook());

    const result = runCli(["validate", workbookPath]);

    assert.equal(result.status, 0);
    assert.match(result.stdout, /"errors": 0/);
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("CLI validate exits nonzero for validation errors", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-cli-"));

  try {
    const workbookPath = resolve(directory, "invalid.xlsx");
    const pkg = await openPackage(await createMinimalWorkbook());
    const relsXml = await pkg.readText("xl/_rels/workbook.xml.rels");
    pkg.setText(
      "xl/_rels/workbook.xml.rels",
      relsXml.replace(
        "</Relationships>",
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
      )
    );
    await writeFile(workbookPath, await pkg.write());

    const result = runCli(["validate", workbookPath]);

    assert.equal(result.status, 1);
    assert.match(result.stdout, /RELATIONSHIP_ID_DUPLICATE/);
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("CLI template-manifest reports patchable anchors", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-cli-"));

  try {
    const workbookPath = resolve(directory, "template.xlsx");
    await writeFile(
      workbookPath,
      await createMinimalWorkbook({
        includeDefinedName: true,
        includeDrawing: true,
        includeTable: true
      })
    );

    const result = runCli(["template-manifest", workbookPath]);

    assert.equal(result.status, 0);
    assert.match(result.stdout, /"namedRanges"/);
    assert.match(result.stdout, /"RevenueRange"/);
    assert.match(result.stdout, /"RevenueTable"/);
    assert.match(result.stdout, /"xl\/media\/image1\.png"/);
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("CLI patch-named-range updates a defined-name target", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-cli-"));

  try {
    const inputPath = resolve(directory, "input.xlsx");
    const outputPath = resolve(directory, "output.xlsx");
    await writeFile(inputPath, await createMinimalWorkbook({ includeDefinedName: true }));

    const result = runCli([
      "patch-named-range",
      inputPath,
      outputPath,
      "RevenueRange",
      '[["West",100],["East",200]]'
    ]);

    assert.equal(result.status, 0);
    const workbook = await openWorkbook(new Uint8Array(await readFile(outputPath)));
    assert.deepEqual(
      (await workbook.readNamedRange("RevenueRange")).cells.map((row) =>
        row.map((cell) => cell?.value)
      ),
      [
        ["West", 100],
        ["East", 200]
      ]
    );
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("CLI render-template-safe reports validation and writes output", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-cli-"));

  try {
    const inputPath = resolve(directory, "input.xlsx");
    const outputPath = resolve(directory, "output.xlsx");
    const patchPath = resolve(directory, "patch.json");
    await writeFile(inputPath, await createMinimalWorkbook({ includeDefinedName: true }));
    await writeFile(patchPath, '{"names":[{"name":"RevenueRange","values":[["North",10]]}]}');

    const result = runCli(["render-template-safe", inputPath, outputPath, `@${patchPath}`]);

    assert.equal(result.status, 0);
    assert.match(result.stdout, /"wrote": true/);
    assert.match(result.stdout, /"errors": 0/);

    const workbook = await openWorkbook(new Uint8Array(await readFile(outputPath)));
    assert.equal((await workbook.readCell("Sheet1", "A1"))?.value, "North");
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

test("CLI render-template accepts image files in JSON patches", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-cli-"));

  try {
    const inputPath = resolve(directory, "input.xlsx");
    const outputPath = resolve(directory, "output.xlsx");
    const imagePath = resolve(directory, "logo.png");
    const replacement = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 4, 3, 2, 1]);
    await writeFile(inputPath, await createMinimalWorkbook({ includeDrawing: true }));
    await writeFile(imagePath, replacement);

    const result = runCli([
      "render-template",
      inputPath,
      outputPath,
      JSON.stringify({
        images: [{ imagePartName: "xl/media/image1.png", path: imagePath }]
      })
    ]);

    assert.equal(result.status, 0);
    assert.match(result.stdout, /"images": 1/);

    const pkg = await openPackage(new Uint8Array(await readFile(outputPath)));
    assert.deepEqual(Array.from(await pkg.readPart("xl/media/image1.png")), [...replacement]);
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

function runCli(args: string[]): { status: number | null; stdout: string; stderr: string } {
  const result = spawnSync(
    process.execPath,
    ["--import", "tsx", "packages/cli/src/cli.ts", ...args],
    {
      cwd: resolve("."),
      encoding: "utf8",
      shell: false
    }
  );

  return {
    status: result.status,
    stdout: result.stdout,
    stderr: result.stderr
  };
}
