import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { resolve } from "node:path";
import test from "node:test";
import { validateWorkbookFile } from "../packages/node/src/index.ts";
import { buildTemplates } from "../scripts/build-templates.ts";

test("starter templates build into valid workbooks", async () => {
  const directory = await mkdtemp(resolve(tmpdir(), "ironsheet-templates-"));

  try {
    await copyTemplateManifest(directory);
    const written = await buildTemplates(directory);

    assert.deepEqual(written.map((path) => path.slice(directory.length + 1)).sort(), [
      "templates/generated/dashboard-template.xlsx",
      "templates/generated/large-export-template.xlsx",
      "templates/generated/macro-model-template.xlsm",
      "templates/generated/pivot-source-template.xlsx",
      "templates/generated/styled-report-template.xlsx"
    ]);

    for (const workbookPath of written) {
      const report = await validateWorkbookFile(workbookPath);
      assert.deepEqual(report.summary, { errors: 0, warnings: 0, infos: 0 });
    }
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});

async function copyTemplateManifest(directory: string): Promise<void> {
  await mkdir(resolve(directory, "templates"), { recursive: true });
  await writeFile(
    resolve(directory, "templates/manifest.json"),
    await readFile(resolve("templates/manifest.json"))
  );
}
