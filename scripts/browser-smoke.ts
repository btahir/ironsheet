#!/usr/bin/env tsx
import { mkdir, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import process from "node:process";
import { build } from "esbuild";

const entryPath = resolve(".cache/browser-smoke/entry.ts");

await mkdir(resolve(".cache/browser-smoke"), { recursive: true });
await writeFile(
  entryPath,
  `import { browserCompressionAdapter, inspectWorkbookArchiveFromBytes, openWorkbookFromBytes, writeWorkbookToBlobSafely } from "@ironsheet/browser";

export async function smokeWorkbook(bytes: Uint8Array): Promise<Blob> {
  const archive = inspectWorkbookArchiveFromBytes(bytes);
  if (!archive.accepted) {
    throw new Error(archive.issues[0]?.message ?? "Workbook archive was rejected");
  }

  const workbook = await openWorkbookFromBytes(bytes);
  await workbook.inspect();
  const result = await writeWorkbookToBlobSafely(workbook);
  if (!result.wrote || result.blob === undefined) {
    throw new Error("Workbook validation failed");
  }
  return result.blob;
}

export const smokeAdapter = browserCompressionAdapter;
`
);

try {
  const result = await build({
    absWorkingDir: process.cwd(),
    bundle: true,
    entryPoints: [entryPath],
    format: "esm",
    metafile: true,
    platform: "browser",
    sourcemap: false,
    target: ["es2022"],
    write: false
  });

  const bundle = result.outputFiles[0]?.text;
  if (bundle === undefined) {
    throw new Error("Browser smoke did not produce a bundle");
  }

  const forbiddenPatterns = [
    /\bnode:/,
    /\brequire\s*\(/,
    /\bfrom\s+["'](?:fs|path|zlib|stream|buffer)["']/
  ];
  const match = forbiddenPatterns.find((pattern) => pattern.test(bundle));
  if (match !== undefined) {
    throw new Error(`Browser bundle contains forbidden Node-only code matching ${match}`);
  }

  const inputs = Object.keys(result.metafile.inputs);
  const hasBrowser = inputs.some((input) =>
    /packages\/browser\/(?:src|dist)\/index\.(?:ts|js)$/.test(input)
  );
  const hasCore = inputs.some((input) =>
    /packages\/core\/(?:src|dist)\/index\.(?:ts|js)$/.test(input)
  );
  if (!hasBrowser || !hasCore) {
    throw new Error("Browser bundle did not include @ironsheet/browser and @ironsheet/core");
  }

  console.log(
    `browser-smoke: bundled @ironsheet/browser for browser (${bundle.length} bytes, ${inputs.length} inputs)`
  );
} finally {
  await rm(resolve(".cache/browser-smoke"), { force: true, recursive: true });
}
