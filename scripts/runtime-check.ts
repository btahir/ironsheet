#!/usr/bin/env tsx
import { readFileSync } from "node:fs";
import { join } from "node:path";

const coreFiles = [
  "packages/core/src/index.ts",
  "packages/core/src/address.ts",
  "packages/core/src/diff.ts",
  "packages/core/src/errors.ts",
  "packages/core/src/opc.ts",
  "packages/core/src/shared-strings.ts",
  "packages/core/src/table.ts",
  "packages/core/src/workbook.ts",
  "packages/core/src/worksheet.ts",
  "packages/core/src/xml.ts",
  "packages/core/src/zip/binary.ts",
  "packages/core/src/zip/crc32.ts",
  "packages/core/src/zip/index.ts"
];

const forbiddenImports = [
  "node:",
  '"fs"',
  '"path"',
  '"stream"',
  '"zlib"',
  '"buffer"',
  "'fs'",
  "'path'",
  "'stream'",
  "'zlib'",
  "'buffer'"
];

const violations: string[] = [];

for (const file of coreFiles) {
  const source = readFileSync(join(process.cwd(), file), "utf8");
  for (const forbidden of forbiddenImports) {
    if (source.includes(forbidden)) {
      violations.push(`${file}: forbidden runtime-neutral core import/reference ${forbidden}`);
    }
  }
}

if (violations.length > 0) {
  console.error(violations.join("\n"));
  process.exit(1);
}

console.log("runtime: core package has no obvious Node-only imports");
