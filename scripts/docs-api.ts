#!/usr/bin/env tsx
/**
 * Regenerate the generated API reference under
 * website/content/docs/api-reference/.
 *
 * Regeneration contract (idempotent):
 *   1. Build every package so fresh `dist/*.d.ts` declarations exist
 *      (TypeDoc's `packages` strategy reads each package's published entry
 *      point, and `dist/` is gitignored so it must be rebuilt).
 *   2. Delete everything already inside the api-reference directory EXCEPT
 *      the hand-written `meta.json` (Fumadocs sidebar config). This removes
 *      stale generated pages for exports that no longer exist.
 *   3. Run TypeDoc, which writes `index.md` (the landing, merged from
 *      docs/api-landing.md) plus the `core/`, `node/`, and `browser/`
 *      sections.
 *
 * Files that are NEVER touched by this script:
 *   - website/content/docs/api-reference/meta.json  (hand-written)
 *   - docs/api-landing.md                           (hand-written landing prose)
 */
import { spawnSync } from "node:child_process";
import { readdirSync, rmSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import process from "node:process";

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), "..");
const outDir = join(repoRoot, "website", "content", "docs", "api-reference");
const preserve = new Set(["meta.json"]);

function run(command: string, args: string[]): void {
  const result = spawnSync(command, args, {
    stdio: "inherit",
    shell: false,
    cwd: repoRoot
  });
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

// 1. Fresh declarations.
run(process.execPath, ["node_modules/tsx/dist/cli.mjs", "scripts/build.ts"]);

// 2. Clean generated output, preserving hand-written files.
for (const entry of readdirSync(outDir)) {
  if (preserve.has(entry)) {
    continue;
  }
  rmSync(join(outDir, entry), { force: true, recursive: true });
}

// 3. Generate.
run(process.execPath, ["node_modules/typedoc/bin/typedoc", "--options", "docs/typedoc.json"]);
