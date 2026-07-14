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
import { readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
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

// 4. Dot-prefix bare-relative Markdown link targets in the generated output.
//
// Why this pass exists: Fumadocs' resolveHref only rewrites Markdown links whose
// href starts with "./" or "../". TypeDoc emits sibling/child links as bare
// relative paths (e.g. "core/index.md", "classes/Workbook.md", "PackageError.md"),
// which Fumadocs passes through verbatim so they 404 in the browser. Prefixing
// "./" makes resolveHref resolve them to the correct page URL.
//
// Scope is deliberately conservative: only link targets ending in ".md" (optionally
// with a #fragment) are touched, and never those already prefixed with "./"/"../",
// absolute ("/"), external (http:/https:/mailto:), or pure anchors ("#"). A single
// regex over the whole file (rather than an AST walk) is acceptable because a
// literal `](…​.md)` sequence inside a fenced code block is vanishingly unlikely in
// TypeDoc output; the tradeoff is that such content would also be rewritten.
const BARE_MD_LINK = /\]\((?!\.\/|\.\.\/|\/|https?:|#|mailto:)([^)]+\.md(?:#[^)]*)?)\)/g;

function rewriteLinks(dir: string): void {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      rewriteLinks(full);
    } else if (entry.isFile() && entry.name.endsWith(".md")) {
      const original = readFileSync(full, "utf8");
      const rewritten = original.replace(BARE_MD_LINK, "](./$1)");
      if (rewritten !== original) {
        writeFileSync(full, rewritten);
      }
    }
  }
}

rewriteLinks(outDir);
