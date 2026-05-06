#!/usr/bin/env tsx
import { readdirSync, readFileSync } from "node:fs";
import { join, normalize, relative, resolve } from "node:path";

type PackageJson = {
  dependencies?: Record<string, string>;
  optionalDependencies?: Record<string, string>;
};

const root = process.cwd();
const coreRoot = resolve(root, "packages/core/src");
const forbiddenBareImports = new Set(["buffer", "fs", "path", "stream", "zlib"]);
const importPattern =
  /(?:import|export)\s+(?:[^'"]*?\s+from\s+)?["']([^"']+)["']|import\(\s*["']([^"']+)["']\s*\)/g;

const violations: string[] = [];

for (const file of listTypeScriptFiles(coreRoot)) {
  const relativeFile = relative(root, file);
  const source = readFileSync(file, "utf8");
  for (const match of source.matchAll(importPattern)) {
    const specifier = match[1] ?? match[2];
    if (specifier === undefined) {
      continue;
    }

    if (specifier.startsWith("node:") || forbiddenBareImports.has(specifier)) {
      violations.push(`${relativeFile}: forbidden Node-only import ${specifier}`);
      continue;
    }

    if (specifier.startsWith(".")) {
      const resolved = normalize(resolve(file, "..", specifier));
      if (!resolved.startsWith(coreRoot)) {
        violations.push(`${relativeFile}: core import escapes packages/core/src: ${specifier}`);
      }
    }
  }
}

const packageJson = JSON.parse(readFileSync(join(root, "package.json"), "utf8")) as PackageJson;
for (const dependencyField of ["dependencies", "optionalDependencies"] as const) {
  const dependencies = packageJson[dependencyField] ?? {};
  for (const dependency of Object.keys(dependencies)) {
    violations.push(
      `package.json: runtime dependency is not allowed yet: ${dependencyField}.${dependency}`
    );
  }
}

if (violations.length > 0) {
  console.error(violations.join("\n"));
  process.exit(1);
}

console.log("runtime: core package has no Node-only imports or runtime dependencies");

function listTypeScriptFiles(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) {
      return listTypeScriptFiles(path);
    }

    return entry.isFile() && path.endsWith(".ts") ? [path] : [];
  });
}
