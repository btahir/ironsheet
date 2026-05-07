#!/usr/bin/env tsx
import { spawnSync, type SpawnSyncOptionsWithStringEncoding } from "node:child_process";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { extname, join } from "node:path";
import process from "node:process";

type PackageJson = {
  scripts?: Record<string, string>;
};

const blockedPaths = new Set<string>(["IRONSHEET_SPEC.md"]);
const optionalChecks = [
  "format:check",
  "lint",
  "typecheck",
  "test",
  "runtime:check",
  "build",
  "browser:smoke"
] as const;
const forbiddenScriptExtensions = new Set<string>([".js", ".mjs", ".cjs"]);

function run(command: string, args: string[]): void {
  const result = spawnSync(command, args, {
    stdio: "inherit",
    shell: false
  });

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

function output(command: string, args: string[]): string {
  const options: SpawnSyncOptionsWithStringEncoding = {
    encoding: "utf8",
    shell: false
  };
  const result = spawnSync(command, args, options);

  if (result.status !== 0) {
    return "";
  }

  return result.stdout.trim();
}

function fail(message: string): never {
  console.error(`verify: ${message}`);
  process.exit(1);
}

function readPackageJson(): PackageJson {
  if (!existsSync("package.json")) {
    fail("package.json not found");
  }

  return JSON.parse(readFileSync("package.json", "utf8")) as PackageJson;
}

function assertNodeVersion(): void {
  const nodeMajor = Number.parseInt(process.versions.node.split(".")[0] ?? "0", 10);
  if (nodeMajor < 18) {
    fail(`Node.js >=18 is required; found ${process.versions.node}`);
  }
}

function assertPlanningDocsNotStaged(): void {
  const insideGit = output("git", ["rev-parse", "--is-inside-work-tree"]);
  if (insideGit !== "true") {
    return;
  }

  const staged = output("git", ["diff", "--cached", "--name-only"])
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const blocked = staged.filter((file) => blockedPaths.has(file));
  if (blocked.length > 0) {
    fail(`planning docs must not be committed: ${blocked.join(", ")}`);
  }
}

function assertScriptsAreTypeScript(): void {
  if (!existsSync("scripts")) {
    return;
  }

  const forbidden = readdirSync("scripts", { recursive: true, withFileTypes: true })
    .filter((entry) => entry.isFile())
    .map((entry) => join(entry.parentPath, entry.name))
    .filter((file) => forbiddenScriptExtensions.has(extname(file)));

  if (forbidden.length > 0) {
    fail(`scripts must be TypeScript, not JavaScript: ${forbidden.join(", ")}`);
  }
}

function runConfiguredChecks(packageJson: PackageJson): void {
  const scripts = packageJson.scripts ?? {};
  const checksToRun = optionalChecks.filter((name) => Object.hasOwn(scripts, name));

  if (checksToRun.length === 0) {
    console.log("verify: no optional checks configured yet; repository guard checks passed");
    return;
  }

  for (const check of checksToRun) {
    console.log(`verify: running npm run ${check}`);
    run("npm", ["run", check]);
  }

  console.log("verify: all checks passed");
}

assertNodeVersion();
assertPlanningDocsNotStaged();
assertScriptsAreTypeScript();
runConfiguredChecks(readPackageJson());
