#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import process from "node:process";

const blockedPaths = new Set(["IRONSHEET_SPEC.md"]);
const optionalChecks = ["format:check", "lint", "typecheck", "test", "build"];

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    stdio: "inherit",
    shell: false,
    ...options
  });

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

function output(command, args) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    shell: false
  });

  if (result.status !== 0) {
    return "";
  }

  return result.stdout.trim();
}

function fail(message) {
  console.error(`verify: ${message}`);
  process.exit(1);
}

const nodeMajor = Number.parseInt(process.versions.node.split(".")[0] ?? "0", 10);
if (nodeMajor < 18) {
  fail(`Node.js >=18 is required; found ${process.versions.node}`);
}

const insideGit = output("git", ["rev-parse", "--is-inside-work-tree"]);
if (insideGit === "true") {
  const staged = output("git", ["diff", "--cached", "--name-only"])
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const blocked = staged.filter((file) => blockedPaths.has(file));
  if (blocked.length > 0) {
    fail(`planning docs must not be committed: ${blocked.join(", ")}`);
  }
}

if (!existsSync("package.json")) {
  fail("package.json not found");
}

const packageJson = JSON.parse(readFileSync("package.json", "utf8"));
const scripts = packageJson.scripts ?? {};
const checksToRun = optionalChecks.filter((name) => Object.hasOwn(scripts, name));

if (checksToRun.length === 0) {
  console.log("verify: no optional checks configured yet; repository guard checks passed");
  process.exit(0);
}

for (const check of checksToRun) {
  console.log(`verify: running npm run ${check}`);
  run("npm", ["run", check]);
}

console.log("verify: all checks passed");

