#!/usr/bin/env tsx
import { spawnSync } from "node:child_process";
import process from "node:process";

function run(command: string, args: string[]): void {
  const result = spawnSync(command, args, {
    stdio: "inherit",
    shell: false
  });

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

function quiet(command: string, args: string[]): number {
  const result = spawnSync(command, args, {
    stdio: "ignore",
    shell: false
  });

  return result.status ?? 1;
}

const message = process.argv.slice(2).join(" ").trim();

if (message.length === 0) {
  console.error('usage: npm run commit:safe -- "commit message"');
  process.exit(2);
}

if (quiet("git", ["rev-parse", "--is-inside-work-tree"]) !== 0) {
  console.error("commit:safe: not inside a git repository");
  process.exit(1);
}

run("git", ["add", "-A"]);

// The planning spec is useful locally, but should not be part of normal product commits.
quiet("git", ["reset", "--", "IRONSHEET_SPEC.md"]);

if (quiet("git", ["diff", "--cached", "--quiet"]) === 0) {
  console.log("commit:safe: no staged changes to commit");
  process.exit(0);
}

run("npm", ["run", "verify"]);
run("git", ["commit", "-m", message]);
