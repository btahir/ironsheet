#!/usr/bin/env tsx
import { chmodSync, rmSync } from "node:fs";
import { join } from "node:path";
import process from "node:process";
import { spawnSync } from "node:child_process";

const packages = ["core", "compat", "browser", "node", "cli"] as const;

for (const packageName of packages) {
  rmSync(join("packages", packageName, "dist"), { force: true, recursive: true });
}

for (const packageName of packages) {
  const config = join("packages", packageName, "tsconfig.build.json");
  console.log(`build: ${packageName}`);
  const result = spawnSync(process.execPath, ["node_modules/typescript/bin/tsc", "-p", config], {
    stdio: "inherit",
    shell: false
  });

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

if (process.platform !== "win32") {
  chmodSync(join("packages", "cli", "dist", "cli.js"), 0o755);
}
