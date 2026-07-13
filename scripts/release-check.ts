#!/usr/bin/env tsx
import { spawnSync } from "node:child_process";
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { join, resolve } from "node:path";
import process from "node:process";

type PackageJson = {
  name?: string;
  version?: string;
  private?: boolean;
  license?: string;
  repository?: string | { type?: string; url?: string; directory?: string };
  main?: string;
  types?: string;
  exports?: Record<string, unknown>;
  files?: string[];
  bin?: Record<string, string>;
  dependencies?: Record<string, string>;
};

type PackageLock = {
  packages?: Record<
    string,
    {
      name?: string;
      version?: string;
      dependencies?: Record<string, string>;
    }
  >;
};

type Capability = {
  name: string;
  status: "available" | "manual" | "missing";
  message: string;
};

const releaseMode = process.argv.includes("--release");
const publishOrder = [
  "@ironsheet/core",
  "@ironsheet/node",
  "@ironsheet/browser",
  "@ironsheet/compat",
  "@ironsheet/cli"
] as const;
const packageDirs = readdirSync("packages", { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .map((entry) => join("packages", entry.name))
  .sort();

function run(command: string, args: string[]): void {
  console.log(`release: running ${command} ${args.join(" ")}`);
  const result = spawnSync(command, args, {
    stdio: "inherit",
    shell: false
  });

  if (result.status !== 0) {
    console.error(`release: command failed: ${command} ${args.join(" ")}`);
    process.exit(result.status ?? 1);
  }
}

function output(command: string, args: string[]): string {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    shell: false
  });

  if (result.status !== 0) {
    return "";
  }

  return result.stdout.trim();
}

function commandExists(command: string): boolean {
  return (
    spawnSync("sh", ["-lc", `command -v ${command} >/dev/null 2>&1`], {
      stdio: "ignore",
      shell: false
    }).status === 0
  );
}

function fail(message: string): never {
  console.error(`release: ${message}`);
  process.exit(1);
}

function readJson<T>(path: string): T {
  return JSON.parse(readFileSync(path, "utf8")) as T;
}

function assertFile(path: string, message: string): void {
  if (!existsSync(path) || !statSync(path).isFile()) {
    fail(message);
  }
}

function assertDirectory(path: string, message: string): void {
  if (!existsSync(path) || !statSync(path).isDirectory()) {
    fail(message);
  }
}

function assertWorkspaceMetadata(): void {
  assertFile("LICENSE", "LICENSE file is missing at the repository root");
  const packageLock = readJson<PackageLock>("package-lock.json");

  const packages = packageDirs.map((dir) => ({
    dir,
    json: readJson<PackageJson>(join(dir, "package.json"))
  }));
  const workspaceVersions = new Map(packages.map(({ json }) => [json.name, json.version] as const));

  for (const { dir, json } of packages) {
    assertFile(join(dir, "README.md"), `${dir}/README.md is missing`);
    assertFile(join(dir, "LICENSE"), `${dir}/LICENSE is missing`);

    if (json.name === undefined || !json.name.startsWith("@ironsheet/")) {
      fail(`${dir}/package.json must use an @ironsheet/* package name`);
    }

    if (json.private === true) {
      fail(`${json.name} must not be private when preparing publishable packages`);
    }

    if (json.license !== "Apache-2.0") {
      fail(`${json.name} must be licensed Apache-2.0, found ${json.license ?? "no license field"}`);
    }

    if (json.repository === undefined) {
      fail(`${json.name} must define a repository field`);
    }

    if (json.version === undefined || json.version.length === 0) {
      fail(`${json.name} must define a version`);
    }

    if (releaseMode && json.version === "0.0.0") {
      fail(`${json.name} still has placeholder version 0.0.0`);
    }

    const lockedPackage = packageLock.packages?.[dir];
    if (lockedPackage?.name !== json.name || lockedPackage.version !== json.version) {
      fail(
        `${json.name} package-lock metadata is stale; expected ${json.version}, found ${lockedPackage?.version ?? "no version"}`
      );
    }

    if (json.main === undefined || json.types === undefined) {
      fail(`${json.name} must define main and types`);
    }

    assertFile(join(dir, json.main), `${json.name} main output is missing: ${json.main}`);
    assertFile(join(dir, json.types), `${json.name} declaration output is missing: ${json.types}`);

    if (json.exports?.["."] === undefined) {
      fail(`${json.name} must export its package root`);
    }

    if (json.files?.includes("dist") !== true) {
      fail(`${json.name} package files must include dist`);
    }

    assertDirectory(join(dir, "dist"), `${json.name} dist directory is missing`);

    for (const [binName, binPath] of Object.entries(json.bin ?? {})) {
      assertFile(join(dir, binPath), `${json.name} bin ${binName} is missing: ${binPath}`);
    }

    for (const [dependency, version] of Object.entries(json.dependencies ?? {})) {
      if (!dependency.startsWith("@ironsheet/")) {
        continue;
      }

      const workspaceVersion = workspaceVersions.get(dependency);
      if (workspaceVersion === undefined) {
        fail(`${json.name} depends on unknown workspace package ${dependency}`);
      }

      if (version !== workspaceVersion) {
        fail(
          `${json.name} depends on ${dependency}@${version}, expected workspace version ${workspaceVersion}`
        );
      }

      if (lockedPackage.dependencies?.[dependency] !== version) {
        fail(
          `${json.name} package-lock dependency ${dependency}@${lockedPackage.dependencies?.[dependency] ?? "missing"}, expected ${version}`
        );
      }
    }
  }
}

function assertNpmAuthentication(): void {
  const identity = output("npm", ["whoami"]);
  if (identity.length === 0) {
    fail("npm authentication is required for a release; run npm login and retry");
  }

  console.log(`release: npm identity: ${identity}`);
}

function validatorCapabilities(): Capability[] {
  const capabilities: Capability[] = [];
  const isMacOs = process.platform === "darwin";
  const numbersInstalled =
    isMacOs &&
    output("mdfind", ["kMDItemCFBundleIdentifier == 'com.apple.iWork.Numbers'"]).length > 0;
  const excelInstalled =
    isMacOs && output("mdfind", ["kMDItemCFBundleIdentifier == 'com.microsoft.Excel'"]).length > 0;
  const hasLibreOffice = commandExists("soffice");
  const hasDotnet = commandExists("dotnet");
  const openXmlProject = resolve("tools/openxml-validator/OpenXmlValidator.csproj");

  capabilities.push({
    name: "Numbers",
    status: numbersInstalled ? "manual" : "missing",
    message: numbersInstalled
      ? "Numbers.app is available for manual import smoke checks"
      : "Numbers.app is not installed"
  });
  capabilities.push({
    name: "LibreOffice",
    status: hasLibreOffice ? "available" : "missing",
    message: hasLibreOffice
      ? "soffice is available for optional headless import/export checks"
      : "LibreOffice soffice is not installed"
  });
  capabilities.push({
    name: "Open XML SDK",
    status: hasDotnet && existsSync(openXmlProject) ? "available" : "missing",
    message:
      hasDotnet && existsSync(openXmlProject)
        ? "dotnet and the Open XML SDK validator harness are available"
        : hasDotnet
          ? "dotnet is installed, but tools/openxml-validator/OpenXmlValidator.csproj is missing"
          : ".NET SDK is not installed"
  });
  capabilities.push({
    name: "Excel",
    status: excelInstalled ? "manual" : "missing",
    message: excelInstalled
      ? "Microsoft Excel is available for manual release smoke checks"
      : "Microsoft Excel is not installed"
  });

  return capabilities;
}

if (releaseMode) {
  run("npm", ["run", "verify"]);
  run("npm", ["run", "compat:corpus:strict"]);
} else {
  run("npm", ["run", "ci"]);
}
run("npm", ["run", "pack:check"]);
assertWorkspaceMetadata();

if (releaseMode) {
  assertNpmAuthentication();
}

console.log("release: validator capability report");
for (const capability of validatorCapabilities()) {
  console.log(`release: ${capability.name}: ${capability.status} - ${capability.message}`);
}

for (const packageName of publishOrder) {
  run("npm", ["publish", "--workspace", packageName, "--dry-run", "--access", "public"]);
}
console.log(releaseMode ? "release: strict preflight passed" : "release: preflight passed");
