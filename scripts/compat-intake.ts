#!/usr/bin/env tsx
import { copyFile, mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import process from "node:process";
import {
  parseCompatibilityFixtureManifest,
  requiredValidatorsPassed,
  type CompatibilityFixture,
  type CompatibilityReport,
  type CompatibilityValidator
} from "../packages/compat/src/index.ts";
import { runCompatibilityChecks, writeReport } from "./compat.ts";

export type FixtureIntakeOptions = {
  activate?: boolean;
  fixtureId: string;
  manifestPath?: string;
  requiredValidators?: CompatibilityValidator[];
  sourcePath: string;
};

export type FixtureIntakeResult = {
  fixture: CompatibilityFixture;
  manifestPath: string;
  report: CompatibilityReport;
  reportPath: string;
  workbookPath: string;
};

type MutableFixtureManifest = {
  schemaVersion: 1;
  fixtures: Array<Record<string, unknown>>;
};

export async function intakeFixture(options: FixtureIntakeOptions): Promise<FixtureIntakeResult> {
  const manifestPath = resolve(options.manifestPath ?? "fixtures/corpus/manifest.json");
  const rawManifest = JSON.parse(await readFile(manifestPath, "utf8")) as MutableFixtureManifest;
  const parsedManifest = parseCompatibilityFixtureManifest(rawManifest);
  const fixtureIndex = parsedManifest.fixtures.findIndex(
    (fixture) => fixture.id === options.fixtureId
  );
  if (fixtureIndex === -1) {
    throw new Error(`Unknown fixture id ${options.fixtureId}`);
  }

  const fixture = parsedManifest.fixtures[fixtureIndex];
  if (fixture === undefined) {
    throw new Error(`Unknown fixture id ${options.fixtureId}`);
  }

  const manifestDir = dirname(manifestPath);
  const workbookPath = resolve(manifestDir, fixture.path);
  const sourcePath = resolve(options.sourcePath);

  const report = await runCompatibilityChecks(sourcePath);
  const reportPath = writeReport(report);
  const requiredValidators = options.requiredValidators ?? fixture.requiredValidators;
  if (!requiredValidatorsPassed(report, requiredValidators)) {
    throw new Error(`Required validator did not pass: ${requiredValidators.join(", ")}`);
  }

  await mkdir(dirname(workbookPath), { recursive: true });
  await copyFile(sourcePath, workbookPath);

  if (options.activate === true || options.requiredValidators !== undefined) {
    const rawFixture = rawManifest.fixtures[fixtureIndex];
    if (rawFixture === undefined) {
      throw new Error(`Manifest fixture ${options.fixtureId} disappeared during intake`);
    }

    if (options.activate === true) {
      rawFixture.status = "active";
    }
    if (options.requiredValidators !== undefined) {
      rawFixture.requiredValidators = requiredValidators;
    }

    parseCompatibilityFixtureManifest(rawManifest);
    await writeFile(manifestPath, `${JSON.stringify(rawManifest, null, 2)}\n`);
  }

  const updatedManifest = parseCompatibilityFixtureManifest(
    JSON.parse(await readFile(manifestPath, "utf8")) as unknown
  );
  const updatedFixture = updatedManifest.fixtures.find(
    (candidate) => candidate.id === options.fixtureId
  );
  if (updatedFixture === undefined) {
    throw new Error(`Unknown fixture id ${options.fixtureId}`);
  }

  return {
    fixture: updatedFixture,
    manifestPath,
    report,
    reportPath,
    workbookPath
  };
}

function parseArgs(args: string[]): FixtureIntakeOptions {
  const positional: string[] = [];
  let activate = false;
  let manifestPath: string | undefined;
  let requiredValidators: CompatibilityValidator[] | undefined;

  for (const arg of args) {
    if (arg === "--activate") {
      activate = true;
      continue;
    }

    if (arg.startsWith("--manifest=")) {
      manifestPath = arg.slice("--manifest=".length);
      continue;
    }

    if (arg.startsWith("--require=")) {
      requiredValidators = arg
        .slice("--require=".length)
        .split(",")
        .map((validator) => parseCompatibilityValidator(validator.trim()));
      continue;
    }

    positional.push(arg);
  }

  const [fixtureId, sourcePath] = positional;
  if (fixtureId === undefined || sourcePath === undefined || positional.length !== 2) {
    throw new Error(
      "usage: npm run compat:intake -- <fixture-id> <source.xlsx|source.xlsm> [--activate] [--manifest=fixtures/corpus/manifest.json] [--require=file,zip,ironsheet]"
    );
  }

  return {
    activate,
    fixtureId,
    ...(manifestPath === undefined ? {} : { manifestPath }),
    ...(requiredValidators === undefined ? {} : { requiredValidators }),
    sourcePath
  };
}

function parseCompatibilityValidator(value: string): CompatibilityValidator {
  if (
    value === "file" ||
    value === "ironsheet" ||
    value === "zip" ||
    value === "numbers" ||
    value === "libreoffice" ||
    value === "openxml-sdk" ||
    value === "excel"
  ) {
    return value;
  }

  throw new Error(`Unknown compatibility validator ${value}`);
}

if (process.argv[1]?.endsWith("compat-intake.ts")) {
  try {
    const result = await intakeFixture(parseArgs(process.argv.slice(2)));
    console.log(JSON.stringify(result, null, 2));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(1);
  }
}
