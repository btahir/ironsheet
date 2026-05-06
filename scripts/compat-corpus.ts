#!/usr/bin/env tsx
import { existsSync, mkdirSync } from "node:fs";
import { readFile, writeFile } from "node:fs/promises";
import { basename, dirname, resolve } from "node:path";
import process from "node:process";
import {
  hasFailingChecks,
  parseCompatibilityFixtureManifest,
  requiredValidatorsPassed,
  type CompatibilityFixture,
  type CompatibilityReport
} from "../packages/compat/src/index.ts";
import { runCompatibilityChecks, writeReport } from "./compat.ts";

type CorpusFixtureResult = {
  id: string;
  workbookPath: string;
  status: "pass" | "fail" | "skip";
  message: string;
  reportPath?: string;
  report?: CompatibilityReport;
};

type CorpusReport = {
  schemaVersion: 1;
  generatedAt: string;
  manifestPath: string;
  strict: boolean;
  summary: {
    passed: number;
    failed: number;
    skipped: number;
    total: number;
  };
  fixtures: CorpusFixtureResult[];
};

async function runCorpus(
  manifestPath: string,
  options: { strict: boolean }
): Promise<CorpusReport> {
  const absoluteManifestPath = resolve(manifestPath);
  const manifest = parseCompatibilityFixtureManifest(
    JSON.parse(await readFile(absoluteManifestPath, "utf8")) as unknown
  );
  const manifestDir = dirname(absoluteManifestPath);
  const fixtures: CorpusFixtureResult[] = [];

  for (const fixture of manifest.fixtures) {
    fixtures.push(await runFixture(manifestDir, fixture));
  }

  return {
    schemaVersion: 1,
    generatedAt: new Date().toISOString(),
    manifestPath: absoluteManifestPath,
    strict: options.strict,
    summary: {
      passed: fixtures.filter((fixture) => fixture.status === "pass").length,
      failed: fixtures.filter((fixture) => fixture.status === "fail").length,
      skipped: fixtures.filter((fixture) => fixture.status === "skip").length,
      total: fixtures.length
    },
    fixtures
  };
}

async function runFixture(
  manifestDir: string,
  fixture: CompatibilityFixture
): Promise<CorpusFixtureResult> {
  const workbookPath = resolve(manifestDir, fixture.path);

  if (fixture.status === "pending") {
    return {
      id: fixture.id,
      workbookPath,
      status: "skip",
      message: "Fixture is documented but not active yet"
    };
  }

  if (!existsSync(workbookPath)) {
    return {
      id: fixture.id,
      workbookPath,
      status: "fail",
      message: "Active fixture workbook is missing"
    };
  }

  const report = await runCompatibilityChecks(workbookPath);
  const reportPath = writeReport(report);
  if (hasFailingChecks(report)) {
    return {
      id: fixture.id,
      workbookPath,
      status: "fail",
      message: "One or more compatibility checks failed",
      reportPath,
      report
    };
  }

  if (!requiredValidatorsPassed(report, fixture.requiredValidators)) {
    return {
      id: fixture.id,
      workbookPath,
      status: "fail",
      message: `Required validator did not pass: ${fixture.requiredValidators.join(", ")}`,
      reportPath,
      report
    };
  }

  return {
    id: fixture.id,
    workbookPath,
    status: "pass",
    message: "Compatibility fixture passed",
    reportPath,
    report
  };
}

async function writeCorpusReport(report: CorpusReport): Promise<string> {
  const outputDir = resolve("compat-output");
  mkdirSync(outputDir, { recursive: true });
  const outputPath = resolve(outputDir, `${basename(report.manifestPath)}.compat-corpus.json`);
  await writeFile(outputPath, `${JSON.stringify(report, null, 2)}\n`);
  console.log(`compat: wrote ${outputPath}`);
  return outputPath;
}

const args = process.argv.slice(2);
const manifestPath = args.find((arg) => arg !== "--strict") ?? "fixtures/corpus/manifest.json";
const strict = args.includes("--strict") || process.env.IRONSHEET_STRICT_CORPUS === "1";
const report = await runCorpus(manifestPath, { strict });
console.log(JSON.stringify(report, null, 2));
await writeCorpusReport(report);

if (report.summary.failed > 0 || (report.strict && report.summary.skipped > 0)) {
  process.exit(1);
}
