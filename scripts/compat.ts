#!/usr/bin/env tsx
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, realpathSync, writeFileSync } from "node:fs";
import { basename, resolve } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import {
  createCompatibilityReport,
  hasFailingChecks,
  type CompatibilityCheck,
  type CompatibilityReport
} from "../packages/compat/src/index.ts";
import { validateWorkbookFile } from "../packages/node/src/index.ts";

type CommandResult = {
  status: number | null;
  stdout: string;
  stderr: string;
};

function run(command: string, args: string[]): CommandResult {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    shell: false
  });

  return {
    status: result.status,
    stdout: result.stdout.trim(),
    stderr: result.stderr.trim()
  };
}

function commandExists(command: string): boolean {
  return run("sh", ["-lc", `command -v ${command} >/dev/null 2>&1`]).status === 0;
}

function isMacOs(): boolean {
  return process.platform === "darwin";
}

function checkFile(workbookPath: string): CompatibilityCheck {
  if (!existsSync(workbookPath)) {
    return {
      validator: "file",
      status: "fail",
      message: "Workbook file does not exist"
    };
  }

  return {
    validator: "file",
    status: "pass",
    message: "Workbook file exists"
  };
}

function checkZip(workbookPath: string): CompatibilityCheck {
  if (!commandExists("unzip")) {
    return {
      validator: "zip",
      status: "skip",
      message: "`unzip` is not available"
    };
  }

  const result = run("unzip", ["-t", workbookPath]);
  if (result.status !== 0) {
    return {
      validator: "zip",
      status: "fail",
      message: "ZIP package integrity check failed",
      details: { stderr: result.stderr, stdout: result.stdout }
    };
  }

  return {
    validator: "zip",
    status: "pass",
    message: "ZIP package integrity check passed"
  };
}

async function checkIronsheetValidation(workbookPath: string): Promise<CompatibilityCheck> {
  try {
    const report = await validateWorkbookFile(workbookPath);
    if (report.summary.errors > 0) {
      return {
        validator: "ironsheet",
        status: "fail",
        message: "Ironsheet semantic validation found workbook errors",
        details: {
          summary: report.summary,
          issues: report.issues
        }
      };
    }

    return {
      validator: "ironsheet",
      status: "pass",
      message:
        report.summary.warnings === 0
          ? "Ironsheet semantic validation passed"
          : "Ironsheet semantic validation passed with warnings",
      details: {
        summary: report.summary,
        issues: report.issues
      }
    };
  } catch (error) {
    return {
      validator: "ironsheet",
      status: "fail",
      message: "Ironsheet semantic validation could not read the workbook",
      details: {
        error: error instanceof Error ? error.message : String(error)
      }
    };
  }
}

function checkNumbers(workbookPath: string): CompatibilityCheck {
  if (!isMacOs()) {
    return {
      validator: "numbers",
      status: "skip",
      message: "Numbers check only runs on macOS"
    };
  }

  const appLookup = run("mdfind", ["kMDItemCFBundleIdentifier == 'com.apple.iWork.Numbers'"]);
  if (appLookup.status !== 0 || appLookup.stdout.length === 0) {
    return {
      validator: "numbers",
      status: "skip",
      message: "Numbers.app is not installed"
    };
  }

  if (process.env.IRONSHEET_RUN_NUMBERS !== "1") {
    return {
      validator: "numbers",
      status: "manual",
      message:
        "Numbers.app is available. Set IRONSHEET_RUN_NUMBERS=1 to open the workbook for an interactive import smoke check."
    };
  }

  const script = `
set workbookPath to POSIX file "${workbookPath.replaceAll('"', '\\"')}"
tell application "Numbers"
  activate
  open workbookPath
end tell
`;

  const result = run("osascript", ["-e", script]);
  if (result.status !== 0) {
    return {
      validator: "numbers",
      status: "fail",
      message: "Numbers.app failed to open the workbook",
      details: { stderr: result.stderr, stdout: result.stdout }
    };
  }

  return {
    validator: "numbers",
    status: "manual",
    message:
      "Workbook opened in Numbers.app. Use Computer Use or manual inspection to confirm import quality and close the document."
  };
}

function checkLibreOffice(workbookPath: string): CompatibilityCheck {
  const soffice = commandExists("soffice") ? "soffice" : "";
  if (soffice.length === 0) {
    return {
      validator: "libreoffice",
      status: "skip",
      message: "LibreOffice `soffice` is not installed"
    };
  }

  if (process.env.IRONSHEET_RUN_LIBREOFFICE !== "1") {
    return {
      validator: "libreoffice",
      status: "manual",
      message:
        "LibreOffice is available. Set IRONSHEET_RUN_LIBREOFFICE=1 to run headless import/export."
    };
  }

  const outputDir = resolve("compat-output", "libreoffice");
  mkdirSync(outputDir, { recursive: true });

  const result = run(soffice, [
    "--headless",
    "--convert-to",
    "xlsx",
    "--outdir",
    outputDir,
    workbookPath
  ]);

  if (result.status !== 0) {
    return {
      validator: "libreoffice",
      status: "fail",
      message: "LibreOffice headless import/export failed",
      details: { stderr: result.stderr, stdout: result.stdout }
    };
  }

  return {
    validator: "libreoffice",
    status: "pass",
    message: "LibreOffice headless import/export completed",
    details: { outputDir }
  };
}

function checkOpenXmlSdk(workbookPath: string): CompatibilityCheck {
  if (!commandExists("dotnet")) {
    return {
      validator: "openxml-sdk",
      status: "skip",
      message: ".NET SDK is not installed; Open XML SDK validation is unavailable"
    };
  }

  const projectPath = resolve("tools/openxml-validator/OpenXmlValidator.csproj");
  if (!existsSync(projectPath)) {
    return {
      validator: "openxml-sdk",
      status: "manual",
      message:
        ".NET SDK is available, but tools/openxml-validator/OpenXmlValidator.csproj is not scaffolded yet"
    };
  }

  if (process.env.IRONSHEET_RUN_OPENXML_SDK !== "1") {
    return {
      validator: "openxml-sdk",
      status: "manual",
      message:
        "Open XML SDK validator harness is available. Set IRONSHEET_RUN_OPENXML_SDK=1 to run it."
    };
  }

  const result = run("dotnet", ["run", "--project", projectPath, "--", workbookPath]);
  if (result.status !== 0) {
    return {
      validator: "openxml-sdk",
      status: "fail",
      message: "Open XML SDK validation failed",
      details: { stderr: result.stderr, stdout: result.stdout }
    };
  }

  return {
    validator: "openxml-sdk",
    status: "pass",
    message: "Open XML SDK validation passed",
    details: { stdout: result.stdout }
  };
}

function checkExcel(): CompatibilityCheck {
  if (!isMacOs()) {
    return {
      validator: "excel",
      status: "skip",
      message: "Local Excel check is only configured for macOS right now"
    };
  }

  const appLookup = run("mdfind", ["kMDItemCFBundleIdentifier == 'com.microsoft.Excel'"]);
  if (appLookup.status !== 0 || appLookup.stdout.length === 0) {
    return {
      validator: "excel",
      status: "skip",
      message: "Microsoft Excel is not installed"
    };
  }

  return {
    validator: "excel",
    status: "manual",
    message:
      "Microsoft Excel is installed. Add an Excel-specific smoke validator before enabling this check."
  };
}

export async function runCompatibilityChecks(workbookPath: string): Promise<CompatibilityReport> {
  const absolutePath = existsSync(workbookPath)
    ? realpathSync(workbookPath)
    : resolve(workbookPath);
  const checks = [
    checkFile(absolutePath),
    checkZip(absolutePath),
    await checkIronsheetValidation(absolutePath),
    checkNumbers(absolutePath),
    checkLibreOffice(absolutePath),
    checkOpenXmlSdk(absolutePath),
    checkExcel()
  ];

  return createCompatibilityReport(absolutePath, checks);
}

export function writeReport(report: CompatibilityReport): string {
  const outputDir = resolve("compat-output");
  mkdirSync(outputDir, { recursive: true });

  const outputPath = resolve(outputDir, `${basename(report.workbookPath)}.compat.json`);
  writeFileSync(outputPath, `${JSON.stringify(report, null, 2)}\n`);
  console.log(`compat: wrote ${outputPath}`);
  return outputPath;
}

async function main(): Promise<void> {
  const workbookPath = process.argv[2];
  if (workbookPath === undefined) {
    console.error("usage: npm run compat:check -- path/to/workbook.xlsx");
    process.exit(2);
  }

  const resolvedPath = resolve(workbookPath);
  const report = await runCompatibilityChecks(resolvedPath);
  console.log(JSON.stringify(report, null, 2));
  writeReport(report);

  if (hasFailingChecks(report)) {
    process.exit(1);
  }
}

if (process.argv[1] !== undefined && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  await main();
}
