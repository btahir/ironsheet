export type CompatibilityStatus = "pass" | "fail" | "skip" | "manual";

export type CompatibilityValidator =
  | "file"
  | "zip"
  | "numbers"
  | "libreoffice"
  | "openxml-sdk"
  | "excel";

export type CompatibilityCheck = {
  validator: CompatibilityValidator;
  status: CompatibilityStatus;
  message: string;
  details?: Record<string, unknown>;
};

export type CompatibilityReport = {
  schemaVersion: 1;
  generatedAt: string;
  workbookPath: string;
  checks: CompatibilityCheck[];
};

export function createCompatibilityReport(
  workbookPath: string,
  checks: CompatibilityCheck[],
  generatedAt = new Date().toISOString()
): CompatibilityReport {
  return {
    schemaVersion: 1,
    generatedAt,
    workbookPath,
    checks
  };
}

export function hasFailingChecks(report: CompatibilityReport): boolean {
  return report.checks.some((check) => check.status === "fail");
}
