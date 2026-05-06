export type CompatibilityStatus = "pass" | "fail" | "skip" | "manual";

export type CompatibilityValidator =
  | "file"
  | "ironsheet"
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

export type CompatibilityFixtureStatus = "active" | "pending";

export type CompatibilityFixture = {
  id: string;
  path: string;
  description: string;
  features: string[];
  status: CompatibilityFixtureStatus;
  requiredValidators: CompatibilityValidator[];
};

export type CompatibilityFixtureManifest = {
  schemaVersion: 1;
  fixtures: CompatibilityFixture[];
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

export function parseCompatibilityFixtureManifest(value: unknown): CompatibilityFixtureManifest {
  if (!isRecord(value) || value.schemaVersion !== 1 || !Array.isArray(value.fixtures)) {
    throw new Error("Fixture manifest must have schemaVersion 1 and a fixtures array");
  }

  const seenIds = new Set<string>();
  return {
    schemaVersion: 1,
    fixtures: value.fixtures.map((fixture, index) => {
      if (!isRecord(fixture)) {
        throw new Error(`Fixture ${index} must be an object`);
      }

      const parsed = parseFixture(fixture, index);
      if (seenIds.has(parsed.id)) {
        throw new Error(`Duplicate fixture id ${parsed.id}`);
      }

      seenIds.add(parsed.id);
      return parsed;
    })
  };
}

export function requiredValidatorsPassed(
  report: CompatibilityReport,
  requiredValidators: CompatibilityValidator[]
): boolean {
  return requiredValidators.every((validator) => {
    return report.checks.some((check) => check.validator === validator && check.status === "pass");
  });
}

function parseFixture(fixture: Record<string, unknown>, index: number): CompatibilityFixture {
  const id = requiredString(fixture, "id", index);
  const path = requiredString(fixture, "path", index);
  const description = requiredString(fixture, "description", index);
  const status = requiredFixtureStatus(fixture.status, index);

  if (
    !Array.isArray(fixture.features) ||
    !fixture.features.every((item) => typeof item === "string")
  ) {
    throw new Error(`Fixture ${index} features must be an array of strings`);
  }

  if (
    !Array.isArray(fixture.requiredValidators) ||
    !fixture.requiredValidators.every(isCompatibilityValidator)
  ) {
    throw new Error(`Fixture ${index} requiredValidators contains an unknown validator`);
  }

  return {
    id,
    path,
    description,
    features: fixture.features,
    status,
    requiredValidators: fixture.requiredValidators
  };
}

function requiredString(fixture: Record<string, unknown>, key: string, index: number): string {
  const value = fixture[key];
  if (typeof value !== "string" || value.length === 0) {
    throw new Error(`Fixture ${index} ${key} must be a non-empty string`);
  }

  return value;
}

function requiredFixtureStatus(value: unknown, index: number): CompatibilityFixtureStatus {
  if (value === "active" || value === "pending") {
    return value;
  }

  throw new Error(`Fixture ${index} status must be active or pending`);
}

function isCompatibilityValidator(value: unknown): value is CompatibilityValidator {
  return (
    value === "file" ||
    value === "ironsheet" ||
    value === "zip" ||
    value === "numbers" ||
    value === "libreoffice" ||
    value === "openxml-sdk" ||
    value === "excel"
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
