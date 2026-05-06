import assert from "node:assert/strict";
import { resolve } from "node:path";
import test from "node:test";
import {
  createCompatibilityReport,
  hasFailingChecks,
  parseCompatibilityFixtureManifest,
  requiredValidatorsPassed
} from "../packages/compat/src/index.ts";
import { runCompatibilityChecks } from "../scripts/compat.ts";

test("compatibility reports detect failing checks", () => {
  const report = createCompatibilityReport("/tmp/workbook.xlsx", [
    {
      validator: "file",
      status: "pass",
      message: "exists"
    },
    {
      validator: "zip",
      status: "fail",
      message: "invalid zip"
    }
  ]);

  assert.equal(hasFailingChecks(report), true);
});

test("compatibility reports allow skipped optional validators", () => {
  const report = createCompatibilityReport("/tmp/workbook.xlsx", [
    {
      validator: "excel",
      status: "skip",
      message: "not installed"
    }
  ]);

  assert.equal(hasFailingChecks(report), false);
});

test("compatibility reports treat ironsheet validation as required", () => {
  const report = createCompatibilityReport("/tmp/workbook.xlsx", [
    {
      validator: "ironsheet",
      status: "fail",
      message: "semantic validation failed"
    }
  ]);

  assert.equal(hasFailingChecks(report), true);
});

test("compatibility fixture manifests validate fixture shape", () => {
  assert.deepEqual(
    parseCompatibilityFixtureManifest({
      schemaVersion: 1,
      fixtures: [
        {
          id: "chart-dashboard",
          path: "workbooks/chart-dashboard.xlsx",
          description: "Workbook with chart formulas",
          features: ["charts", "formulas"],
          status: "pending",
          requiredValidators: ["file", "zip", "ironsheet"]
        }
      ]
    }),
    {
      schemaVersion: 1,
      fixtures: [
        {
          id: "chart-dashboard",
          path: "workbooks/chart-dashboard.xlsx",
          description: "Workbook with chart formulas",
          features: ["charts", "formulas"],
          status: "pending",
          requiredValidators: ["file", "zip", "ironsheet"]
        }
      ]
    }
  );
});

test("compatibility fixture manifests reject duplicate ids", () => {
  assert.throws(
    () =>
      parseCompatibilityFixtureManifest({
        schemaVersion: 1,
        fixtures: [
          {
            id: "same",
            path: "a.xlsx",
            description: "A",
            features: [],
            status: "pending",
            requiredValidators: ["file"]
          },
          {
            id: "same",
            path: "b.xlsx",
            description: "B",
            features: [],
            status: "pending",
            requiredValidators: ["file"]
          }
        ]
      }),
    /Duplicate fixture id/
  );
});

test("required fixture validators must pass", () => {
  const report = createCompatibilityReport("/tmp/workbook.xlsx", [
    {
      validator: "file",
      status: "pass",
      message: "exists"
    },
    {
      validator: "zip",
      status: "skip",
      message: "missing unzip"
    }
  ]);

  assert.equal(requiredValidatorsPassed(report, ["file"]), true);
  assert.equal(requiredValidatorsPassed(report, ["file", "zip"]), false);
});

test("compatibility checks report missing files without throwing", async () => {
  const report = await runCompatibilityChecks(resolve("compat-output/missing-test-workbook.xlsx"));
  const fileCheck = report.checks.find((check) => check.validator === "file");

  assert.equal(fileCheck?.status, "fail");
});
