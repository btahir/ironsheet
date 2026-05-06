import assert from "node:assert/strict";
import test from "node:test";
import { createCompatibilityReport, hasFailingChecks } from "../packages/compat/src/index.ts";

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
