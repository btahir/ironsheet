import assert from "node:assert/strict";
import test from "node:test";

test("TypeScript test runner is wired", () => {
  assert.equal("Move fast and break no spreadsheets.".includes("spreadsheets"), true);
});
