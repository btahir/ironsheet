import assert from "node:assert/strict";
import type { PackageDiff, PackageDiffStatus } from "../../packages/core/src/index.ts";

type TrackedPackageDiffStatus = Exclude<PackageDiffStatus, "unchanged">;

export type ExpectedPackageDiff = Partial<Record<TrackedPackageDiffStatus, string[]>> & {
  unchanged?: string[];
};

export function assertPackageDiff(diff: PackageDiff, expected: ExpectedPackageDiff): void {
  for (const status of ["added", "removed", "changed", "repacked"] as const) {
    assert.deepEqual(namesWithStatus(diff, status), sorted(expected[status] ?? []), status);
  }

  for (const name of expected.unchanged ?? []) {
    assert.equal(
      diff.entries.find((entry) => entry.name === name)?.status,
      "unchanged",
      `${name} should be unchanged`
    );
  }
}

function namesWithStatus(diff: PackageDiff, status: TrackedPackageDiffStatus): string[] {
  return sorted(diff.entries.filter((entry) => entry.status === status).map((entry) => entry.name));
}

function sorted(values: string[]): string[] {
  return [...values].sort();
}
