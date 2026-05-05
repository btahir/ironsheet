import { parseZip } from "./zip/index.ts";

export type PackageDiffStatus = "added" | "removed" | "changed" | "unchanged";

export type PackageEntryDiff = {
  name: string;
  status: PackageDiffStatus;
  before?: PackageEntrySnapshot;
  after?: PackageEntrySnapshot;
};

export type PackageEntrySnapshot = {
  compressionMethod: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
};

export type PackageDiff = {
  summary: Record<PackageDiffStatus, number>;
  entries: PackageEntryDiff[];
};

export function diffZipPackages(beforeData: Uint8Array, afterData: Uint8Array): PackageDiff {
  const before = new Map(parseZip(beforeData).entries.map((entry) => [entry.name, entry]));
  const after = new Map(parseZip(afterData).entries.map((entry) => [entry.name, entry]));
  const names = [...new Set([...before.keys(), ...after.keys()])].sort();
  const entries: PackageEntryDiff[] = names.map((name) => {
    const beforeEntry = before.get(name);
    const afterEntry = after.get(name);

    if (beforeEntry === undefined && afterEntry !== undefined) {
      return {
        name,
        status: "added",
        after: snapshotEntry(afterEntry)
      };
    }

    if (beforeEntry !== undefined && afterEntry === undefined) {
      return {
        name,
        status: "removed",
        before: snapshotEntry(beforeEntry)
      };
    }

    if (beforeEntry === undefined || afterEntry === undefined) {
      throw new Error(`Unexpected package diff state for ${name}`);
    }

    return {
      name,
      status: entriesEqual(beforeEntry, afterEntry) ? "unchanged" : "changed",
      before: snapshotEntry(beforeEntry),
      after: snapshotEntry(afterEntry)
    };
  });

  return {
    summary: summarize(entries),
    entries
  };
}

function snapshotEntry(entry: {
  compressionMethod: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
}): PackageEntrySnapshot {
  return {
    compressionMethod: entry.compressionMethod,
    crc32: entry.crc32,
    compressedSize: entry.compressedSize,
    uncompressedSize: entry.uncompressedSize
  };
}

function entriesEqual(
  before: { crc32: number; compressedData: Uint8Array; compressionMethod: number },
  after: { crc32: number; compressedData: Uint8Array; compressionMethod: number }
): boolean {
  return (
    before.crc32 === after.crc32 &&
    before.compressionMethod === after.compressionMethod &&
    bytesEqual(before.compressedData, after.compressedData)
  );
}

function bytesEqual(before: Uint8Array, after: Uint8Array): boolean {
  if (before.byteLength !== after.byteLength) {
    return false;
  }

  for (let index = 0; index < before.byteLength; index += 1) {
    if (before[index] !== after[index]) {
      return false;
    }
  }

  return true;
}

function summarize(entries: PackageEntryDiff[]): Record<PackageDiffStatus, number> {
  const summary: Record<PackageDiffStatus, number> = {
    added: 0,
    removed: 0,
    changed: 0,
    unchanged: 0
  };

  for (const entry of entries) {
    summary[entry.status] += 1;
  }

  return summary;
}
