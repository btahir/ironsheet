import assert from "node:assert/strict";
import test from "node:test";
import { IronsheetError, type Workbook } from "../packages/core/src/index.ts";
import { openWorkbook } from "../packages/node/src/index.ts";
import { createMinimalWorkbook, type MinimalWorkbookOptions } from "./helpers/minimal-xlsx.ts";

const iterations = 50;
const mutationsPerIteration = 8;
const baseSeed = 0x1205eed;

type Rng = () => number;

function mulberry32(seed: number): Rng {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function pick<T>(rng: Rng, values: T[]): T {
  const value = values[Math.floor(rng() * values.length)];
  if (value === undefined) {
    throw new Error("pick from empty list");
  }

  return value;
}

function randomInt(rng: Rng, min: number, max: number): number {
  return min + Math.floor(rng() * (max - min + 1));
}

function randomFixtureOptions(rng: Rng): MinimalWorkbookOptions {
  return {
    ...(rng() < 0.4 ? { includeCalcChain: true, includeFormulaCell: true } : {}),
    ...(rng() < 0.3 ? { includeConditionalFormatting: true } : {}),
    ...(rng() < 0.3 ? { includeDataValidation: true } : {}),
    ...(rng() < 0.4 ? { includeDefinedName: true } : {}),
    ...(rng() < 0.3 ? { includeHiddenSheet: true } : {}),
    ...(rng() < 0.3 ? { includeHyperlink: true } : {}),
    ...(rng() < 0.3 ? { includeMerge: true } : {}),
    ...(rng() < 0.4
      ? {
          includeTable: true,
          tableRows: [
            ["Old", 1],
            ["New", 2]
          ] as Array<[string, number]>
        }
      : {}),
    ...(rng() < 0.3 ? { useSharedStrings: true } : {})
  };
}

function randomCellValue(rng: Rng): string | number | boolean | null {
  const kind = randomInt(rng, 0, 4);
  if (kind === 0) {
    return `text-${randomInt(rng, 0, 999)}`;
  }
  if (kind === 1) {
    return randomInt(rng, -100_000, 100_000) / 100;
  }
  if (kind === 2) {
    return rng() < 0.5;
  }
  if (kind === 3) {
    return null;
  }
  return `<&"' special ${randomInt(rng, 0, 99)}`;
}

function randomAddress(rng: Rng): string {
  const column = String.fromCharCode(65 + randomInt(rng, 0, 7));
  return `${column}${randomInt(rng, 12, 60)}`;
}

type Mutation = (workbook: Workbook, rng: Rng) => Promise<void>;

const mutations: Array<{ name: string; run: Mutation }> = [
  {
    name: "patchCell",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      await workbook.patchCell(sheet, randomAddress(rng), randomCellValue(rng));
    }
  },
  {
    name: "patchRange",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      await workbook.patchRange(sheet, randomAddress(rng), [
        [randomCellValue(rng), randomCellValue(rng)],
        [randomCellValue(rng), randomCellValue(rng)]
      ]);
    }
  },
  {
    name: "patchFormula",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      await workbook.patchCell(sheet, randomAddress(rng), {
        formula: `SUM(A1:A${randomInt(rng, 2, 9)})`,
        result: 0
      });
    }
  },
  {
    name: "styleCell",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      await workbook.styleCell(sheet, randomAddress(rng), {
        ...(rng() < 0.5 ? { font: { bold: true, size: randomInt(rng, 8, 16) } } : {}),
        ...(rng() < 0.5 ? { fill: "1F4E79" } : {}),
        ...(rng() < 0.4 ? { border: { all: { style: "thin" } } } : {}),
        ...(rng() < 0.4 ? { numberFormat: "0.00%" } : {}),
        ...(rng() < 0.4 ? { alignment: { horizontal: "center", wrapText: true } } : {})
      });
    }
  },
  {
    name: "styleRange",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      const row = randomInt(rng, 12, 40);
      await workbook.styleRange(sheet, `A${row}:C${row + 2}`, {
        fill: "EEF2F7",
        font: { italic: true }
      });
    }
  },
  {
    name: "appendRows",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      await workbook.appendRows(sheet, [[randomCellValue(rng), randomCellValue(rng)]]);
    }
  },
  {
    name: "clearRange",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      const row = randomInt(rng, 12, 50);
      await workbook.clearRange(sheet, `A${row}:D${row + 1}`, {
        keepStyles: rng() < 0.5
      });
    }
  },
  {
    name: "insertRows",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      await workbook.insertRows(sheet, randomInt(rng, 12, 30), randomInt(rng, 1, 3));
    }
  },
  {
    name: "deleteRows",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      await workbook.deleteRows(sheet, randomInt(rng, 12, 30), randomInt(rng, 1, 3));
    }
  },
  {
    name: "addSheet",
    run: async (workbook, rng) => {
      await workbook.addSheet(`Fuzz${randomInt(rng, 0, 9999)}`);
    }
  },
  {
    name: "copySheet",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      await workbook.copySheet(sheet, `Copy${randomInt(rng, 0, 9999)}`);
    }
  },
  {
    name: "renameSheet",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      await workbook.renameSheet(sheet, `Renamed${randomInt(rng, 0, 9999)}`);
    }
  },
  {
    name: "deleteSheet",
    run: async (workbook, rng) => {
      const visible = workbook.sheets().filter((sheet) => sheet.state === undefined);
      if (visible.length < 2) {
        return;
      }
      await workbook.deleteSheet(pick(rng, visible).name);
    }
  },
  {
    name: "setDefinedName",
    run: async (workbook, rng) => {
      const sheet = pick(rng, workbook.sheets()).name;
      const quoted = sheet.includes(" ") ? `'${sheet}'` : sheet;
      await workbook.setDefinedName(`FuzzName${randomInt(rng, 0, 999)}`, `${quoted}!$A$1`);
    }
  }
];

test("fuzz: random mutation sequences keep workbooks valid across round trips", async () => {
  for (let iteration = 0; iteration < iterations; iteration += 1) {
    const seed = baseSeed + iteration * 7919;
    const rng = mulberry32(seed);
    const applied: string[] = [];

    try {
      const workbook = await openWorkbook(await createMinimalWorkbook(randomFixtureOptions(rng)));

      for (let step = 0; step < mutationsPerIteration; step += 1) {
        const mutation = pick(rng, mutations);
        applied.push(mutation.name);
        try {
          await mutation.run(workbook, rng);
        } catch (error) {
          if (!(error instanceof IronsheetError)) {
            throw error;
          }
          applied[applied.length - 1] = `${mutation.name}(refused)`;
        }
      }

      const validation = await workbook.validate();
      assert.equal(
        validation.summary.errors,
        0,
        JSON.stringify(validation.issues.filter((issue) => issue.severity === "error"))
      );

      const output = await workbook.write();
      const reopened = await openWorkbook(output);
      const revalidation = await reopened.validate();
      assert.equal(
        revalidation.summary.errors,
        0,
        JSON.stringify(revalidation.issues.filter((issue) => issue.severity === "error"))
      );
    } catch (error) {
      throw new Error(
        `fuzz iteration ${iteration} failed (seed ${seed}, mutations: ${applied.join(" -> ")})`,
        { cause: error }
      );
    }
  }
});
