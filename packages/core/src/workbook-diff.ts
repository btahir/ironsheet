import type { Workbook } from "./workbook.ts";
import type { CellPrimitive, ReadCellResult } from "./worksheet.ts";

export type WorkbookCellDiffKind = "added" | "changed" | "removed";

export type WorkbookCellDiffSide = {
  formula?: string;
  styleId?: string;
  value: CellPrimitive;
};

export type WorkbookCellDiff = {
  address: string;
  after?: WorkbookCellDiffSide;
  before?: WorkbookCellDiffSide;
  changed?: Array<"formula" | "style" | "value">;
  kind: WorkbookCellDiffKind;
  sheetName: string;
};

export type WorkbookNameListDiff = {
  added: string[];
  changed: string[];
  removed: string[];
};

export type WorkbookSemanticDiff = {
  cells: WorkbookCellDiff[];
  definedNames: WorkbookNameListDiff;
  sheets: { added: string[]; removed: string[] };
  summary: {
    addedCells: number;
    changedCells: number;
    removedCells: number;
    truncated: boolean;
  };
  tables: WorkbookNameListDiff;
};

export type DiffWorkbooksOptions = {
  maxCellDiffs?: number;
};

const defaultMaxCellDiffs = 10_000;

export async function diffWorkbooks(
  before: Workbook,
  after: Workbook,
  options: DiffWorkbooksOptions = {}
): Promise<WorkbookSemanticDiff> {
  const maxCellDiffs = options.maxCellDiffs ?? defaultMaxCellDiffs;
  const beforeSheets = new Map(before.sheets().map((sheet) => [sheet.name, sheet]));
  const afterSheets = new Map(after.sheets().map((sheet) => [sheet.name, sheet]));

  const sheets = {
    added: [...afterSheets.keys()].filter((name) => !beforeSheets.has(name)),
    removed: [...beforeSheets.keys()].filter((name) => !afterSheets.has(name))
  };

  const cells: WorkbookCellDiff[] = [];
  let addedCells = 0;
  let changedCells = 0;
  let removedCells = 0;
  let truncated = false;

  const recordCellDiff = (diff: WorkbookCellDiff): void => {
    if (diff.kind === "added") {
      addedCells += 1;
    } else if (diff.kind === "removed") {
      removedCells += 1;
    } else {
      changedCells += 1;
    }

    if (cells.length < maxCellDiffs) {
      cells.push(diff);
    } else {
      truncated = true;
    }
  };

  for (const sheetName of beforeSheets.keys()) {
    if (!afterSheets.has(sheetName)) {
      continue;
    }

    const beforeCells = cellMap(await before.readSheetCells(sheetName));
    const afterCells = cellMap(await after.readSheetCells(sheetName));

    for (const [address, beforeCell] of beforeCells) {
      const afterCell = afterCells.get(address);
      if (afterCell === undefined) {
        if (!isBlankCell(beforeCell)) {
          recordCellDiff({
            address,
            before: cellSide(beforeCell),
            kind: "removed",
            sheetName
          });
        }
        continue;
      }

      const changed: Array<"formula" | "style" | "value"> = [];
      if (!cellValuesEqual(beforeCell.value, afterCell.value)) {
        changed.push("value");
      }
      if ((beforeCell.formula ?? null) !== (afterCell.formula ?? null)) {
        changed.push("formula");
      }
      if ((beforeCell.styleId ?? "0") !== (afterCell.styleId ?? "0")) {
        changed.push("style");
      }

      if (changed.length > 0) {
        recordCellDiff({
          address,
          after: cellSide(afterCell),
          before: cellSide(beforeCell),
          changed,
          kind: "changed",
          sheetName
        });
      }
    }

    for (const [address, afterCell] of afterCells) {
      if (!beforeCells.has(address) && !isBlankCell(afterCell)) {
        recordCellDiff({
          address,
          after: cellSide(afterCell),
          kind: "added",
          sheetName
        });
      }
    }
  }

  return {
    cells,
    definedNames: await diffDefinedNames(before, after),
    sheets,
    summary: {
      addedCells,
      changedCells,
      removedCells,
      truncated
    },
    tables: await diffTables(before, after)
  };
}

async function diffDefinedNames(before: Workbook, after: Workbook): Promise<WorkbookNameListDiff> {
  const beforeNames = new Map(
    (await before.definedNames()).map((name) => [
      `${name.name}::${name.localSheetId ?? ""}`,
      name.text
    ])
  );
  const afterNames = new Map(
    (await after.definedNames()).map((name) => [
      `${name.name}::${name.localSheetId ?? ""}`,
      name.text
    ])
  );

  return keyedDiff(beforeNames, afterNames);
}

async function diffTables(before: Workbook, after: Workbook): Promise<WorkbookNameListDiff> {
  const tableKey = (table: {
    columns: Array<{ name?: string }>;
    ref: string;
  }): string => `${table.ref}::${table.columns.map((column) => column.name ?? "").join(",")}`;

  const beforeTables = new Map(
    (await before.tables()).map((table) => [table.name, tableKey(table)])
  );
  const afterTables = new Map((await after.tables()).map((table) => [table.name, tableKey(table)]));

  return keyedDiff(beforeTables, afterTables);
}

function keyedDiff(before: Map<string, string>, after: Map<string, string>): WorkbookNameListDiff {
  const added: string[] = [];
  const changed: string[] = [];
  const removed: string[] = [];

  for (const [key, beforeValue] of before) {
    const afterValue = after.get(key);
    if (afterValue === undefined) {
      removed.push(displayKey(key));
    } else if (afterValue !== beforeValue) {
      changed.push(displayKey(key));
    }
  }

  for (const key of after.keys()) {
    if (!before.has(key)) {
      added.push(displayKey(key));
    }
  }

  return { added, changed, removed };
}

function displayKey(key: string): string {
  const separator = key.indexOf("::");
  if (separator === -1) {
    return key;
  }

  const scope = key.slice(separator + 2);
  const name = key.slice(0, separator);
  return scope === "" ? name : `${name} (sheet ${scope})`;
}

function cellMap(cells: ReadCellResult[]): Map<string, ReadCellResult> {
  return new Map(cells.map((cell) => [cell.address, cell]));
}

function cellSide(cell: ReadCellResult): WorkbookCellDiffSide {
  return {
    value: cell.value,
    ...(cell.formula === undefined ? {} : { formula: cell.formula }),
    ...(cell.styleId === undefined ? {} : { styleId: cell.styleId })
  };
}

function isBlankCell(cell: ReadCellResult): boolean {
  return cell.value === null && cell.formula === undefined;
}

function cellValuesEqual(left: CellPrimitive, right: CellPrimitive): boolean {
  if (left instanceof Date || right instanceof Date) {
    return left instanceof Date && right instanceof Date && left.getTime() === right.getTime();
  }

  return left === right;
}
