import { formatCellAddress, numberToColumnLabel, parseCellAddress } from "./address.ts";
import { WorkbookError } from "./errors.ts";
import { rewriteFormulaElements } from "./formula-rewrite.ts";
import {
  renameFormulaStructuredReferenceColumn,
  renameFormulaStructuredReferenceTable
} from "./formula.ts";
import type { OoxmlPackage } from "./opc.ts";
import { resolveRelationshipTarget } from "./opc.ts";
import { patchCell, removeCellsInRange, replaceRowsInRange, type CellInput } from "./worksheet.ts";
import {
  escapeXmlAttribute,
  findElementCloseStart,
  findElementEnd,
  findFirstStartTag,
  findStartTags
} from "./xml.ts";

const tableRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

export type WorkbookTable = {
  name: string;
  displayName: string;
  columns: WorkbookTableColumn[];
  partName: string;
  worksheetPartName: string;
  ref: string;
  totalsRowCount: number;
};

export type WorkbookTableColumn = {
  id?: string;
  name?: string;
  totalsRowFunction?: string;
};

export async function listWorkbookTables(pkg: OoxmlPackage): Promise<WorkbookTable[]> {
  const tables: WorkbookTable[] = [];

  for (const partName of pkg.listParts().filter((part) => /^xl\/tables\/.+\.xml$/.test(part))) {
    const tableXml = await pkg.readText(partName);
    const table = findFirstStartTag(tableXml, "table");
    if (table === undefined) {
      continue;
    }

    const name = table.attributes.name ?? table.attributes.displayName;
    const displayName = table.attributes.displayName ?? name;
    const ref = table.attributes.ref;

    if (name === undefined || displayName === undefined || ref === undefined) {
      throw new WorkbookError(`Table part ${partName} is missing name/displayName/ref`);
    }

    tables.push({
      name,
      displayName,
      columns: parseTableColumns(tableXml),
      partName,
      worksheetPartName: await findWorksheetForTable(pkg, partName),
      ref,
      totalsRowCount: parseTotalsRowCount(table.attributes)
    });
  }

  return tables;
}

export async function findWorkbookTable(
  pkg: OoxmlPackage,
  tableName: string
): Promise<WorkbookTable> {
  for (const table of await listWorkbookTables(pkg)) {
    if (table.name === tableName || table.displayName === tableName) {
      return table;
    }
  }

  throw new WorkbookError(`Unknown table ${tableName}`);
}

export async function replaceTableRows(
  pkg: OoxmlPackage,
  tableName: string,
  rows: CellInput[][]
): Promise<WorkbookTable> {
  const table = await findWorkbookTable(pkg, tableName);
  const range = parseTableRange(table.ref);
  const bodyStartRow = range.start.row + 1;
  const bodyEndRow = range.end.row - table.totalsRowCount;
  const newEndRow = bodyStartRow + rows.length - 1 + table.totalsRowCount;
  const newRef = `${numberToColumnLabel(range.start.column)}${range.start.row}:${numberToColumnLabel(range.end.column)}${Math.max(range.start.row, newEndRow)}`;

  const worksheetXml = await pkg.readText(table.worksheetPartName);
  await assertTableRowsCanResize(pkg, table, worksheetXml, newRef);
  pkg.setText(
    table.worksheetPartName,
    replaceRowsInRange(
      worksheetXml,
      {
        startRow: bodyStartRow,
        endRow: bodyEndRow,
        startColumn: range.start.column,
        endColumn: range.end.column
      },
      rows,
      { trailingRows: table.totalsRowCount }
    )
  );

  const tableXml = await pkg.readText(table.partName);
  pkg.setText(table.partName, updateTableRef(tableXml, newRef));

  return {
    ...table,
    ref: newRef
  };
}

export async function renameWorkbookTable(
  pkg: OoxmlPackage,
  tableName: string,
  nextName: string
): Promise<WorkbookTable> {
  validateTableName(nextName);

  const tables = await listWorkbookTables(pkg);
  const table = tables.find((candidate) => {
    return candidate.name === tableName || candidate.displayName === tableName;
  });

  if (table === undefined) {
    throw new WorkbookError(`Unknown table ${tableName}`);
  }

  const duplicate = tables.find((candidate) => {
    if (candidate.partName === table.partName) {
      return false;
    }

    return (
      candidate.name.toLowerCase() === nextName.toLowerCase() ||
      candidate.displayName.toLowerCase() === nextName.toLowerCase()
    );
  });

  if (duplicate !== undefined) {
    throw new WorkbookError(`Table name ${nextName} is already used by ${duplicate.displayName}`);
  }

  const oldNames = [...new Set([table.name, table.displayName])];
  const tableXml = await pkg.readText(table.partName);
  pkg.setText(table.partName, updateTableIdentity(tableXml, nextName));

  await rewriteTableFormulaReferences(pkg, oldNames, nextName);

  return {
    ...table,
    name: nextName,
    displayName: nextName
  };
}

export async function renameWorkbookTableColumn(
  pkg: OoxmlPackage,
  tableName: string,
  columnName: string,
  nextName: string
): Promise<WorkbookTable> {
  validateTableColumnName(nextName);

  const table = await findWorkbookTable(pkg, tableName);
  const columnIndex = table.columns.findIndex(
    (column) => column.name?.toLowerCase() === columnName.toLowerCase()
  );
  if (columnIndex === -1) {
    throw new WorkbookError(`Unknown column ${columnName} in table ${table.displayName}`);
  }
  const existingColumnName = table.columns[columnIndex]?.name ?? columnName;

  const duplicate = table.columns.find((column, index) => {
    return index !== columnIndex && column.name?.toLowerCase() === nextName.toLowerCase();
  });
  if (duplicate !== undefined) {
    throw new WorkbookError(
      `Column name ${nextName} is already used by table ${table.displayName}`
    );
  }

  const tableXml = await pkg.readText(table.partName);
  pkg.setText(table.partName, updateTableColumnName(tableXml, existingColumnName, nextName));

  const range = parseTableRange(table.ref);
  const headerAddress = formatCellAddress(range.start.column + columnIndex, range.start.row);
  const worksheetXml = await pkg.readText(table.worksheetPartName);
  pkg.setText(table.worksheetPartName, patchCell(worksheetXml, headerAddress, nextName).xml);

  await rewriteTableColumnFormulaReferences(pkg, table, existingColumnName, nextName);

  return {
    ...table,
    columns: table.columns.map((column, index) =>
      index === columnIndex ? { ...column, name: nextName } : column
    )
  };
}

export async function appendWorkbookTableColumn(
  pkg: OoxmlPackage,
  tableName: string,
  columnName: string,
  values: CellInput[] = []
): Promise<WorkbookTable> {
  validateTableColumnName(columnName);

  const table = await findWorkbookTable(pkg, tableName);
  assertTableColumnNameAvailable(table, columnName);

  const range = parseTableRange(table.ref);
  const column = range.end.column + 1;
  const newRef = tableRef(range.start.column, range.start.row, column, range.end.row);
  const bodyStartRow = range.start.row + 1;
  const bodyEndRow = range.end.row - table.totalsRowCount;

  let worksheetXml = await pkg.readText(table.worksheetPartName);
  await assertTableColumnCanAppend(pkg, table, worksheetXml, newRef, column);
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    const value =
      row === range.start.row
        ? columnName
        : row > bodyEndRow
          ? null
          : (values[row - bodyStartRow] ?? null);
    worksheetXml = patchCell(worksheetXml, formatCellAddress(column, row), value).xml;
  }
  pkg.setText(table.worksheetPartName, worksheetXml);

  const tableXml = await pkg.readText(table.partName);
  const nextColumn = {
    id: String(nextTableColumnId(table.columns)),
    name: columnName
  };
  pkg.setText(table.partName, appendTableColumnXml(updateTableRef(tableXml, newRef), nextColumn));

  return {
    ...table,
    ref: newRef,
    columns: [...table.columns, nextColumn]
  };
}

export async function removeRightmostWorkbookTableColumn(
  pkg: OoxmlPackage,
  tableName: string,
  columnName: string
): Promise<WorkbookTable> {
  const table = await findWorkbookTable(pkg, tableName);
  const columnIndex = table.columns.findIndex(
    (column) => column.name?.toLowerCase() === columnName.toLowerCase()
  );
  if (columnIndex === -1) {
    throw new WorkbookError(`Unknown column ${columnName} in table ${table.displayName}`);
  }
  if (table.columns.length <= 1) {
    throw new WorkbookError(`Cannot remove the only column from table ${table.displayName}`);
  }
  if (columnIndex !== table.columns.length - 1) {
    throw new WorkbookError("Only the rightmost table column can be removed safely");
  }

  const range = parseTableRange(table.ref);
  const column = range.start.column + columnIndex;
  const newRef = tableRef(range.start.column, range.start.row, column - 1, range.end.row);
  const worksheetXml = await pkg.readText(table.worksheetPartName);
  await assertTableColumnCanRemove(pkg, table, column);
  pkg.setText(
    table.worksheetPartName,
    removeCellsInRange(worksheetXml, {
      startRow: range.start.row,
      endRow: range.end.row,
      startColumn: column,
      endColumn: column
    }).xml
  );

  const existingColumnName = table.columns[columnIndex]?.name ?? columnName;
  const tableXml = await pkg.readText(table.partName);
  pkg.setText(
    table.partName,
    updateTableRef(removeTableColumnXml(tableXml, existingColumnName), newRef)
  );

  return {
    ...table,
    ref: newRef,
    columns: table.columns.slice(0, -1)
  };
}

function parseTableRange(ref: string): {
  start: { row: number; column: number };
  end: { row: number; column: number };
} {
  const [start, end] = ref.split(":");
  if (start === undefined || end === undefined) {
    throw new WorkbookError(`Invalid table ref ${ref}`);
  }

  const parsedStart = parseCellAddress(start);
  const parsedEnd = parseCellAddress(end);

  return {
    start: { row: parsedStart.row, column: parsedStart.column },
    end: { row: parsedEnd.row, column: parsedEnd.column }
  };
}

async function assertTableRowsCanResize(
  pkg: OoxmlPackage,
  table: WorkbookTable,
  worksheetXml: string,
  newRef: string
): Promise<void> {
  const oldRange = parseTableRange(table.ref);
  const newRange = parseTableRange(newRef);
  if (newRange.end.row > oldRange.end.row) {
    const occupiedRows = worksheetOccupiedRows(worksheetXml);
    for (let row = oldRange.end.row + 1; row <= newRange.end.row; row += 1) {
      if (occupiedRows.has(row)) {
        throw new WorkbookError(
          `Cannot expand table ${table.displayName} through occupied worksheet row ${row}`
        );
      }
    }
  }

  await assertTableRangeDoesNotOverlapOtherTables(pkg, table, newRef);
}

async function assertTableColumnCanAppend(
  pkg: OoxmlPackage,
  table: WorkbookTable,
  worksheetXml: string,
  newRef: string,
  column: number
): Promise<void> {
  const range = parseTableRange(table.ref);
  const occupiedAddress = firstOccupiedCellInRange(worksheetXml, {
    start: { row: range.start.row, column },
    end: { row: range.end.row, column }
  });
  if (occupiedAddress !== undefined) {
    throw new WorkbookError(
      `Cannot append column to table ${table.displayName}; ${occupiedAddress} is already occupied`
    );
  }

  await assertTableRangeDoesNotOverlapOtherTables(pkg, table, newRef);
}

async function assertTableColumnCanRemove(
  pkg: OoxmlPackage,
  table: WorkbookTable,
  column: number
): Promise<void> {
  const range = parseTableRange(table.ref);
  await assertTableRangeDoesNotOverlapOtherTables(
    pkg,
    table,
    tableRef(column, range.start.row, column, range.end.row)
  );
}

async function assertTableRangeDoesNotOverlapOtherTables(
  pkg: OoxmlPackage,
  table: WorkbookTable,
  ref: string
): Promise<void> {
  const range = parseTableRange(ref);
  for (const other of await listWorkbookTables(pkg)) {
    if (other.partName === table.partName || other.worksheetPartName !== table.worksheetPartName) {
      continue;
    }

    if (tableRangesIntersect(range, parseTableRange(other.ref))) {
      throw new WorkbookError(
        `Cannot resize table ${table.displayName}; new range ${ref} overlaps table ${other.displayName} at ${other.ref}`
      );
    }
  }
}

function worksheetOccupiedRows(xml: string): Set<number> {
  const rows = new Set<number>();
  for (const row of findStartTags(xml, "row")) {
    const rowNumber = Number.parseInt(row.attributes.r ?? "", 10);
    if (Number.isInteger(rowNumber) && rowNumber > 0) {
      rows.add(rowNumber);
    }
  }

  for (const cell of findStartTags(xml, "c")) {
    const address = cell.attributes.r;
    if (address === undefined) {
      continue;
    }

    rows.add(parseCellAddress(address).row);
  }

  return rows;
}

function firstOccupiedCellInRange(
  xml: string,
  range: {
    start: { row: number; column: number };
    end: { row: number; column: number };
  }
): string | undefined {
  for (const cell of findStartTags(xml, "c")) {
    const address = cell.attributes.r;
    if (address === undefined) {
      continue;
    }

    const parsed = parseCellAddress(address);
    if (
      parsed.column >= range.start.column &&
      parsed.column <= range.end.column &&
      parsed.row >= range.start.row &&
      parsed.row <= range.end.row
    ) {
      return parsed.address;
    }
  }

  return undefined;
}

function tableRangesIntersect(
  left: {
    start: { row: number; column: number };
    end: { row: number; column: number };
  },
  right: {
    start: { row: number; column: number };
    end: { row: number; column: number };
  }
): boolean {
  return (
    left.start.column <= right.end.column &&
    left.end.column >= right.start.column &&
    left.start.row <= right.end.row &&
    left.end.row >= right.start.row
  );
}

function tableRef(
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number
): string {
  return `${numberToColumnLabel(startColumn)}${startRow}:${numberToColumnLabel(endColumn)}${endRow}`;
}

function parseTotalsRowCount(attributes: Record<string, string>): number {
  const explicitCount = Number.parseInt(attributes.totalsRowCount ?? "", 10);
  if (Number.isInteger(explicitCount) && explicitCount > 0) {
    return explicitCount;
  }

  return attributes.totalsRowShown === "1" ? 1 : 0;
}

function parseTableColumns(xml: string): WorkbookTableColumn[] {
  return findStartTags(xml, "tableColumn").map((column) => ({
    ...(column.attributes.id === undefined ? {} : { id: column.attributes.id }),
    ...(column.attributes.name === undefined ? {} : { name: column.attributes.name }),
    ...(column.attributes.totalsRowFunction === undefined
      ? {}
      : { totalsRowFunction: column.attributes.totalsRowFunction })
  }));
}

async function findWorksheetForTable(pkg: OoxmlPackage, tablePartName: string): Promise<string> {
  for (const partName of pkg.listParts().filter((part) => /^xl\/worksheets\/.+\.xml$/.test(part))) {
    const relationships = await pkg.relationshipsFor(partName);
    const match = relationships.find((relationship) => {
      const target = resolveRelationshipTarget(partName, relationship.target);
      return relationship.type === tableRelationship && target === tablePartName;
    });

    if (match !== undefined) {
      return partName;
    }
  }

  throw new WorkbookError(`No worksheet relationship points to ${tablePartName}`);
}

function updateTableRef(xml: string, ref: string): string {
  const table = findFirstStartTag(xml, "table");
  if (table === undefined) {
    throw new WorkbookError("Table XML is missing table tag");
  }

  let nextXml = `${xml.slice(0, table.start)}${upsertRef(table.raw, ref)}${xml.slice(table.end)}`;
  const autoFilter = findFirstStartTag(nextXml, "autoFilter");
  if (autoFilter !== undefined) {
    nextXml = `${nextXml.slice(0, autoFilter.start)}${upsertRef(autoFilter.raw, ref)}${nextXml.slice(autoFilter.end)}`;
  }

  return nextXml;
}

function updateTableIdentity(xml: string, name: string): string {
  const table = findFirstStartTag(xml, "table");
  if (table === undefined) {
    throw new WorkbookError("Table XML is missing table tag");
  }

  return `${xml.slice(0, table.start)}${upsertAttributes(table.raw, {
    displayName: name,
    name
  })}${xml.slice(table.end)}`;
}

function updateTableColumnName(xml: string, columnName: string, nextName: string): string {
  const column = findStartTags(xml, "tableColumn").find(
    (candidate) => candidate.attributes.name === columnName
  );
  if (column === undefined) {
    throw new WorkbookError(`Table XML is missing column ${columnName}`);
  }

  return `${xml.slice(0, column.start)}${upsertAttributes(column.raw, {
    name: nextName
  })}${xml.slice(column.end)}`;
}

function appendTableColumnXml(
  xml: string,
  column: Required<Pick<WorkbookTableColumn, "id" | "name">>
): string {
  const tableColumns = findFirstStartTag(xml, "tableColumns");
  if (tableColumns === undefined) {
    throw new WorkbookError("Table XML is missing tableColumns");
  }

  const count = findStartTags(
    xml.slice(tableColumns.end, findElementCloseStart(xml, tableColumns)),
    "tableColumn"
  ).length;
  const close = findElementCloseStart(xml, tableColumns);
  const tableColumnXml = `<tableColumn id="${escapeXmlAttribute(column.id)}" name="${escapeXmlAttribute(column.name)}"/>`;
  const withColumn = `${xml.slice(0, close)}${tableColumnXml}${xml.slice(close)}`;
  return `${withColumn.slice(0, tableColumns.start)}${upsertAttributes(tableColumns.raw, {
    count: String(count + 1)
  })}${withColumn.slice(tableColumns.end)}`;
}

function removeTableColumnXml(xml: string, columnName: string): string {
  const tableColumns = findFirstStartTag(xml, "tableColumns");
  if (tableColumns === undefined) {
    throw new WorkbookError("Table XML is missing tableColumns");
  }

  const column = findStartTags(xml, "tableColumn").find(
    (candidate) => candidate.attributes.name?.toLowerCase() === columnName.toLowerCase()
  );
  if (column === undefined) {
    throw new WorkbookError(`Table XML is missing column ${columnName}`);
  }

  const nextXml = `${xml.slice(0, column.start)}${xml.slice(findElementEnd(xml, column))}`;
  const currentCount = Number.parseInt(tableColumns.attributes.count ?? "1", 10);
  const count = Math.max(0, (Number.isInteger(currentCount) ? currentCount : 1) - 1);
  return `${nextXml.slice(0, tableColumns.start)}${upsertAttributes(tableColumns.raw, {
    count: String(count)
  })}${nextXml.slice(tableColumns.end)}`;
}

function assertTableColumnNameAvailable(table: WorkbookTable, columnName: string): void {
  const duplicate = table.columns.find(
    (column) => column.name?.toLowerCase() === columnName.toLowerCase()
  );
  if (duplicate !== undefined) {
    throw new WorkbookError(
      `Column name ${columnName} is already used by table ${table.displayName}`
    );
  }
}

function nextTableColumnId(columns: WorkbookTableColumn[]): number {
  return (
    Math.max(
      0,
      ...columns.map((column) => Number.parseInt(column.id ?? "", 10)).filter(Number.isInteger)
    ) + 1
  );
}

async function rewriteTableFormulaReferences(
  pkg: OoxmlPackage,
  oldNames: string[],
  nextName: string
): Promise<void> {
  for (const partName of pkg.listParts().filter((part) => part.endsWith(".xml"))) {
    const xml = await pkg.readText(partName);
    const nextXml = rewriteFormulaElements(xml, (formula) =>
      renameFormulaStructuredReferenceTable(formula, oldNames, nextName)
    );
    if (nextXml !== xml) {
      pkg.setText(partName, nextXml);
    }
  }
}

async function rewriteTableColumnFormulaReferences(
  pkg: OoxmlPackage,
  table: WorkbookTable,
  columnName: string,
  nextName: string
): Promise<void> {
  const tableNames = [...new Set([table.name, table.displayName])];

  for (const partName of pkg.listParts().filter((part) => part.endsWith(".xml"))) {
    const xml = await pkg.readText(partName);
    const nextXml = rewriteFormulaElements(xml, (formula) =>
      renameFormulaStructuredReferenceColumn(formula, tableNames, columnName, nextName, {
        includeUnqualified: partName === table.partName
      })
    );
    if (nextXml !== xml) {
      pkg.setText(partName, nextXml);
    }
  }
}

function upsertRef(rawTag: string, ref: string): string {
  if (/\sref=(["']).*?\1/.test(rawTag)) {
    return rawTag.replace(/\sref=(["']).*?\1/, ` ref="${escapeXmlAttribute(ref)}"`);
  }

  const closing = rawTag.endsWith("/>") ? "/>" : ">";
  return `${rawTag.slice(0, -closing.length)} ref="${escapeXmlAttribute(ref)}"${closing}`;
}

function upsertAttributes(rawTag: string, attributes: Record<string, string>): string {
  return Object.entries(attributes).reduce((nextTag, [name, value]) => {
    const escapedValue = escapeXmlAttribute(value);
    const attributePattern = new RegExp(`\\s${escapeRegExp(name)}=(["']).*?\\1`);
    if (attributePattern.test(nextTag)) {
      return nextTag.replace(attributePattern, ` ${name}="${escapedValue}"`);
    }

    const closing = nextTag.endsWith("/>") ? "/>" : ">";
    return `${nextTag.slice(0, -closing.length)} ${name}="${escapedValue}"${closing}`;
  }, rawTag);
}

function validateTableName(name: string): void {
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) {
    throw new WorkbookError(`Invalid table name ${name}`);
  }

  if (/^[A-Za-z]{1,3}[1-9][0-9]{0,6}$/.test(name)) {
    throw new WorkbookError(`Table name ${name} cannot look like a cell reference`);
  }
}

function validateTableColumnName(name: string): void {
  if (name.length === 0) {
    throw new WorkbookError("Table column name cannot be empty");
  }
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
