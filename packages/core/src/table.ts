import { numberToColumnLabel, parseCellAddress } from "./address.ts";
import { WorkbookError } from "./errors.ts";
import type { OoxmlPackage } from "./opc.ts";
import { resolveRelationshipTarget } from "./opc.ts";
import { replaceRowsInRange, type CellInput } from "./worksheet.ts";
import { escapeXmlAttribute, findFirstStartTag } from "./xml.ts";

const tableRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

export type WorkbookTable = {
  name: string;
  displayName: string;
  partName: string;
  worksheetPartName: string;
  ref: string;
  totalsRowCount: number;
};

export async function findWorkbookTable(
  pkg: OoxmlPackage,
  tableName: string
): Promise<WorkbookTable> {
  for (const partName of pkg.listParts().filter((part) => /^xl\/tables\/.+\.xml$/.test(part))) {
    const tableXml = await pkg.readText(partName);
    const table = findFirstStartTag(tableXml, "table");
    if (table === undefined) {
      continue;
    }

    const name = table.attributes.name ?? table.attributes.displayName;
    const displayName = table.attributes.displayName ?? name;
    const ref = table.attributes.ref;

    if (name !== tableName && displayName !== tableName) {
      continue;
    }

    if (name === undefined || displayName === undefined || ref === undefined) {
      throw new WorkbookError(`Table part ${partName} is missing name/displayName/ref`);
    }

    return {
      name,
      displayName,
      partName,
      worksheetPartName: await findWorksheetForTable(pkg, partName),
      ref,
      totalsRowCount: parseTotalsRowCount(table.attributes)
    };
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

function parseTotalsRowCount(attributes: Record<string, string>): number {
  const explicitCount = Number.parseInt(attributes.totalsRowCount ?? "", 10);
  if (Number.isInteger(explicitCount) && explicitCount > 0) {
    return explicitCount;
  }

  return attributes.totalsRowShown === "1" ? 1 : 0;
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

function upsertRef(rawTag: string, ref: string): string {
  if (/\sref=(["']).*?\1/.test(rawTag)) {
    return rawTag.replace(/\sref=(["']).*?\1/, ` ref="${escapeXmlAttribute(ref)}"`);
  }

  const closing = rawTag.endsWith("/>") ? "/>" : ">";
  return `${rawTag.slice(0, -closing.length)} ref="${escapeXmlAttribute(ref)}"${closing}`;
}
