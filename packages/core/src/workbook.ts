import { PackageError, WorkbookError } from "./errors.ts";
import { type OoxmlPackage, resolveRelationshipTarget } from "./opc.ts";
import { parseSharedStrings } from "./shared-strings.ts";
import { replaceTableRows, type WorkbookTable } from "./table.ts";
import { patchCell, readCell, type CellInput, type ReadCellResult } from "./worksheet.ts";
import { findFirstStartTag, findStartTags } from "./xml.ts";

const officeDocumentRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const worksheetRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const calcChainRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

export type WorkbookSheet = {
  name: string;
  id: string;
  relationshipId: string;
  partName: string;
};

export type WorkbookInspectResult = {
  workbookPart: string;
  sheets: WorkbookSheet[];
  parts: string[];
  features: WorkbookFeatureSummary;
};

export type WorkbookFeatureSummary = {
  calcChains: number;
  charts: number;
  drawings: number;
  macros: number;
  media: number;
  pivotTables: number;
  sharedStrings: number;
  tables: number;
};

export class Workbook {
  private sharedStringsCache: string[] | undefined;

  private constructor(
    readonly pkg: OoxmlPackage,
    readonly workbookPart: string,
    private readonly sheetsByName: Map<string, WorkbookSheet>
  ) {}

  static async fromPackage(pkg: OoxmlPackage): Promise<Workbook> {
    const workbookPart = await resolveWorkbookPart(pkg);
    const sheets = await parseSheets(pkg, workbookPart);
    const sheetsByName = new Map(sheets.map((sheet) => [sheet.name, sheet]));
    return new Workbook(pkg, workbookPart, sheetsByName);
  }

  sheets(): WorkbookSheet[] {
    return [...this.sheetsByName.values()];
  }

  sheet(name: string): WorkbookSheet {
    const sheet = this.sheetsByName.get(name);
    if (sheet === undefined) {
      throw new WorkbookError(`Unknown worksheet ${name}`);
    }

    return sheet;
  }

  async patchCell(sheetName: string, address: string, value: CellInput): Promise<void> {
    const sheet = this.sheet(sheetName);
    const xml = await this.pkg.readText(sheet.partName);
    const result = patchCell(xml, address, value);
    this.pkg.setText(sheet.partName, result.xml);

    if (result.formulaChanged) {
      await this.forceRecalculateOnOpen();
    }
  }

  async readCell(sheetName: string, address: string): Promise<ReadCellResult | undefined> {
    const sheet = this.sheet(sheetName);
    const xml = await this.pkg.readText(sheet.partName);
    return readCell(xml, address, { sharedStrings: await this.sharedStrings() });
  }

  replaceTableRows(tableName: string, rows: CellInput[][]): Promise<WorkbookTable> {
    return replaceTableRows(this.pkg, tableName, rows);
  }

  async inspect(): Promise<WorkbookInspectResult> {
    const parts = this.pkg.listParts();
    return {
      workbookPart: this.workbookPart,
      sheets: this.sheets(),
      parts,
      features: summarizeFeatures(parts)
    };
  }

  write(): Promise<Uint8Array> {
    return this.pkg.write();
  }

  private async forceRecalculateOnOpen(): Promise<void> {
    const xml = await this.pkg.readText(this.workbookPart);
    const calcPr = findFirstStartTag(xml, "calcPr");

    if (calcPr === undefined) {
      const workbookClose = xml.lastIndexOf("</workbook>");
      if (workbookClose === -1) {
        throw new WorkbookError("workbook.xml is missing closing workbook tag");
      }

      this.pkg.setText(
        this.workbookPart,
        `${xml.slice(0, workbookClose)}<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>${xml.slice(workbookClose)}`
      );
      await this.removeCalcChain();
      return;
    }

    const replacement = upsertAttributes(calcPr.raw, {
      calcMode: "auto",
      fullCalcOnLoad: "1",
      forceFullCalc: "1"
    });

    this.pkg.setText(
      this.workbookPart,
      `${xml.slice(0, calcPr.start)}${replacement}${xml.slice(calcPr.end)}`
    );

    await this.removeCalcChain();
  }

  private async removeCalcChain(): Promise<void> {
    if (!this.pkg.hasPart("xl/calcChain.xml")) {
      return;
    }

    this.pkg.deletePart("xl/calcChain.xml");
    await this.pkg.removeContentTypeOverride("xl/calcChain.xml");
    await this.pkg.removeRelationships(this.workbookPart, (relationship) => {
      const target = resolveRelationshipTarget(this.workbookPart, relationship.target);
      return relationship.type === calcChainRelationship || target === "xl/calcChain.xml";
    });
  }

  private async sharedStrings(): Promise<string[]> {
    if (this.sharedStringsCache !== undefined) {
      return this.sharedStringsCache;
    }

    if (!this.pkg.hasPart("xl/sharedStrings.xml")) {
      this.sharedStringsCache = [];
      return this.sharedStringsCache;
    }

    this.sharedStringsCache = parseSharedStrings(await this.pkg.readText("xl/sharedStrings.xml"));
    return this.sharedStringsCache;
  }
}

async function resolveWorkbookPart(pkg: OoxmlPackage): Promise<string> {
  const rootRelationships = await pkg.rootRelationships();
  const workbookRelationship = rootRelationships.find(
    (rel) => rel.type === officeDocumentRelationship
  );

  if (workbookRelationship === undefined) {
    throw new PackageError("Root relationships do not point to an office document");
  }

  return resolveRelationshipTarget("", workbookRelationship.target);
}

async function parseSheets(pkg: OoxmlPackage, workbookPart: string): Promise<WorkbookSheet[]> {
  const workbookXml = await pkg.readText(workbookPart);
  const relationships = await pkg.relationshipsFor(workbookPart);
  const worksheetRelationships = new Map(
    relationships
      .filter((relationship) => relationship.type === worksheetRelationship)
      .map((relationship) => [
        relationship.id,
        resolveRelationshipTarget(workbookPart, relationship.target)
      ])
  );

  return findStartTags(workbookXml, "sheet").map((tag) => {
    const name = tag.attributes.name;
    const id = tag.attributes.sheetId;
    const relationshipId = tag.attributes["r:id"];

    if (name === undefined || id === undefined || relationshipId === undefined) {
      throw new WorkbookError("Sheet entry is missing name, sheetId, or r:id");
    }

    const partName = worksheetRelationships.get(relationshipId);
    if (partName === undefined) {
      throw new WorkbookError(`Sheet ${name} points to missing relationship ${relationshipId}`);
    }

    return {
      name,
      id,
      relationshipId,
      partName
    };
  });
}

function upsertAttributes(rawTag: string, attributes: Record<string, string>): string {
  const closing = rawTag.endsWith("/>") ? "/>" : ">";
  let tag = rawTag.slice(0, -closing.length);

  for (const [name, value] of Object.entries(attributes)) {
    const pattern = new RegExp(`\\s${name}=(["']).*?\\1`);
    if (pattern.test(tag)) {
      tag = tag.replace(pattern, ` ${name}="${value}"`);
      continue;
    }

    tag = `${tag} ${name}="${value}"`;
  }

  return `${tag}${closing}`;
}

function summarizeFeatures(parts: string[]): WorkbookFeatureSummary {
  return {
    calcChains: countParts(parts, /^xl\/calcChain\.xml$/),
    charts: countParts(parts, /^xl\/charts\//),
    drawings: countParts(parts, /^xl\/drawings\//),
    macros: countParts(parts, /^xl\/vbaProject\.bin$/),
    media: countParts(parts, /^xl\/media\//),
    pivotTables: countParts(parts, /^xl\/pivotTables\//),
    sharedStrings: countParts(parts, /^xl\/sharedStrings\.xml$/),
    tables: countParts(parts, /^xl\/tables\//)
  };
}

function countParts(parts: string[], pattern: RegExp): number {
  return parts.filter((part) => pattern.test(part)).length;
}
