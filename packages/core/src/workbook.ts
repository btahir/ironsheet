import type { Diagnostic } from "./diagnostics.ts";
import { PackageError, WorkbookError } from "./errors.ts";
import { type OoxmlPackage, type Relationship, resolveRelationshipTarget } from "./opc.ts";
import { parseSharedStrings } from "./shared-strings.ts";
import { replaceTableRows, type WorkbookTable } from "./table.ts";
import {
  patchCell,
  patchCells,
  patchRange,
  readCell,
  readRange,
  type CellInput,
  type CellPatch,
  type FormulaValue,
  type ReadCellResult,
  type ReadRangeResult
} from "./worksheet.ts";
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
  state?: WorkbookSheetState;
};

export type WorkbookSheetState = "hidden" | "veryHidden";

export type WorkbookInspectResult = {
  workbookPart: string;
  sheets: WorkbookSheet[];
  parts: string[];
  features: WorkbookFeatureSummary;
  diagnostics: Diagnostic[];
};

export type WorkbookFeatureSummary = {
  calcChains: number;
  charts: number;
  comments: number;
  conditionalFormats: number;
  dataValidations: number;
  definedNames: number;
  drawings: number;
  hiddenSheets: number;
  hyperlinks: number;
  macros: number;
  media: number;
  merges: number;
  pivotTables: number;
  sharedStrings: number;
  tables: number;
};

export class Workbook {
  private sharedStringsCache: string[] | undefined;
  private readonly diagnosticJournal: Diagnostic[] = [];

  private constructor(
    readonly pkg: OoxmlPackage,
    readonly workbookPart: string,
    private readonly sheetsByName: Map<string, WorkbookSheet>
  ) {}

  static async fromPackage(pkg: OoxmlPackage): Promise<Workbook> {
    const workbookPart = await resolveWorkbookPart(pkg);
    const sheets = await parseSheets(pkg, workbookPart);
    const sheetsByName = new Map(sheets.map((sheet) => [sheet.name, sheet]));
    const workbook = new Workbook(pkg, workbookPart, sheetsByName);
    await workbook.recordInitialDiagnostics();
    return workbook;
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

  async patchCells(sheetName: string, patches: CellPatch[]): Promise<void> {
    const sheet = this.sheet(sheetName);
    const xml = await this.pkg.readText(sheet.partName);
    const result = patchCells(xml, patches);
    this.pkg.setText(sheet.partName, result.xml);

    if (result.formulaChanged) {
      await this.forceRecalculateOnOpen();
    }
  }

  async patchRange(sheetName: string, startAddress: string, values: CellInput[][]): Promise<void> {
    const sheet = this.sheet(sheetName);
    const xml = await this.pkg.readText(sheet.partName);
    const result = patchRange(xml, startAddress, values);
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

  async readRange(sheetName: string, rangeRef: string): Promise<ReadRangeResult> {
    const sheet = this.sheet(sheetName);
    const xml = await this.pkg.readText(sheet.partName);
    return readRange(xml, rangeRef, { sharedStrings: await this.sharedStrings() });
  }

  async replaceTableRows(tableName: string, rows: CellInput[][]): Promise<WorkbookTable> {
    const table = await replaceTableRows(this.pkg, tableName, rows);
    if (rows.some((row) => row.some(isFormulaValue))) {
      await this.forceRecalculateOnOpen();
    }

    return table;
  }

  async inspect(): Promise<WorkbookInspectResult> {
    const parts = this.pkg.listParts();
    return {
      workbookPart: this.workbookPart,
      sheets: this.sheets(),
      parts,
      features: await summarizeFeatures(this.pkg, parts, this.workbookPart),
      diagnostics: this.diagnostics()
    };
  }

  diagnostics(): Diagnostic[] {
    return [...this.diagnosticJournal];
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
    this.addDiagnostic({
      severity: "info",
      code: "FORMULA_CALC_CHAIN_REMOVED",
      message: "Removed stale calcChain.xml after formula mutation",
      part: "xl/calcChain.xml"
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

  private async recordInitialDiagnostics(): Promise<void> {
    if (this.pkg.hasPart("xl/vbaProject.bin")) {
      this.addDiagnostic({
        severity: "info",
        code: "MACRO_PROJECT_PRESERVED",
        message: "Workbook contains a VBA project; Ironsheet preserves but never executes macros",
        part: "xl/vbaProject.bin"
      });
    }

    for (const [part, relationships] of Object.entries((await this.pkg.inspect()).relationships)) {
      for (const relationship of relationships) {
        this.recordRelationshipDiagnostic(part, relationship);
      }
    }
  }

  private recordRelationshipDiagnostic(part: string, relationship: Relationship): void {
    if (relationship.targetMode === "External") {
      this.addDiagnostic({
        severity: "warning",
        code: "EXTERNAL_RELATIONSHIP_PRESERVED",
        message: `Preserved external relationship ${relationship.id}`,
        part
      });
    }
  }

  private addDiagnostic(diagnostic: Diagnostic): void {
    this.diagnosticJournal.push(diagnostic);
  }
}

function isFormulaValue(value: CellInput): value is FormulaValue {
  return (
    typeof value === "object" && value !== null && !(value instanceof Date) && "formula" in value
  );
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

    const state = sheetState(tag.attributes.state);
    return {
      name,
      id,
      relationshipId,
      partName,
      ...(state === undefined ? {} : { state })
    };
  });
}

function sheetState(value: string | undefined): WorkbookSheetState | undefined {
  if (value === "hidden" || value === "veryHidden") {
    return value;
  }

  return undefined;
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

async function summarizeFeatures(
  pkg: OoxmlPackage,
  parts: string[],
  workbookPart: string
): Promise<WorkbookFeatureSummary> {
  const worksheetXml = await Promise.all(
    parts.filter((part) => /^xl\/worksheets\/.+\.xml$/.test(part)).map((part) => pkg.readText(part))
  );
  const workbookXml = await pkg.readText(workbookPart);

  return {
    calcChains: countParts(parts, /^xl\/calcChain\.xml$/),
    charts: countParts(parts, /^xl\/charts\//),
    comments:
      countParts(parts, /^xl\/comments\d*\.xml$/) + countParts(parts, /^xl\/threadedComments\//),
    conditionalFormats: countXmlStartTags(worksheetXml, "conditionalFormatting"),
    dataValidations: countXmlStartTags(worksheetXml, "dataValidation"),
    definedNames: countXmlStartTags([workbookXml], "definedName"),
    drawings: countParts(parts, /^xl\/drawings\/drawing\d+\.xml$/),
    hiddenSheets: findStartTags(workbookXml, "sheet").filter((tag) =>
      ["hidden", "veryHidden"].includes(tag.attributes.state ?? "")
    ).length,
    hyperlinks: countXmlStartTags(worksheetXml, "hyperlink"),
    macros: countParts(parts, /^xl\/vbaProject\.bin$/),
    media: countParts(parts, /^xl\/media\//),
    merges: countXmlStartTags(worksheetXml, "mergeCell"),
    pivotTables: countParts(parts, /^xl\/pivotTables\//),
    sharedStrings: countParts(parts, /^xl\/sharedStrings\.xml$/),
    tables: countParts(parts, /^xl\/tables\//)
  };
}

function countParts(parts: string[], pattern: RegExp): number {
  return parts.filter((part) => pattern.test(part)).length;
}

function countXmlStartTags(xmlParts: string[], localName: string): number {
  return xmlParts.reduce((count, xml) => count + findStartTags(xml, localName).length, 0);
}
