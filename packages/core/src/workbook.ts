import { type CellRange, parseCellRange } from "./address.ts";
import type { Diagnostic } from "./diagnostics.ts";
import { parseDefinedNames, type WorkbookDefinedName } from "./defined-names.ts";
import { PackageError, WorkbookError } from "./errors.ts";
import {
  parseFormulaReferences,
  parseFormulaSheetReferences,
  parseFormulaStructuredReferences,
  type FormulaReference,
  type FormulaSheetReference,
  type FormulaStructuredReference
} from "./formula.ts";
import { type OoxmlPackage, type Relationship, resolveRelationshipTarget } from "./opc.ts";
import { parseSharedStrings } from "./shared-strings.ts";
import { replaceTableRows, type WorkbookTable } from "./table.ts";
import { validateWorkbookPackage, type ValidationReport } from "./validation.ts";
import {
  appendRows,
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
import {
  decodeXml,
  findElementCloseStart,
  findElementEnd,
  findFirstStartTag,
  findStartTags
} from "./xml.ts";

const officeDocumentRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const worksheetRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const calcChainRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

type MutationImpactOptions = {
  operation: "cell" | "range" | "appendRows" | "table";
  sheetPartName?: string;
  affectedRanges?: CellRange[];
};

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
  definedNames: WorkbookDefinedName[];
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
  externalRelationships: number;
  formulaCells: number;
  hiddenSheets: number;
  hyperlinks: number;
  macros: number;
  media: number;
  merges: number;
  pivotTables: number;
  sharedStrings: number;
  tables: number;
};

export type WorkbookFormula = {
  sheetName: string;
  sheetPartName: string;
  address: string;
  formula: string;
  formulaRef?: string;
  formulaType?: string;
  references: FormulaReference[];
  sheetReferences: FormulaSheetReference[];
  sharedIndex?: string;
  structuredReferences: FormulaStructuredReference[];
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
    } else {
      await this.forceRecalculateIfFormulaDependenciesTouched(
        sheet.partName,
        result.affectedRanges
      );
    }

    await this.recordMutationImpactDiagnostics({
      operation: "cell",
      sheetPartName: sheet.partName,
      affectedRanges: result.affectedRanges
    });
  }

  async patchCells(sheetName: string, patches: CellPatch[]): Promise<void> {
    const sheet = this.sheet(sheetName);
    const xml = await this.pkg.readText(sheet.partName);
    const result = patchCells(xml, patches);
    this.pkg.setText(sheet.partName, result.xml);

    if (result.formulaChanged) {
      await this.forceRecalculateOnOpen();
    } else {
      await this.forceRecalculateIfFormulaDependenciesTouched(
        sheet.partName,
        result.affectedRanges
      );
    }

    await this.recordMutationImpactDiagnostics({
      operation: "range",
      sheetPartName: sheet.partName,
      affectedRanges: result.affectedRanges
    });
  }

  async patchRange(sheetName: string, startAddress: string, values: CellInput[][]): Promise<void> {
    const sheet = this.sheet(sheetName);
    const xml = await this.pkg.readText(sheet.partName);
    const result = patchRange(xml, startAddress, values);
    this.pkg.setText(sheet.partName, result.xml);

    if (result.formulaChanged) {
      await this.forceRecalculateOnOpen();
    } else {
      await this.forceRecalculateIfFormulaDependenciesTouched(
        sheet.partName,
        result.affectedRanges
      );
    }

    await this.recordMutationImpactDiagnostics({
      operation: "range",
      sheetPartName: sheet.partName,
      affectedRanges: result.affectedRanges
    });
  }

  async appendRows(
    sheetName: string,
    rows: CellInput[][],
    options: { startColumn?: number } = {}
  ): Promise<void> {
    const sheet = this.sheet(sheetName);
    const xml = await this.pkg.readText(sheet.partName);
    const result = appendRows(xml, rows, options);
    this.pkg.setText(sheet.partName, result.xml);

    if (result.formulaChanged) {
      await this.forceRecalculateOnOpen();
    } else {
      await this.forceRecalculateIfFormulaDependenciesTouched(
        sheet.partName,
        result.affectedRanges
      );
    }

    await this.recordMutationImpactDiagnostics({
      operation: "appendRows",
      sheetPartName: sheet.partName,
      affectedRanges: result.affectedRanges
    });
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
    if (table.totalsRowCount > 0 || rows.some((row) => row.some(isFormulaValue))) {
      await this.forceRecalculateOnOpen();
    }

    await this.recordMutationImpactDiagnostics({
      operation: "table",
      sheetPartName: table.worksheetPartName
    });

    return table;
  }

  async inspect(): Promise<WorkbookInspectResult> {
    const parts = this.pkg.listParts();
    return {
      workbookPart: this.workbookPart,
      sheets: this.sheets(),
      definedNames: await this.definedNames(),
      parts,
      features: await summarizeFeatures(this.pkg, parts, this.workbookPart),
      diagnostics: this.diagnostics()
    };
  }

  validate(): Promise<ValidationReport> {
    return validateWorkbookPackage(this.pkg);
  }

  async definedNames(): Promise<WorkbookDefinedName[]> {
    return parseDefinedNames(await this.pkg.readText(this.workbookPart));
  }

  async formulas(): Promise<WorkbookFormula[]> {
    const formulas: WorkbookFormula[] = [];

    for (const sheet of this.sheets()) {
      const xml = await this.pkg.readText(sheet.partName);
      for (const cell of findStartTags(xml, "c")) {
        const address = cell.attributes.r;
        if (address === undefined || cell.selfClosing) {
          continue;
        }

        const cellXml = xml.slice(cell.start, findElementEnd(xml, cell));
        const formulaTag = findFirstStartTag(cellXml, "f");
        if (formulaTag === undefined) {
          continue;
        }

        const formula = formulaTag.selfClosing
          ? ""
          : decodeXml(cellXml.slice(formulaTag.end, findElementCloseStart(cellXml, formulaTag)));
        const workbookFormula: WorkbookFormula = {
          sheetName: sheet.name,
          sheetPartName: sheet.partName,
          address,
          formula,
          references: parseFormulaReferences(formula),
          sheetReferences: parseFormulaSheetReferences(formula),
          structuredReferences: parseFormulaStructuredReferences(formula)
        };

        if (formulaTag.attributes.t !== undefined) {
          workbookFormula.formulaType = formulaTag.attributes.t;
        }
        if (formulaTag.attributes.si !== undefined) {
          workbookFormula.sharedIndex = formulaTag.attributes.si;
        }
        if (formulaTag.attributes.ref !== undefined) {
          workbookFormula.formulaRef = formulaTag.attributes.ref;
        }

        formulas.push(workbookFormula);
      }
    }

    return formulas;
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
      const workbook = findFirstStartTag(xml, "workbook");
      if (workbook === undefined) {
        throw new WorkbookError("workbook.xml is missing closing workbook tag");
      }

      const workbookClose = findElementCloseStart(xml, workbook);
      const calcPrTag = qualifiedName(xmlPrefix(workbook.name), "calcPr");
      this.pkg.setText(
        this.workbookPart,
        `${xml.slice(0, workbookClose)}<${calcPrTag} calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>${xml.slice(workbookClose)}`
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

  private async forceRecalculateIfFormulaDependenciesTouched(
    sheetPartName: string,
    affectedRanges: CellRange[]
  ): Promise<void> {
    if (affectedRanges.length === 0) {
      return;
    }

    const affected = await this.formulasAffectedByRanges(sheetPartName, affectedRanges);
    if (affected.length === 0) {
      return;
    }

    await this.forceRecalculateOnOpen();
    this.addDiagnostic({
      severity: "info",
      code: "FORMULA_DEPENDENCIES_RECALCULATED",
      message: `Marked workbook for recalculation because ${affected.length} formula cell(s) reference edited range(s)`,
      part: sheetPartName
    });
  }

  private async formulasAffectedByRanges(
    sheetPartName: string,
    affectedRanges: CellRange[]
  ): Promise<WorkbookFormula[]> {
    const sheetPartByName = new Map(
      this.sheets().map((sheet) => [sheet.name.toLowerCase(), sheet.partName])
    );
    const affected: WorkbookFormula[] = [];

    for (const formula of await this.formulas()) {
      for (const reference of formula.references) {
        const referenceSheetPart =
          reference.sheetName === undefined
            ? formula.sheetPartName
            : sheetPartByName.get(reference.sheetName.toLowerCase());
        if (referenceSheetPart !== sheetPartName) {
          continue;
        }

        const referenceRange = formulaReferenceRange(reference);
        if (
          affectedRanges.some((affectedRange) => rangesIntersect(referenceRange, affectedRange))
        ) {
          affected.push(formula);
          break;
        }
      }
    }

    return affected;
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

  private async recordMutationImpactDiagnostics(options: MutationImpactOptions): Promise<void> {
    const parts = this.pkg.listParts();
    const definedNames = await this.definedNames();

    if (definedNames.length > 0) {
      this.addDiagnostic({
        severity: "warning",
        code: "DEFINED_NAMES_MAY_NEED_REVIEW",
        message: `Workbook has ${definedNames.length} defined name(s); verify edited ranges still match template intent`
      });
    }

    if (countParts(parts, /^xl\/charts\//) > 0) {
      this.addDiagnostic({
        severity: "warning",
        code: "CHARTS_MAY_NEED_REFRESH",
        message: "Workbook contains charts; verify chart ranges after worksheet data edits"
      });
    }

    if (countParts(parts, /^xl\/pivotTables\//) > 0) {
      this.addDiagnostic({
        severity: "warning",
        code: "PIVOT_TABLES_MAY_NEED_REFRESH",
        message:
          "Workbook contains pivot tables; Excel may need to refresh pivot caches after data edits"
      });
    }

    if (
      options.operation !== "table" &&
      options.sheetPartName !== undefined &&
      (await worksheetHasTableParts(this.pkg, options.sheetPartName))
    ) {
      this.addDiagnostic({
        severity: "warning",
        code: "WORKSHEET_TABLES_NOT_RESIZED",
        message:
          "Worksheet contains tables; direct cell/range edits do not resize table refs. Use replaceTableRows for table body changes.",
        part: options.sheetPartName
      });
    }
  }

  private addDiagnostic(diagnostic: Diagnostic): void {
    if (
      this.diagnosticJournal.some(
        (existing) => existing.code === diagnostic.code && existing.part === diagnostic.part
      )
    ) {
      return;
    }

    this.diagnosticJournal.push(diagnostic);
  }
}

function isFormulaValue(value: CellInput): value is FormulaValue {
  return (
    typeof value === "object" && value !== null && !(value instanceof Date) && "formula" in value
  );
}

function formulaReferenceRange(reference: FormulaReference): CellRange {
  if (reference.kind === "range") {
    return reference.range;
  }

  return parseCellRange(reference.ref);
}

function rangesIntersect(left: CellRange, right: CellRange): boolean {
  return (
    left.start.column <= right.end.column &&
    left.end.column >= right.start.column &&
    left.start.row <= right.end.row &&
    left.end.row >= right.start.row
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

async function worksheetHasTableParts(pkg: OoxmlPackage, sheetPartName: string): Promise<boolean> {
  if (!pkg.hasPart(sheetPartName)) {
    return false;
  }

  return findStartTags(await pkg.readText(sheetPartName), "tablePart").length > 0;
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

function xmlPrefix(name: string): string | undefined {
  const colon = name.indexOf(":");
  return colon === -1 ? undefined : name.slice(0, colon);
}

function qualifiedName(prefix: string | undefined, localName: string): string {
  return prefix === undefined ? localName : `${prefix}:${localName}`;
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
  const inspection = await pkg.inspect();

  return {
    calcChains: countParts(parts, /^xl\/calcChain\.xml$/),
    charts: countParts(parts, /^xl\/charts\//),
    comments:
      countParts(parts, /^xl\/comments\d*\.xml$/) + countParts(parts, /^xl\/threadedComments\//),
    conditionalFormats: countXmlStartTags(worksheetXml, "conditionalFormatting"),
    dataValidations: countXmlStartTags(worksheetXml, "dataValidation"),
    definedNames: countXmlStartTags([workbookXml], "definedName"),
    drawings: countParts(parts, /^xl\/drawings\/drawing\d+\.xml$/),
    externalRelationships: Object.values(inspection.relationships).reduce(
      (count, relationships) =>
        count +
        relationships.filter((relationship) => relationship.targetMode === "External").length,
      0
    ),
    formulaCells: countXmlStartTags(worksheetXml, "f"),
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
