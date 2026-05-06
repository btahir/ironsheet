import { type CellRange, parseCellRange } from "./address.ts";
import { retargetWorkbookChartFormulas, type ChartFormulaRetarget } from "./chart.ts";
import {
  parseWorksheetComments,
  worksheetCommentsRelationship,
  type WorkbookComment
} from "./comments.ts";
import type { Diagnostic } from "./diagnostics.ts";
import { parseDefinedNames, type WorkbookDefinedName } from "./defined-names.ts";
import { PackageError, WorkbookError } from "./errors.ts";
import { rewriteFormulaElements } from "./formula-rewrite.ts";
import {
  parseFormulaReferences,
  parseFormulaSheetReferences,
  parseFormulaStructuredReferences,
  renameFormulaSheetReferences,
  type FormulaReference,
  type FormulaSheetReference,
  type FormulaStructuredReference
} from "./formula.ts";
import { drawingRelationship, imageRelationship, type WorkbookImage } from "./images.ts";
import {
  normalizePartName,
  type OoxmlPackage,
  type Relationship,
  resolveRelationshipTarget
} from "./opc.ts";
import { retargetWorkbookPivotCacheSources, type PivotCacheSourceRetarget } from "./pivot.ts";
import { parseSharedStrings } from "./shared-strings.ts";
import {
  ensureWorkbookCellFormat,
  ensureWorkbookNumberFormat,
  parseWorkbookStyles,
  type WorkbookCellFormat,
  type WorkbookCellStyleInput,
  type WorkbookStyles
} from "./styles.ts";
import {
  appendWorkbookTableColumn,
  listWorkbookTables,
  removeRightmostWorkbookTableColumn,
  renameWorkbookTableColumn,
  renameWorkbookTable,
  replaceTableRows,
  type WorkbookTable
} from "./table.ts";
import { validateWorkbookPackage, type ValidationReport } from "./validation.ts";
import {
  appendRows,
  applyCellStyle,
  deleteWorksheetAutoFilter,
  deleteWorksheetConditionalFormat,
  deleteWorksheetDataValidation,
  deleteWorksheetHyperlink,
  listWorksheetAutoFilters,
  listWorksheetConditionalFormats,
  listWorksheetDataValidations,
  listWorksheetHyperlinks,
  listWorksheetMergedCells,
  mergeWorksheetCells,
  patchCell,
  patchCells,
  patchRange,
  readCell,
  readRange,
  setWorksheetAutoFilter,
  setWorksheetConditionalFormat,
  setWorksheetDataValidation,
  setWorksheetHyperlink,
  unmergeWorksheetCells,
  type CellInput,
  type CellPatch,
  type FormulaValue,
  type ReadCellResult,
  type ReadRangeResult,
  type WorksheetAutoFilter,
  type WorksheetConditionalFormat,
  type WorksheetDataValidation
} from "./worksheet.ts";
import {
  decodeXml,
  escapeXmlAttribute,
  escapeXmlText,
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
const hyperlinkRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

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

export type WorkbookHyperlink = {
  sheetName: string;
  sheetPartName: string;
  ref: string;
  display?: string;
  location?: string;
  relationshipId?: string;
  target?: string;
  targetMode?: string;
  tooltip?: string;
};

export type WorkbookMergedCell = {
  sheetName: string;
  sheetPartName: string;
  ref: string;
};

export type WorkbookDataValidation = WorksheetDataValidation & {
  sheetName: string;
  sheetPartName: string;
};

export type WorkbookConditionalFormat = WorksheetConditionalFormat & {
  sheetName: string;
  sheetPartName: string;
};

export type WorkbookAutoFilter = WorksheetAutoFilter & {
  sheetName: string;
  sheetPartName: string;
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

  async hyperlinks(sheetName?: string): Promise<WorkbookHyperlink[]> {
    const sheets = sheetName === undefined ? this.sheets() : [this.sheet(sheetName)];
    const hyperlinks: WorkbookHyperlink[] = [];

    for (const sheet of sheets) {
      const xml = await this.pkg.readText(sheet.partName);
      const relationships = new Map(
        (await this.pkg.relationshipsFor(sheet.partName)).map((relationship) => [
          relationship.id,
          relationship
        ])
      );

      for (const hyperlink of listWorksheetHyperlinks(xml)) {
        const relationship =
          hyperlink.relationshipId === undefined
            ? undefined
            : relationships.get(hyperlink.relationshipId);
        hyperlinks.push({
          sheetName: sheet.name,
          sheetPartName: sheet.partName,
          ref: hyperlink.ref,
          ...(hyperlink.display === undefined ? {} : { display: hyperlink.display }),
          ...(hyperlink.location === undefined ? {} : { location: hyperlink.location }),
          ...(hyperlink.relationshipId === undefined
            ? {}
            : { relationshipId: hyperlink.relationshipId }),
          ...(relationship?.target === undefined ? {} : { target: relationship.target }),
          ...(relationship?.targetMode === undefined
            ? {}
            : { targetMode: relationship.targetMode }),
          ...(hyperlink.tooltip === undefined ? {} : { tooltip: hyperlink.tooltip })
        });
      }
    }

    return hyperlinks;
  }

  async mergedCells(sheetName?: string): Promise<WorkbookMergedCell[]> {
    const sheets = sheetName === undefined ? this.sheets() : [this.sheet(sheetName)];
    const merges: WorkbookMergedCell[] = [];

    for (const sheet of sheets) {
      const xml = await this.pkg.readText(sheet.partName);
      for (const merge of listWorksheetMergedCells(xml)) {
        merges.push({
          sheetName: sheet.name,
          sheetPartName: sheet.partName,
          ref: merge.ref
        });
      }
    }

    return merges;
  }

  async dataValidations(sheetName?: string): Promise<WorkbookDataValidation[]> {
    const sheets = sheetName === undefined ? this.sheets() : [this.sheet(sheetName)];
    const dataValidations: WorkbookDataValidation[] = [];

    for (const sheet of sheets) {
      const xml = await this.pkg.readText(sheet.partName);
      for (const dataValidation of listWorksheetDataValidations(xml)) {
        dataValidations.push({
          sheetName: sheet.name,
          sheetPartName: sheet.partName,
          ...dataValidation
        });
      }
    }

    return dataValidations;
  }

  async conditionalFormats(sheetName?: string): Promise<WorkbookConditionalFormat[]> {
    const sheets = sheetName === undefined ? this.sheets() : [this.sheet(sheetName)];
    const conditionalFormats: WorkbookConditionalFormat[] = [];

    for (const sheet of sheets) {
      const xml = await this.pkg.readText(sheet.partName);
      for (const conditionalFormat of listWorksheetConditionalFormats(xml)) {
        conditionalFormats.push({
          sheetName: sheet.name,
          sheetPartName: sheet.partName,
          ...conditionalFormat
        });
      }
    }

    return conditionalFormats;
  }

  async autoFilters(sheetName?: string): Promise<WorkbookAutoFilter[]> {
    const sheets = sheetName === undefined ? this.sheets() : [this.sheet(sheetName)];
    const autoFilters: WorkbookAutoFilter[] = [];

    for (const sheet of sheets) {
      const xml = await this.pkg.readText(sheet.partName);
      for (const autoFilter of listWorksheetAutoFilters(xml)) {
        autoFilters.push({
          sheetName: sheet.name,
          sheetPartName: sheet.partName,
          ...autoFilter
        });
      }
    }

    return autoFilters;
  }

  async comments(sheetName?: string): Promise<WorkbookComment[]> {
    const sheets = sheetName === undefined ? this.sheets() : [this.sheet(sheetName)];
    const comments: WorkbookComment[] = [];

    for (const sheet of sheets) {
      const relationships = await this.pkg.relationshipsFor(sheet.partName);
      for (const relationship of relationships) {
        if (relationship.type !== worksheetCommentsRelationship) {
          continue;
        }

        const commentPartName = resolveRelationshipTarget(sheet.partName, relationship.target);
        if (!this.pkg.hasPart(commentPartName)) {
          continue;
        }

        for (const comment of parseWorksheetComments(await this.pkg.readText(commentPartName))) {
          comments.push({
            sheetName: sheet.name,
            sheetPartName: sheet.partName,
            commentPartName,
            relationshipId: relationship.id,
            ...comment
          });
        }
      }
    }

    return comments;
  }

  async images(sheetName?: string): Promise<WorkbookImage[]> {
    const sheets = sheetName === undefined ? this.sheets() : [this.sheet(sheetName)];
    const images: WorkbookImage[] = [];

    for (const sheet of sheets) {
      const worksheetXml = await this.pkg.readText(sheet.partName);
      const worksheetRelationships = new Map(
        (await this.pkg.relationshipsFor(sheet.partName)).map((relationship) => [
          relationship.id,
          relationship
        ])
      );

      for (const drawing of findStartTags(worksheetXml, "drawing")) {
        const drawingRelationshipId = drawing.attributes["r:id"];
        if (drawingRelationshipId === undefined) {
          continue;
        }

        const worksheetRelationship = worksheetRelationships.get(drawingRelationshipId);
        if (worksheetRelationship?.type !== drawingRelationship) {
          continue;
        }

        const drawingPartName = resolveRelationshipTarget(
          sheet.partName,
          worksheetRelationship.target
        );
        if (!this.pkg.hasPart(drawingPartName)) {
          continue;
        }

        const drawingXml = await this.pkg.readText(drawingPartName);
        const drawingRelationships = new Map(
          (await this.pkg.relationshipsFor(drawingPartName)).map((relationship) => [
            relationship.id,
            relationship
          ])
        );

        for (const blip of findStartTags(drawingXml, "blip")) {
          const imageRelationshipId = blip.attributes["r:embed"] ?? blip.attributes["r:link"];
          if (imageRelationshipId === undefined) {
            continue;
          }

          const image = drawingRelationships.get(imageRelationshipId);
          if (image?.type !== imageRelationship) {
            continue;
          }

          const imagePartName =
            image.targetMode === "External"
              ? undefined
              : resolveRelationshipTarget(drawingPartName, image.target);
          images.push({
            sheetName: sheet.name,
            sheetPartName: sheet.partName,
            drawingPartName,
            drawingRelationshipId,
            imageRelationshipId,
            target: image.target,
            ...(imagePartName === undefined ? {} : { imagePartName }),
            ...(image.targetMode === undefined ? {} : { targetMode: image.targetMode })
          });
        }
      }
    }

    return images;
  }

  async replaceImage(imagePartName: string, data: Uint8Array): Promise<WorkbookImage> {
    const normalized = normalizePartName(imagePartName);
    const image = (await this.images()).find((candidate) => candidate.imagePartName === normalized);
    if (image === undefined) {
      throw new WorkbookError(`Unknown workbook image part ${normalized}`);
    }

    this.pkg.setPart(normalized, data);
    return image;
  }

  async setDataValidation(
    sheetName: string,
    dataValidation: WorksheetDataValidation
  ): Promise<WorkbookDataValidation> {
    const sheet = this.sheet(sheetName);
    const result = setWorksheetDataValidation(
      await this.pkg.readText(sheet.partName),
      dataValidation
    );
    this.pkg.setText(sheet.partName, result.xml);

    return {
      sheetName: sheet.name,
      sheetPartName: sheet.partName,
      ...result.dataValidation
    };
  }

  async setAutoFilter(
    sheetName: string,
    autoFilter: WorksheetAutoFilter
  ): Promise<WorkbookAutoFilter> {
    const sheet = this.sheet(sheetName);
    const result = setWorksheetAutoFilter(await this.pkg.readText(sheet.partName), autoFilter);
    this.pkg.setText(sheet.partName, result.xml);

    return {
      sheetName: sheet.name,
      sheetPartName: sheet.partName,
      ...result.autoFilter
    };
  }

  async setConditionalFormat(
    sheetName: string,
    conditionalFormat: WorksheetConditionalFormat
  ): Promise<WorkbookConditionalFormat> {
    const sheet = this.sheet(sheetName);
    const result = setWorksheetConditionalFormat(
      await this.pkg.readText(sheet.partName),
      conditionalFormat
    );
    this.pkg.setText(sheet.partName, result.xml);

    return {
      sheetName: sheet.name,
      sheetPartName: sheet.partName,
      ...result.conditionalFormat
    };
  }

  async deleteDataValidation(sheetName: string, sqref: string): Promise<boolean> {
    const sheet = this.sheet(sheetName);
    const result = deleteWorksheetDataValidation(await this.pkg.readText(sheet.partName), sqref);
    if (!result.deleted) {
      return false;
    }

    this.pkg.setText(sheet.partName, result.xml);
    return true;
  }

  async deleteAutoFilter(sheetName: string): Promise<boolean> {
    const sheet = this.sheet(sheetName);
    const result = deleteWorksheetAutoFilter(await this.pkg.readText(sheet.partName));
    if (!result.deleted) {
      return false;
    }

    this.pkg.setText(sheet.partName, result.xml);
    return true;
  }

  async deleteConditionalFormat(sheetName: string, sqref: string): Promise<boolean> {
    const sheet = this.sheet(sheetName);
    const result = deleteWorksheetConditionalFormat(await this.pkg.readText(sheet.partName), sqref);
    if (!result.deleted) {
      return false;
    }

    this.pkg.setText(sheet.partName, result.xml);
    return true;
  }

  async mergeCells(sheetName: string, ref: string): Promise<WorkbookMergedCell> {
    const sheet = this.sheet(sheetName);
    const result = mergeWorksheetCells(await this.pkg.readText(sheet.partName), ref);
    if (result.merged) {
      this.pkg.setText(sheet.partName, result.xml);
    }

    return {
      sheetName: sheet.name,
      sheetPartName: sheet.partName,
      ref: result.merge.ref
    };
  }

  async unmergeCells(sheetName: string, ref: string): Promise<boolean> {
    const sheet = this.sheet(sheetName);
    const result = unmergeWorksheetCells(await this.pkg.readText(sheet.partName), ref);
    if (!result.unmerged) {
      return false;
    }

    this.pkg.setText(sheet.partName, result.xml);
    return true;
  }

  async setHyperlink(
    sheetName: string,
    ref: string,
    target: string,
    options: { display?: string; tooltip?: string } = {}
  ): Promise<WorkbookHyperlink> {
    const sheet = this.sheet(sheetName);
    const xml = await this.pkg.readText(sheet.partName);
    const normalizedRef = parseCellRange(ref).ref;
    const existing = listWorksheetHyperlinks(xml).find(
      (hyperlink) => hyperlink.ref === normalizedRef
    );
    const relationshipId =
      existing?.relationshipId ?? (await this.pkg.nextRelationshipId(sheet.partName));
    const result = setWorksheetHyperlink(xml, {
      ref,
      relationshipId,
      ...(options.display === undefined ? {} : { display: options.display }),
      ...(options.tooltip === undefined ? {} : { tooltip: options.tooltip })
    });

    this.pkg.setText(sheet.partName, result.xml);
    await this.pkg.upsertRelationship(sheet.partName, {
      id: relationshipId,
      type: hyperlinkRelationship,
      target,
      targetMode: "External"
    });

    return {
      sheetName: sheet.name,
      sheetPartName: sheet.partName,
      ref: result.hyperlink.ref,
      relationshipId,
      target,
      targetMode: "External",
      ...(options.display === undefined ? {} : { display: options.display }),
      ...(options.tooltip === undefined ? {} : { tooltip: options.tooltip })
    };
  }

  async deleteHyperlink(sheetName: string, ref: string): Promise<boolean> {
    const sheet = this.sheet(sheetName);
    const result = deleteWorksheetHyperlink(await this.pkg.readText(sheet.partName), ref);
    if (!result.deleted) {
      return false;
    }

    this.pkg.setText(sheet.partName, result.xml);
    const relationshipIds = new Set(result.relationshipIds);
    await this.pkg.removeRelationships(sheet.partName, (relationship) =>
      relationshipIds.has(relationship.id)
    );
    return true;
  }

  async ensureCellStyle(style: WorkbookCellStyleInput): Promise<string> {
    let stylesXml = await this.readOrCreateStylesXml();
    const styleInput = { ...style };
    if (styleInput.numberFormat !== undefined) {
      const numberFormat = ensureWorkbookNumberFormat(stylesXml, styleInput.numberFormat);
      stylesXml = numberFormat.xml;
      styleInput.numFmtId = numberFormat.numFmtId;
      styleInput.applyNumberFormat = "1";
    }

    const result = ensureWorkbookCellFormat(stylesXml, cellFormatFromStyleInput(styleInput));
    this.pkg.setText("xl/styles.xml", result.xml);
    return result.styleId;
  }

  async styleCell(
    sheetName: string,
    address: string,
    style: WorkbookCellStyleInput
  ): Promise<string> {
    const sheet = this.sheet(sheetName);
    const sheetXml = await this.pkg.readText(sheet.partName);
    const existingCell = readCell(sheetXml, address, { sharedStrings: await this.sharedStrings() });
    const baseStyle =
      existingCell?.styleId === undefined
        ? {}
        : await this.cellFormatForStyleId(existingCell.styleId);
    const styleId = await this.ensureCellStyle({ ...baseStyle, ...style });
    const result = applyCellStyle(sheetXml, address, styleId);
    this.pkg.setText(sheet.partName, result.xml);
    return styleId;
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

  async renameTable(tableName: string, nextName: string): Promise<WorkbookTable> {
    const table = await renameWorkbookTable(this.pkg, tableName, nextName);
    await this.forceRecalculateOnOpen();
    await this.recordMutationImpactDiagnostics({
      operation: "table",
      sheetPartName: table.worksheetPartName
    });

    return table;
  }

  async renameTableColumn(
    tableName: string,
    columnName: string,
    nextName: string
  ): Promise<WorkbookTable> {
    const table = await renameWorkbookTableColumn(this.pkg, tableName, columnName, nextName);
    await this.forceRecalculateOnOpen();
    await this.recordMutationImpactDiagnostics({
      operation: "table",
      sheetPartName: table.worksheetPartName
    });

    return table;
  }

  async appendTableColumn(
    tableName: string,
    columnName: string,
    values: CellInput[] = []
  ): Promise<WorkbookTable> {
    const table = await appendWorkbookTableColumn(this.pkg, tableName, columnName, values);
    await this.forceRecalculateOnOpen();
    await this.recordMutationImpactDiagnostics({
      operation: "table",
      sheetPartName: table.worksheetPartName
    });

    return table;
  }

  async removeRightmostTableColumn(tableName: string, columnName: string): Promise<WorkbookTable> {
    const table = await removeRightmostWorkbookTableColumn(this.pkg, tableName, columnName);
    await this.forceRecalculateOnOpen();
    await this.recordMutationImpactDiagnostics({
      operation: "table",
      sheetPartName: table.worksheetPartName
    });

    return table;
  }

  async renameSheet(sheetName: string, nextName: string): Promise<WorkbookSheet> {
    validateSheetName(nextName);

    const sheet = this.sheet(sheetName);
    const duplicate = this.sheets().find((candidate) => {
      return (
        candidate.partName !== sheet.partName &&
        candidate.name.toLowerCase() === nextName.toLowerCase()
      );
    });
    if (duplicate !== undefined) {
      throw new WorkbookError(`Sheet name ${nextName} is already used by ${duplicate.name}`);
    }

    const workbookXml = await this.pkg.readText(this.workbookPart);
    const sheetTag = findStartTags(workbookXml, "sheet").find(
      (candidate) => candidate.attributes["r:id"] === sheet.relationshipId
    );
    if (sheetTag === undefined) {
      throw new WorkbookError(`Workbook XML is missing sheet ${sheetName}`);
    }

    this.pkg.setText(
      this.workbookPart,
      `${workbookXml.slice(0, sheetTag.start)}${upsertAttributes(sheetTag.raw, {
        name: nextName
      })}${workbookXml.slice(sheetTag.end)}`
    );

    await this.rewriteSheetFormulaReferences(sheet.name, nextName);
    await this.retargetPivotCacheSources([
      { from: { sheet: sheet.name }, to: { sheet: nextName } }
    ]);
    await this.forceRecalculateOnOpen();

    const renamed = { ...sheet, name: nextName };
    this.sheetsByName.delete(sheet.name);
    this.sheetsByName.set(nextName, renamed);
    return renamed;
  }

  async setSheetState(
    sheetName: string,
    state: WorkbookSheetState | undefined
  ): Promise<WorkbookSheet> {
    const sheet = this.sheet(sheetName);
    if (
      state !== undefined &&
      !this.sheets().some(
        (candidate) => candidate.name !== sheetName && candidate.state === undefined
      )
    ) {
      throw new WorkbookError("Workbook must keep at least one visible worksheet");
    }

    const workbookXml = await this.pkg.readText(this.workbookPart);
    const sheetTag = findStartTags(workbookXml, "sheet").find(
      (candidate) => candidate.attributes["r:id"] === sheet.relationshipId
    );
    if (sheetTag === undefined) {
      throw new WorkbookError(`Workbook XML is missing sheet ${sheetName}`);
    }

    const updatedSheetTag =
      state === undefined
        ? removeAttributes(sheetTag.raw, ["state"])
        : upsertAttributes(sheetTag.raw, { state });
    this.pkg.setText(
      this.workbookPart,
      `${workbookXml.slice(0, sheetTag.start)}${updatedSheetTag}${workbookXml.slice(sheetTag.end)}`
    );

    const updated =
      state === undefined
        ? {
            name: sheet.name,
            id: sheet.id,
            relationshipId: sheet.relationshipId,
            partName: sheet.partName
          }
        : { ...sheet, state };
    this.sheetsByName.set(sheet.name, updated);
    return updated;
  }

  hideSheet(sheetName: string, state: WorkbookSheetState = "hidden"): Promise<WorkbookSheet> {
    return this.setSheetState(sheetName, state);
  }

  showSheet(sheetName: string): Promise<WorkbookSheet> {
    return this.setSheetState(sheetName, undefined);
  }

  retargetChartFormulas(retargets: ChartFormulaRetarget[]): Promise<number> {
    return retargetWorkbookChartFormulas(this.pkg, retargets);
  }

  retargetPivotCacheSources(retargets: PivotCacheSourceRetarget[]): Promise<number> {
    return retargetWorkbookPivotCacheSources(this.pkg, retargets);
  }

  tables(): Promise<WorkbookTable[]> {
    return listWorkbookTables(this.pkg);
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

  async setDefinedName(
    name: string,
    text: string,
    options: { comment?: string; hidden?: boolean; sheetName?: string } = {}
  ): Promise<WorkbookDefinedName> {
    validateDefinedName(name);

    const localSheetId =
      options.sheetName === undefined ? undefined : this.localSheetIdForSheet(options.sheetName);
    const definedName: WorkbookDefinedName = {
      name,
      text,
      ...(options.comment === undefined ? {} : { comment: options.comment }),
      ...(options.hidden === undefined ? {} : { hidden: options.hidden }),
      ...(localSheetId === undefined ? {} : { localSheetId })
    };

    this.pkg.setText(
      this.workbookPart,
      setDefinedNameXml(await this.pkg.readText(this.workbookPart), definedName)
    );
    await this.forceRecalculateOnOpen();
    return definedName;
  }

  async deleteDefinedName(name: string, options: { sheetName?: string } = {}): Promise<boolean> {
    const localSheetId =
      options.sheetName === undefined ? undefined : this.localSheetIdForSheet(options.sheetName);
    const result = deleteDefinedNameXml(await this.pkg.readText(this.workbookPart), {
      name,
      ...(localSheetId === undefined ? {} : { localSheetId })
    });
    if (!result.deleted) {
      return false;
    }

    this.pkg.setText(this.workbookPart, result.xml);
    await this.forceRecalculateOnOpen();
    return true;
  }

  async styles(): Promise<WorkbookStyles> {
    if (!this.pkg.hasPart("xl/styles.xml")) {
      return {
        cellStyleXfs: [],
        cellXfs: [],
        counts: {
          borders: 0,
          cellStyleXfs: 0,
          cellXfs: 0,
          fills: 0,
          fonts: 0,
          numFmts: 0
        },
        numberFormats: []
      };
    }

    return parseWorkbookStyles(await this.pkg.readText("xl/styles.xml"));
  }

  private async readOrCreateStylesXml(): Promise<string> {
    if (this.pkg.hasPart("xl/styles.xml")) {
      return this.pkg.readText("xl/styles.xml");
    }

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>';
  }

  private async cellFormatForStyleId(styleId: string): Promise<WorkbookCellFormat> {
    const index = Number.parseInt(styleId, 10);
    if (!Number.isInteger(index) || index < 0) {
      throw new WorkbookError(`Invalid style id ${styleId}`);
    }

    const style = (await this.styles()).cellXfs[index];
    if (style === undefined) {
      throw new WorkbookError(`Unknown style id ${styleId}`);
    }

    return style;
  }

  private localSheetIdForSheet(sheetName: string): string {
    const index = this.sheets().findIndex((sheet) => sheet.name === sheetName);
    if (index === -1) {
      throw new WorkbookError(`Unknown worksheet ${sheetName}`);
    }

    return String(index);
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

  private async rewriteSheetFormulaReferences(sheetName: string, nextName: string): Promise<void> {
    for (const partName of this.pkg.listParts().filter((part) => part.endsWith(".xml"))) {
      const xml = await this.pkg.readText(partName);
      const nextXml = rewriteFormulaElements(xml, (formula) =>
        renameFormulaSheetReferences(formula, sheetName, nextName)
      );
      if (nextXml !== xml) {
        this.pkg.setText(partName, nextXml);
      }
    }
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

function cellFormatFromStyleInput(style: WorkbookCellStyleInput): WorkbookCellFormat {
  return {
    ...(style.applyAlignment === undefined ? {} : { applyAlignment: style.applyAlignment }),
    ...(style.applyBorder === undefined ? {} : { applyBorder: style.applyBorder }),
    ...(style.applyFill === undefined ? {} : { applyFill: style.applyFill }),
    ...(style.applyFont === undefined ? {} : { applyFont: style.applyFont }),
    ...(style.applyNumberFormat === undefined
      ? {}
      : { applyNumberFormat: style.applyNumberFormat }),
    ...(style.borderId === undefined ? {} : { borderId: style.borderId }),
    ...(style.fillId === undefined ? {} : { fillId: style.fillId }),
    ...(style.fontId === undefined ? {} : { fontId: style.fontId }),
    ...(style.numFmtId === undefined ? {} : { numFmtId: style.numFmtId }),
    ...(style.xfId === undefined ? {} : { xfId: style.xfId })
  };
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

function validateSheetName(name: string): void {
  if (name.length === 0) {
    throw new WorkbookError("Sheet name cannot be empty");
  }

  if (name.length > 31) {
    throw new WorkbookError(`Sheet name ${name} exceeds Excel's 31 character limit`);
  }

  if (/[:\\/?*[\]]/.test(name)) {
    throw new WorkbookError(`Sheet name ${name} contains an invalid Excel sheet name character`);
  }

  if (name.startsWith("'") || name.endsWith("'")) {
    throw new WorkbookError("Sheet name cannot start or end with an apostrophe");
  }
}

function validateDefinedName(name: string): void {
  if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) {
    throw new WorkbookError(`Invalid defined name ${name}`);
  }

  if (/^[A-Za-z]{1,3}[1-9][0-9]{0,6}$/.test(name)) {
    throw new WorkbookError(`Defined name ${name} cannot look like a cell reference`);
  }
}

function setDefinedNameXml(workbookXml: string, definedName: WorkbookDefinedName): string {
  const definedNameXml = buildDefinedNameXml(workbookXml, definedName);
  const existing = findMatchingDefinedName(workbookXml, definedName);
  if (existing !== undefined) {
    return `${workbookXml.slice(0, existing.start)}${definedNameXml}${workbookXml.slice(findElementEnd(workbookXml, existing))}`;
  }

  const definedNames = findFirstStartTag(workbookXml, "definedNames");
  if (definedNames !== undefined) {
    const close = findElementCloseStart(workbookXml, definedNames);
    return `${workbookXml.slice(0, close)}${definedNameXml}${workbookXml.slice(close)}`;
  }

  const workbook = findFirstStartTag(workbookXml, "workbook");
  if (workbook === undefined) {
    throw new WorkbookError("workbook.xml is missing workbook tag");
  }

  const prefix = xmlPrefix(workbook.name);
  const definedNamesTag = qualifiedName(prefix, "definedNames");
  const calcPr = findFirstStartTag(workbookXml, "calcPr");
  const insertOffset = calcPr?.start ?? findElementCloseStart(workbookXml, workbook);
  return `${workbookXml.slice(0, insertOffset)}<${definedNamesTag}>${definedNameXml}</${definedNamesTag}>${workbookXml.slice(insertOffset)}`;
}

function deleteDefinedNameXml(
  workbookXml: string,
  target: { localSheetId?: string; name: string }
): { deleted: boolean; xml: string } {
  const existing = findMatchingDefinedName(workbookXml, target);
  if (existing === undefined) {
    return { deleted: false, xml: workbookXml };
  }

  const nextXml = `${workbookXml.slice(0, existing.start)}${workbookXml.slice(findElementEnd(workbookXml, existing))}`;
  const definedNames = findFirstStartTag(nextXml, "definedNames");
  if (definedNames === undefined) {
    return { deleted: true, xml: nextXml };
  }

  const body = nextXml.slice(definedNames.end, findElementCloseStart(nextXml, definedNames));
  if (findStartTags(body, "definedName").length > 0) {
    return { deleted: true, xml: nextXml };
  }

  return {
    deleted: true,
    xml: `${nextXml.slice(0, definedNames.start)}${nextXml.slice(findElementEnd(nextXml, definedNames))}`
  };
}

function buildDefinedNameXml(workbookXml: string, definedName: WorkbookDefinedName): string {
  const tagName = qualifiedName(workbookPrefix(workbookXml), "definedName");
  const attributes = [
    `name="${escapeXmlAttribute(definedName.name)}"`,
    definedName.comment === undefined
      ? undefined
      : `comment="${escapeXmlAttribute(definedName.comment)}"`,
    definedName.hidden === undefined ? undefined : `hidden="${definedName.hidden ? "1" : "0"}"`,
    definedName.localSheetId === undefined
      ? undefined
      : `localSheetId="${escapeXmlAttribute(definedName.localSheetId)}"`
  ]
    .filter((attribute): attribute is string => attribute !== undefined)
    .join(" ");

  return `<${tagName} ${attributes}>${escapeXmlText(definedName.text)}</${tagName}>`;
}

function findMatchingDefinedName(
  workbookXml: string,
  target: { localSheetId?: string; name: string }
) {
  return findStartTags(workbookXml, "definedName").find((definedName) => {
    return (
      definedName.attributes.name?.toLowerCase() === target.name.toLowerCase() &&
      definedName.attributes.localSheetId === target.localSheetId
    );
  });
}

function workbookPrefix(workbookXml: string): string | undefined {
  const workbook = findFirstStartTag(workbookXml, "workbook");
  if (workbook === undefined) {
    throw new WorkbookError("workbook.xml is missing workbook tag");
  }

  return xmlPrefix(workbook.name);
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
    const escapedValue = escapeXmlAttribute(value);
    const pattern = new RegExp(`\\s${name}=(["']).*?\\1`);
    if (pattern.test(tag)) {
      tag = tag.replace(pattern, ` ${name}="${escapedValue}"`);
      continue;
    }

    tag = `${tag} ${name}="${escapedValue}"`;
  }

  return `${tag}${closing}`;
}

function removeAttributes(rawTag: string, names: string[]): string {
  let tag = rawTag;
  for (const name of names) {
    const pattern = new RegExp(`\\s${name}=(["']).*?\\1`);
    tag = tag.replace(pattern, "");
  }

  return tag;
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
