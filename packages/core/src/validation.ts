import { parseCellAddress, parseCellRange } from "./address.ts";
import { parseDefinedNames } from "./defined-names.ts";
import { type OoxmlPackage, type Relationship, resolveRelationshipTarget } from "./opc.ts";
import { findElementEnd, findFirstStartTag, findStartTags } from "./xml.ts";

const chartRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const drawingRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const imageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const tableRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const worksheetRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

export type ValidationSeverity = "error" | "warning" | "info";

export type ValidationIssue = {
  severity: ValidationSeverity;
  code: string;
  message: string;
  part?: string;
  target?: string;
};

export type ValidationReport = {
  issues: ValidationIssue[];
  summary: {
    errors: number;
    warnings: number;
    infos: number;
  };
};

export async function validateWorkbookPackage(pkg: OoxmlPackage): Promise<ValidationReport> {
  const issues: ValidationIssue[] = [];
  const parts = pkg.listParts();
  const partSet = new Set(parts);

  await validateRelationshipTargets(pkg, issues, partSet);
  validateRelationshipParts(issues, parts);
  await validateContentTypes(pkg, issues, parts);
  await validateWorkbookSheets(pkg, issues);
  await validateWorksheetRelationshipIds(pkg, issues, parts);
  await validateDrawingRelationshipIds(pkg, issues, parts);
  await validateDefinedNames(pkg, issues);
  await validateWorksheetDimensions(pkg, issues, parts);
  await validateTableParts(pkg, issues, parts);
  await validateCalcChain(pkg, issues, parts);

  return {
    issues,
    summary: {
      errors: issues.filter((issue) => issue.severity === "error").length,
      warnings: issues.filter((issue) => issue.severity === "warning").length,
      infos: issues.filter((issue) => issue.severity === "info").length
    }
  };
}

async function validateRelationshipTargets(
  pkg: OoxmlPackage,
  issues: ValidationIssue[],
  partSet: Set<string>
): Promise<void> {
  const inspection = await pkg.inspect();

  for (const [sourcePart, relationships] of Object.entries(inspection.relationships)) {
    const relationshipIds = new Set<string>();
    for (const relationship of relationships) {
      if (relationshipIds.has(relationship.id)) {
        issues.push({
          severity: "error",
          code: "RELATIONSHIP_ID_DUPLICATE",
          message: `Relationship part ${sourcePart} contains duplicate relationship id ${relationship.id}`,
          part: sourcePart,
          target: relationship.id
        });
      }
      relationshipIds.add(relationship.id);

      if (relationship.targetMode === "External") {
        continue;
      }

      const source = sourcePart === "/" ? "" : sourcePart;
      const target = resolveRelationshipTarget(source, relationship.target);
      if (!partSet.has(target)) {
        issues.push({
          severity: "error",
          code: "RELATIONSHIP_TARGET_MISSING",
          message: `Relationship ${relationship.id} points to missing part ${target}`,
          part: sourcePart,
          target
        });
      }
    }
  }
}

async function validateContentTypes(
  pkg: OoxmlPackage,
  issues: ValidationIssue[],
  parts: string[]
): Promise<void> {
  if (!pkg.hasPart("[Content_Types].xml")) {
    issues.push({
      severity: "error",
      code: "CONTENT_TYPES_MISSING",
      message: "Package is missing [Content_Types].xml"
    });
    return;
  }

  const contentTypes = parseContentTypes(await pkg.readText("[Content_Types].xml"));
  const partSet = new Set(parts);
  for (const override of contentTypes.overrides) {
    if (!partSet.has(override)) {
      issues.push({
        severity: "warning",
        code: "CONTENT_TYPE_OVERRIDE_ORPHAN",
        message: `Content type override points to missing part ${override}`,
        part: "[Content_Types].xml",
        target: override
      });
    }
  }

  for (const part of parts) {
    if (part === "[Content_Types].xml" || part.endsWith(".rels")) {
      continue;
    }

    const extension = part.includes(".") ? part.slice(part.lastIndexOf(".") + 1) : "";
    if (!contentTypes.overrides.has(part) && !contentTypes.defaults.has(extension)) {
      issues.push({
        severity: "error",
        code: "CONTENT_TYPE_MISSING",
        message: `Part ${part} has no content type override or default`,
        part
      });
    }
  }
}

function validateRelationshipParts(issues: ValidationIssue[], parts: string[]): void {
  const partSet = new Set(parts);
  for (const part of parts.filter((name) => name.endsWith(".rels"))) {
    const sourcePart = sourcePartFromRelationshipPart(part);
    if (sourcePart === undefined || sourcePart === "/") {
      continue;
    }

    if (!partSet.has(sourcePart)) {
      issues.push({
        severity: "warning",
        code: "RELATIONSHIP_PART_ORPHAN",
        message: `Relationship part ${part} has no source part ${sourcePart}`,
        part,
        target: sourcePart
      });
    }
  }
}

async function workbookSheetPartsBySheetId(pkg: OoxmlPackage): Promise<Map<string, string>> {
  if (!pkg.hasPart("xl/workbook.xml")) {
    return new Map();
  }

  const workbookXml = await pkg.readText("xl/workbook.xml");
  const relationshipsById = new Map(
    (await pkg.relationshipsFor("xl/workbook.xml")).map((relationship) => [
      relationship.id,
      relationship
    ])
  );
  const result = new Map<string, string>();

  for (const sheet of findStartTags(workbookXml, "sheet")) {
    const sheetId = sheet.attributes.sheetId;
    const relationshipId = sheet.attributes["r:id"];
    if (sheetId === undefined || relationshipId === undefined) {
      continue;
    }

    const relationship = relationshipsById.get(relationshipId);
    if (relationship?.type === worksheetRelationship) {
      result.set(sheetId, resolveRelationshipTarget("xl/workbook.xml", relationship.target));
    }
  }

  return result;
}

async function validateWorkbookSheets(pkg: OoxmlPackage, issues: ValidationIssue[]): Promise<void> {
  if (!pkg.hasPart("xl/workbook.xml")) {
    return;
  }

  const workbookXml = await pkg.readText("xl/workbook.xml");
  const relationshipsById = new Map(
    (await pkg.relationshipsFor("xl/workbook.xml")).map((relationship) => [
      relationship.id,
      relationship
    ])
  );
  const sheetNames = new Set<string>();
  const sheetIds = new Set<string>();

  for (const sheet of findStartTags(workbookXml, "sheet")) {
    const name = sheet.attributes.name;
    const sheetId = sheet.attributes.sheetId;
    const relationshipId = sheet.attributes["r:id"];

    if (name === undefined || sheetId === undefined || relationshipId === undefined) {
      issues.push({
        severity: "error",
        code: "WORKBOOK_SHEET_ENTRY_INVALID",
        message: "Workbook sheet entry is missing name, sheetId, or r:id",
        part: "xl/workbook.xml"
      });
      continue;
    }

    if (sheetNames.has(name)) {
      issues.push({
        severity: "error",
        code: "WORKBOOK_SHEET_NAME_DUPLICATE",
        message: `Workbook contains duplicate sheet name ${name}`,
        part: "xl/workbook.xml",
        target: name
      });
    }
    sheetNames.add(name);

    if (sheetIds.has(sheetId)) {
      issues.push({
        severity: "warning",
        code: "WORKBOOK_SHEET_ID_DUPLICATE",
        message: `Workbook contains duplicate sheetId ${sheetId}`,
        part: "xl/workbook.xml",
        target: sheetId
      });
    }
    sheetIds.add(sheetId);

    const relationship = relationshipsById.get(relationshipId);
    if (relationship === undefined) {
      issues.push({
        severity: "error",
        code: "WORKBOOK_SHEET_RELATIONSHIP_MISSING",
        message: `Workbook sheet ${name} points to missing relationship ${relationshipId}`,
        part: "xl/workbook.xml",
        target: relationshipId
      });
      continue;
    }

    if (relationship.type !== worksheetRelationship) {
      issues.push({
        severity: "error",
        code: "WORKBOOK_SHEET_RELATIONSHIP_INVALID",
        message: `Workbook sheet ${name} points to a non-worksheet relationship ${relationshipId}`,
        part: "xl/workbook.xml",
        target: relationshipId
      });
      continue;
    }

    const target = resolveRelationshipTarget("xl/workbook.xml", relationship.target);
    if (!/^xl\/worksheets\/.+\.xml$/.test(target)) {
      issues.push({
        severity: "warning",
        code: "WORKBOOK_SHEET_TARGET_UNUSUAL",
        message: `Workbook sheet ${name} points outside the standard worksheet folder: ${target}`,
        part: "xl/workbook.xml",
        target
      });
    }
  }
}

async function validateCalcChain(
  pkg: OoxmlPackage,
  issues: ValidationIssue[],
  parts: string[]
): Promise<void> {
  const calcChainParts = parts.filter((part) => /^xl\/calcChain.*\.xml$/.test(part));
  if (calcChainParts.length === 0) {
    return;
  }

  if (calcChainParts.length > 1) {
    issues.push({
      severity: "warning",
      code: "CALC_CHAIN_MULTIPLE_PARTS",
      message: `Workbook contains ${calcChainParts.length} calculation chain parts`,
      part: "xl/calcChain.xml"
    });
  }

  if (!pkg.hasPart("xl/calcChain.xml")) {
    return;
  }

  const sheetPartsById = await workbookSheetPartsBySheetId(pkg);
  const worksheetXmlByPart = new Map<string, string>();
  const calcChainXml = await pkg.readText("xl/calcChain.xml");
  if (findFirstStartTag(calcChainXml, "calcChain") === undefined) {
    issues.push({
      severity: "error",
      code: "CALC_CHAIN_ROOT_MISSING",
      message: "Calculation chain part is missing calcChain root",
      part: "xl/calcChain.xml"
    });
    return;
  }

  let currentSheetId: string | undefined;
  for (const cell of findStartTags(calcChainXml, "c")) {
    const address = cell.attributes.r;
    if (cell.attributes.i !== undefined) {
      currentSheetId = cell.attributes.i;
    }

    if (address === undefined) {
      issues.push({
        severity: "error",
        code: "CALC_CHAIN_CELL_REF_MISSING",
        message: "Calculation chain cell is missing r attribute",
        part: "xl/calcChain.xml"
      });
      continue;
    }

    try {
      parseCellAddress(address);
    } catch (_error) {
      issues.push({
        severity: "error",
        code: "CALC_CHAIN_CELL_REF_INVALID",
        message: `Calculation chain cell has invalid address ${address}`,
        part: "xl/calcChain.xml",
        target: address
      });
      continue;
    }

    if (currentSheetId === undefined) {
      issues.push({
        severity: "warning",
        code: "CALC_CHAIN_SHEET_ID_MISSING",
        message: `Calculation chain cell ${address} has no sheet id context`,
        part: "xl/calcChain.xml",
        target: address
      });
      continue;
    }

    const worksheetPart = sheetPartsById.get(currentSheetId);
    if (worksheetPart === undefined) {
      issues.push({
        severity: "warning",
        code: "CALC_CHAIN_SHEET_MISSING",
        message: `Calculation chain cell ${address} references missing sheet id ${currentSheetId}`,
        part: "xl/calcChain.xml",
        target: currentSheetId
      });
      continue;
    }

    if (!pkg.hasPart(worksheetPart)) {
      continue;
    }

    let worksheetXml = worksheetXmlByPart.get(worksheetPart);
    if (worksheetXml === undefined) {
      worksheetXml = await pkg.readText(worksheetPart);
      worksheetXmlByPart.set(worksheetPart, worksheetXml);
    }

    if (!worksheetCellHasFormula(worksheetXml, address)) {
      issues.push({
        severity: "warning",
        code: "CALC_CHAIN_CELL_NOT_FORMULA",
        message: `Calculation chain references ${address}, but the worksheet cell has no formula`,
        part: "xl/calcChain.xml",
        target: `${worksheetPart}!${address}`
      });
    }
  }
}

async function validateDefinedNames(pkg: OoxmlPackage, issues: ValidationIssue[]): Promise<void> {
  if (!pkg.hasPart("xl/workbook.xml")) {
    return;
  }

  const workbookXml = await pkg.readText("xl/workbook.xml");
  const sheets = findStartTags(workbookXml, "sheet");
  const sheetNames = new Set(
    sheets.map((tag) => tag.attributes.name).filter((name): name is string => name !== undefined)
  );

  for (const definedName of parseDefinedNames(workbookXml)) {
    if (definedName.localSheetId !== undefined) {
      const localSheetId = Number.parseInt(definedName.localSheetId, 10);
      if (!Number.isInteger(localSheetId) || localSheetId < 0 || localSheetId >= sheets.length) {
        issues.push({
          severity: "warning",
          code: "DEFINED_NAME_LOCAL_SHEET_MISSING",
          message: `Defined name ${definedName.name} has invalid localSheetId ${definedName.localSheetId}`,
          part: "xl/workbook.xml",
          target: definedName.name
        });
      }
    }

    for (const sheetName of sheetReferencesInFormulaText(definedName.text)) {
      if (!sheetNames.has(sheetName)) {
        issues.push({
          severity: "warning",
          code: "DEFINED_NAME_SHEET_MISSING",
          message: `Defined name ${definedName.name} references missing sheet ${sheetName}`,
          part: "xl/workbook.xml",
          target: definedName.name
        });
      }
    }
  }
}

async function validateWorksheetRelationshipIds(
  pkg: OoxmlPackage,
  issues: ValidationIssue[],
  parts: string[]
): Promise<void> {
  for (const part of parts.filter((name) => /^xl\/worksheets\/.+\.xml$/.test(name))) {
    const xml = await pkg.readText(part);
    const relationships = await pkg.relationshipsFor(part);

    for (const drawing of findStartTags(xml, "drawing")) {
      validateRelationshipId({
        issues,
        relationships,
        sourcePart: part,
        relationshipId: drawing.attributes["r:id"],
        expectedType: drawingRelationship,
        missingCode: "DRAWING_RELATIONSHIP_MISSING",
        message: "Worksheet drawing element points to a missing drawing relationship"
      });
    }

    const tableParts = findStartTags(xml, "tablePart");
    const tablePartsContainer = findFirstStartTag(xml, "tableParts");
    const declaredCount = Number.parseInt(tablePartsContainer?.attributes.count ?? "", 10);
    if (Number.isInteger(declaredCount) && declaredCount !== tableParts.length) {
      issues.push({
        severity: "warning",
        code: "TABLE_PART_COUNT_MISMATCH",
        message: `Worksheet declares ${declaredCount} table part(s) but contains ${tableParts.length}`,
        part
      });
    }

    for (const tablePart of tableParts) {
      validateRelationshipId({
        issues,
        relationships,
        sourcePart: part,
        relationshipId: tablePart.attributes["r:id"],
        expectedType: tableRelationship,
        missingCode: "TABLE_PART_RELATIONSHIP_MISSING",
        message: "Worksheet tablePart element points to a missing table relationship"
      });
    }
  }
}

async function validateDrawingRelationshipIds(
  pkg: OoxmlPackage,
  issues: ValidationIssue[],
  parts: string[]
): Promise<void> {
  for (const part of parts.filter((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name))) {
    const xml = await pkg.readText(part);
    const relationships = await pkg.relationshipsFor(part);

    for (const chart of findStartTags(xml, "chart")) {
      validateRelationshipId({
        issues,
        relationships,
        sourcePart: part,
        relationshipId: chart.attributes["r:id"],
        expectedType: chartRelationship,
        missingCode: "DRAWING_CHART_RELATIONSHIP_MISSING",
        message: "Drawing chart element points to a missing chart relationship"
      });
    }

    for (const blip of findStartTags(xml, "blip")) {
      validateRelationshipId({
        issues,
        relationships,
        sourcePart: part,
        relationshipId: blip.attributes["r:embed"] ?? blip.attributes["r:link"],
        expectedType: imageRelationship,
        missingCode: "DRAWING_IMAGE_RELATIONSHIP_MISSING",
        message: "Drawing image element points to a missing image relationship"
      });
    }
  }
}

function validateRelationshipId(options: {
  issues: ValidationIssue[];
  relationships: Relationship[];
  sourcePart: string;
  relationshipId: string | undefined;
  expectedType: string;
  missingCode: string;
  message: string;
}): void {
  if (options.relationshipId === undefined) {
    options.issues.push({
      severity: "error",
      code: options.missingCode,
      message: options.message,
      part: options.sourcePart
    });
    return;
  }

  const relationship = options.relationships.find(
    (candidate) => candidate.id === options.relationshipId
  );
  if (relationship?.type !== options.expectedType) {
    options.issues.push({
      severity: "error",
      code: options.missingCode,
      message: `${options.message}: ${options.relationshipId}`,
      part: options.sourcePart,
      target: options.relationshipId
    });
  }
}

function worksheetCellHasFormula(xml: string, address: string): boolean {
  const target = address.toUpperCase();
  for (const cell of findStartTags(xml, "c")) {
    if ((cell.attributes.r ?? "").toUpperCase() !== target) {
      continue;
    }

    if (cell.selfClosing) {
      return false;
    }

    return findFirstStartTag(xml.slice(cell.start, findElementEnd(xml, cell)), "f") !== undefined;
  }

  return false;
}

function sheetReferencesInFormulaText(text: string): string[] {
  const references = new Set<string>();
  const pattern = /(?:^|[, (])((?:'(?:(?:'')|[^'])+'|[A-Za-z_][A-Za-z0-9_ .]*))!/g;

  for (const match of text.matchAll(pattern)) {
    const rawName = match[1];
    if (rawName === undefined || rawName.includes("[")) {
      continue;
    }

    references.add(unquoteSheetName(rawName));
  }

  return [...references];
}

function unquoteSheetName(name: string): string {
  if (name.startsWith("'") && name.endsWith("'")) {
    return name.slice(1, -1).replaceAll("''", "'");
  }

  return name.trim();
}

function sourcePartFromRelationshipPart(part: string): string | undefined {
  if (part === "_rels/.rels") {
    return "/";
  }

  const marker = "/_rels/";
  const markerIndex = part.lastIndexOf(marker);
  if (markerIndex !== -1 && part.endsWith(".rels")) {
    return `${part.slice(0, markerIndex + 1)}${part.slice(markerIndex + marker.length, -".rels".length)}`;
  }

  if (part.startsWith("_rels/") && part.endsWith(".rels")) {
    return part.slice("_rels/".length, -".rels".length);
  }

  return undefined;
}

async function validateWorksheetDimensions(
  pkg: OoxmlPackage,
  issues: ValidationIssue[],
  parts: string[]
): Promise<void> {
  for (const part of parts.filter((name) => /^xl\/worksheets\/.+\.xml$/.test(name))) {
    const xml = await pkg.readText(part);
    const dimension = findFirstStartTag(xml, "dimension");
    if (dimension === undefined || dimension.attributes.ref === undefined) {
      issues.push({
        severity: "warning",
        code: "WORKSHEET_DIMENSION_MISSING",
        message: "Worksheet is missing a dimension ref",
        part
      });
      continue;
    }

    const range = parseCellRange(dimension.attributes.ref);
    for (const cell of findStartTags(xml, "c")) {
      const address = cell.attributes.r;
      if (address === undefined) {
        continue;
      }

      const parsed = parseCellAddress(address);
      if (
        parsed.column < range.start.column ||
        parsed.column > range.end.column ||
        parsed.row < range.start.row ||
        parsed.row > range.end.row
      ) {
        issues.push({
          severity: "warning",
          code: "WORKSHEET_DIMENSION_EXCLUDES_CELL",
          message: `Worksheet dimension ${range.ref} does not include cell ${parsed.address}`,
          part,
          target: parsed.address
        });
      }
    }
  }
}

async function validateTableParts(
  pkg: OoxmlPackage,
  issues: ValidationIssue[],
  parts: string[]
): Promise<void> {
  for (const part of parts.filter((name) => /^xl\/tables\/.+\.xml$/.test(name))) {
    const xml = await pkg.readText(part);
    const table = findFirstStartTag(xml, "table");
    if (table === undefined || table.attributes.ref === undefined) {
      issues.push({
        severity: "error",
        code: "TABLE_REF_MISSING",
        message: "Table part is missing a table ref",
        part
      });
      continue;
    }

    let range: ReturnType<typeof parseCellRange>;
    try {
      range = parseCellRange(table.attributes.ref);
    } catch (_error) {
      issues.push({
        severity: "error",
        code: "TABLE_REF_INVALID",
        message: `Table part has invalid ref ${table.attributes.ref}`,
        part,
        target: table.attributes.ref
      });
      continue;
    }

    const autoFilter = findFirstStartTag(xml, "autoFilter");
    if (
      autoFilter?.attributes.ref !== undefined &&
      autoFilter.attributes.ref !== table.attributes.ref
    ) {
      issues.push({
        severity: "warning",
        code: "TABLE_AUTOFILTER_REF_MISMATCH",
        message: `Table ref ${table.attributes.ref} does not match autoFilter ref ${autoFilter.attributes.ref}`,
        part
      });
    }

    const tableColumns = findStartTags(xml, "tableColumn");
    const tableColumnsContainer = findFirstStartTag(xml, "tableColumns");
    const declaredCount = Number.parseInt(tableColumnsContainer?.attributes.count ?? "", 10);
    if (Number.isInteger(declaredCount) && declaredCount !== tableColumns.length) {
      issues.push({
        severity: "warning",
        code: "TABLE_COLUMN_COUNT_MISMATCH",
        message: `Table declares ${declaredCount} column(s) but contains ${tableColumns.length}`,
        part
      });
    }

    const refWidth = range.end.column - range.start.column + 1;
    if (tableColumns.length > 0 && tableColumns.length !== refWidth) {
      issues.push({
        severity: "warning",
        code: "TABLE_COLUMN_REF_WIDTH_MISMATCH",
        message: `Table ref ${range.ref} spans ${refWidth} column(s) but defines ${tableColumns.length}`,
        part,
        target: range.ref
      });
    }
  }
}

function parseContentTypes(xml: string): {
  defaults: Set<string>;
  overrides: Set<string>;
} {
  return {
    defaults: new Set(
      findStartTags(xml, "Default")
        .map((tag) => tag.attributes.Extension)
        .filter((extension): extension is string => extension !== undefined)
    ),
    overrides: new Set(
      findStartTags(xml, "Override")
        .map((tag) => tag.attributes.PartName)
        .filter((partName): partName is string => partName !== undefined)
        .map((partName) => partName.replace(/^\/+/, ""))
    )
  };
}
