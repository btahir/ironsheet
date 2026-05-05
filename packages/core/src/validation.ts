import { parseCellAddress, parseCellRange } from "./address.ts";
import { type OoxmlPackage, type Relationship, resolveRelationshipTarget } from "./opc.ts";
import { findFirstStartTag, findStartTags } from "./xml.ts";

const chartRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const drawingRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const imageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const tableRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

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
  await validateContentTypes(pkg, issues, parts);
  await validateWorksheetRelationshipIds(pkg, issues, parts);
  await validateDrawingRelationshipIds(pkg, issues, parts);
  await validateWorksheetDimensions(pkg, issues, parts);
  await validateTableParts(pkg, issues, parts);

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
    for (const relationship of relationships) {
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
