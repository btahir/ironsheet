import { PackageError } from "./errors.ts";
import { escapeXmlAttribute, findElementCloseStart, findElementEnd, findStartTags } from "./xml.ts";
import {
  type CompressionAdapter,
  parseZip,
  readEntryData,
  writeZip,
  type ZipEntry
} from "./zip/index.ts";

const textDecoder = new TextDecoder();
const textEncoder = new TextEncoder();

export type Relationship = {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
};

export type PackagePart = {
  name: string;
  entry?: ZipEntry;
  dirtyText?: string;
};

export type PackageInspectResult = {
  parts: string[];
  relationships: Record<string, Relationship[]>;
};

export class OoxmlPackage {
  readonly parts = new Map<string, PackagePart>();

  private constructor(
    entries: ZipEntry[],
    private readonly compression: CompressionAdapter
  ) {
    for (const entry of entries) {
      this.parts.set(normalizePartName(entry.name), {
        name: normalizePartName(entry.name),
        entry
      });
    }
  }

  static open(data: Uint8Array, compression: CompressionAdapter): OoxmlPackage {
    return new OoxmlPackage(parseZip(data).entries, compression);
  }

  listParts(): string[] {
    return [...this.parts.keys()].sort();
  }

  hasPart(partName: string): boolean {
    return this.parts.has(normalizePartName(partName));
  }

  async readPart(partName: string): Promise<Uint8Array> {
    const part = this.requirePart(partName);
    if (part.dirtyText !== undefined) {
      return textEncoder.encode(part.dirtyText);
    }

    if (part.entry === undefined) {
      throw new PackageError(`Part ${part.name} has no payload`);
    }

    return readEntryData(part.entry, this.compression);
  }

  async readText(partName: string): Promise<string> {
    return textDecoder.decode(await this.readPart(partName));
  }

  setText(partName: string, text: string): void {
    const normalized = normalizePartName(partName);
    const part = this.parts.get(normalized);

    if (part === undefined) {
      throw new PackageError(`Cannot set missing part ${normalized}`);
    }

    part.dirtyText = text;
  }

  addTextPart(partName: string, text: string): void {
    const normalized = normalizePartName(partName);
    if (this.parts.has(normalized)) {
      throw new PackageError(`Cannot add existing part ${normalized}`);
    }

    this.parts.set(normalized, {
      name: normalized,
      dirtyText: text
    });
  }

  deletePart(partName: string): boolean {
    return this.parts.delete(normalizePartName(partName));
  }

  async relationshipsFor(partName: string): Promise<Relationship[]> {
    const normalized = normalizePartName(partName);
    const relationshipPart = relationshipPartName(normalized);

    if (!this.hasPart(relationshipPart)) {
      return [];
    }

    return parseRelationships(await this.readText(relationshipPart));
  }

  async removeRelationships(
    partName: string,
    predicate: (relationship: Relationship) => boolean
  ): Promise<number> {
    const normalized = normalizePartName(partName);
    const relationshipPart = relationshipPartName(normalized);

    if (!this.hasPart(relationshipPart)) {
      return 0;
    }

    const xml = await this.readText(relationshipPart);
    const removals = findStartTags(xml, "Relationship")
      .map((tag) => ({
        tag,
        relationship: relationshipFromAttributes(tag.attributes)
      }))
      .filter(({ relationship }) => predicate(relationship));

    if (removals.length === 0) {
      return 0;
    }

    let nextXml = xml;
    for (const removal of removals.toReversed()) {
      const end = removal.tag.selfClosing ? removal.tag.end : findElementEnd(nextXml, removal.tag);
      nextXml = `${nextXml.slice(0, removal.tag.start)}${nextXml.slice(end)}`;
    }

    this.setText(relationshipPart, nextXml);
    return removals.length;
  }

  async nextRelationshipId(partName: string, prefix = "rId"): Promise<string> {
    const relationships = await this.relationshipsFor(partName);
    const used = new Set(relationships.map((relationship) => relationship.id));
    const highestNumericId = Math.max(
      0,
      ...relationships
        .map((relationship) =>
          relationship.id.match(new RegExp(`^${escapeRegExp(prefix)}([1-9][0-9]*)$`))
        )
        .filter((match): match is RegExpMatchArray => match !== null)
        .map((match) => Number.parseInt(match[1] ?? "0", 10))
    );

    let next = highestNumericId + 1;
    while (used.has(`${prefix}${next}`)) {
      next += 1;
    }

    return `${prefix}${next}`;
  }

  async upsertRelationship(partName: string, relationship: Relationship): Promise<void> {
    const normalized = normalizePartName(partName);
    const relationshipPart = relationshipPartName(normalized);
    const relationshipXml = relationshipElementXml(relationship);

    if (!this.hasPart(relationshipPart)) {
      this.addTextPart(
        relationshipPart,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${relationshipXml}
</Relationships>`
      );
      return;
    }

    const xml = await this.readText(relationshipPart);
    const existing = findStartTags(xml, "Relationship").find(
      (tag) => tag.attributes.Id === relationship.id
    );
    if (existing !== undefined) {
      const end = existing.selfClosing ? existing.end : findElementEnd(xml, existing);
      this.setText(
        relationshipPart,
        `${xml.slice(0, existing.start)}${relationshipXml}${xml.slice(end)}`
      );
      return;
    }

    const relationships = findStartTags(xml, "Relationships")[0];
    if (relationships === undefined) {
      throw new PackageError(`Relationship part ${relationshipPart} is missing Relationships root`);
    }

    if (relationships.selfClosing) {
      const opening = relationships.raw.replace(/\/>$/, ">");
      this.setText(
        relationshipPart,
        `${xml.slice(0, relationships.start)}${opening}
  ${relationshipXml}
</${relationships.name}>${xml.slice(relationships.end)}`
      );
      return;
    }

    const insertOffset = findElementCloseStart(xml, relationships);
    this.setText(
      relationshipPart,
      `${xml.slice(0, insertOffset)}  ${relationshipXml}\n${xml.slice(insertOffset)}`
    );
  }

  async removeContentTypeOverride(partName: string): Promise<boolean> {
    if (!this.hasPart("[Content_Types].xml")) {
      return false;
    }

    const normalized = `/${normalizePartName(partName)}`;
    const xml = await this.readText("[Content_Types].xml");
    const override = findStartTags(xml, "Override").find(
      (tag) => tag.attributes.PartName === normalized
    );

    if (override === undefined) {
      return false;
    }

    const end = override.selfClosing ? override.end : findElementEnd(xml, override);
    this.setText("[Content_Types].xml", `${xml.slice(0, override.start)}${xml.slice(end)}`);
    return true;
  }

  async rootRelationships(): Promise<Relationship[]> {
    if (!this.hasPart("_rels/.rels")) {
      return [];
    }

    return parseRelationships(await this.readText("_rels/.rels"));
  }

  async inspect(): Promise<PackageInspectResult> {
    const relationships: Record<string, Relationship[]> = {};

    if (this.hasPart("_rels/.rels")) {
      relationships["/"] = await this.rootRelationships();
    }

    for (const partName of this.listParts()) {
      if (partName.endsWith(".rels")) {
        continue;
      }

      const rels = await this.relationshipsFor(partName);
      if (rels.length > 0) {
        relationships[partName] = rels;
      }
    }

    return {
      parts: this.listParts(),
      relationships
    };
  }

  async write(): Promise<Uint8Array> {
    const entries = [...this.parts.values()].map((part) => {
      if (part.dirtyText === undefined) {
        if (part.entry === undefined) {
          throw new PackageError(`Part ${part.name} has no payload`);
        }

        return {
          name: part.name,
          compressedData: part.entry.compressedData,
          compressionMethod: part.entry.compressionMethod,
          crc32: part.entry.crc32,
          uncompressedSize: part.entry.uncompressedSize,
          lastModDate: part.entry.lastModDate,
          lastModTime: part.entry.lastModTime,
          externalAttributes: part.entry.externalAttributes
        };
      }

      return {
        name: part.name,
        data: textEncoder.encode(part.dirtyText),
        compressionMethod: 8 as const,
        ...(part.entry === undefined
          ? {}
          : {
              lastModDate: part.entry.lastModDate,
              lastModTime: part.entry.lastModTime,
              externalAttributes: part.entry.externalAttributes
            })
      };
    });

    return writeZip(entries, this.compression);
  }

  private requirePart(partName: string): PackagePart {
    const normalized = normalizePartName(partName);
    const part = this.parts.get(normalized);

    if (part === undefined) {
      throw new PackageError(`Missing package part ${normalized}`);
    }

    if (part.entry === undefined && part.dirtyText === undefined) {
      throw new PackageError(`Missing package part data ${normalized}`);
    }

    return part;
  }
}

export function parseRelationships(xml: string): Relationship[] {
  return findStartTags(xml, "Relationship").map((tag) =>
    relationshipFromAttributes(tag.attributes)
  );
}

export function normalizePartName(partName: string): string {
  return partName.replace(/^\/+/, "").replaceAll("\\", "/");
}

export function resolveRelationshipTarget(sourcePartName: string, target: string): string {
  if (target.startsWith("/")) {
    return normalizePartName(target);
  }

  const source = normalizePartName(sourcePartName);
  const base = source.includes("/") ? source.slice(0, source.lastIndexOf("/") + 1) : "";
  const segments = `${base}${target}`.split("/");
  const resolved: string[] = [];

  for (const segment of segments) {
    if (segment.length === 0 || segment === ".") {
      continue;
    }

    if (segment === "..") {
      resolved.pop();
      continue;
    }

    resolved.push(segment);
  }

  return resolved.join("/");
}

export function relationshipPartName(partName: string): string {
  const normalized = normalizePartName(partName);
  const slash = normalized.lastIndexOf("/");

  if (slash === -1) {
    return `_rels/${normalized}.rels`;
  }

  return `${normalized.slice(0, slash + 1)}_rels/${normalized.slice(slash + 1)}.rels`;
}

function relationshipElementXml(relationship: Relationship): string {
  const attributes = [
    `Id="${escapeXmlAttribute(relationship.id)}"`,
    `Type="${escapeXmlAttribute(relationship.type)}"`,
    `Target="${escapeXmlAttribute(relationship.target)}"`,
    relationship.targetMode === undefined
      ? undefined
      : `TargetMode="${escapeXmlAttribute(relationship.targetMode)}"`
  ]
    .filter((attribute): attribute is string => attribute !== undefined)
    .join(" ");

  return `<Relationship ${attributes}/>`;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function requireAttribute(attributes: Record<string, string>, name: string): string {
  const value = attributes[name];
  if (value === undefined) {
    throw new PackageError(`Missing XML attribute ${name}`);
  }

  return value;
}

function relationshipFromAttributes(attributes: Record<string, string>): Relationship {
  const relationship: Relationship = {
    id: requireAttribute(attributes, "Id"),
    type: requireAttribute(attributes, "Type"),
    target: requireAttribute(attributes, "Target")
  };

  if (attributes.TargetMode !== undefined) {
    relationship.targetMode = attributes.TargetMode;
  }

  return relationship;
}

export function contentTypeOverrideXml(partName: string, contentType: string): string {
  return `<Override PartName="/${escapeXmlAttribute(normalizePartName(partName))}" ContentType="${escapeXmlAttribute(contentType)}"/>`;
}
