import { PackageError } from "./errors.ts";
import { findStartTags } from "./xml.ts";
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
  entry: ZipEntry;
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

  async relationshipsFor(partName: string): Promise<Relationship[]> {
    const normalized = normalizePartName(partName);
    const relationshipPart = relationshipPartName(normalized);

    if (!this.hasPart(relationshipPart)) {
      return [];
    }

    return parseRelationships(await this.readText(relationshipPart));
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
        lastModDate: part.entry.lastModDate,
        lastModTime: part.entry.lastModTime,
        externalAttributes: part.entry.externalAttributes
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

    return part;
  }
}

export function parseRelationships(xml: string): Relationship[] {
  return findStartTags(xml, "Relationship").map((tag) => {
    const relationship: Relationship = {
      id: requireAttribute(tag.attributes, "Id"),
      type: requireAttribute(tag.attributes, "Type"),
      target: requireAttribute(tag.attributes, "Target")
    };

    if (tag.attributes.TargetMode !== undefined) {
      relationship.targetMode = tag.attributes.TargetMode;
    }

    return relationship;
  });
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

function relationshipPartName(partName: string): string {
  const normalized = normalizePartName(partName);
  const slash = normalized.lastIndexOf("/");

  if (slash === -1) {
    return `_rels/${normalized}.rels`;
  }

  return `${normalized.slice(0, slash + 1)}_rels/${normalized.slice(slash + 1)}.rels`;
}

function requireAttribute(attributes: Record<string, string>, name: string): string {
  const value = attributes[name];
  if (value === undefined) {
    throw new PackageError(`Missing XML attribute ${name}`);
  }

  return value;
}
