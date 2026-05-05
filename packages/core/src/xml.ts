import { PackageError } from "./errors.ts";

export type XmlTag = {
  name: string;
  localName: string;
  attributes: Record<string, string>;
  start: number;
  end: number;
  selfClosing: boolean;
  raw: string;
};

export function findStartTags(xml: string, localName: string): XmlTag[] {
  const tags: XmlTag[] = [];
  let offset = 0;

  while (offset < xml.length) {
    const start = xml.indexOf("<", offset);
    if (start === -1) {
      break;
    }

    if (isSpecialTag(xml, start)) {
      offset = start + 1;
      continue;
    }

    const end = findTagEnd(xml, start);
    const raw = xml.slice(start, end + 1);
    const tag = parseStartTag(raw, start, end + 1);
    if (tag !== undefined && tag.localName === localName) {
      tags.push(tag);
    }

    offset = end + 1;
  }

  return tags;
}

export function findFirstStartTag(xml: string, localName: string): XmlTag | undefined {
  return findStartTags(xml, localName)[0];
}

export function parseStartTag(raw: string, start = 0, end = raw.length): XmlTag | undefined {
  if (!raw.startsWith("<") || raw.startsWith("</")) {
    return undefined;
  }

  const body = raw.slice(1, raw.endsWith(">") ? -1 : undefined).trim();
  if (body.length === 0 || body.startsWith("?") || body.startsWith("!")) {
    return undefined;
  }

  const selfClosing = body.endsWith("/");
  const content = selfClosing ? body.slice(0, -1).trimEnd() : body;
  const nameEnd = findNameEnd(content);
  const name = content.slice(0, nameEnd);
  const attributes = parseAttributes(content.slice(nameEnd));

  return {
    name,
    localName: toLocalName(name),
    attributes,
    start,
    end,
    selfClosing,
    raw
  };
}

export function parseAttributes(source: string): Record<string, string> {
  const attributes: Record<string, string> = {};
  let offset = 0;

  while (offset < source.length) {
    while (/\s/.test(source[offset] ?? "")) {
      offset += 1;
    }

    if (offset >= source.length) {
      break;
    }

    const nameStart = offset;
    while (offset < source.length && !/[\s=]/.test(source[offset] ?? "")) {
      offset += 1;
    }

    const name = source.slice(nameStart, offset);
    while (/\s/.test(source[offset] ?? "")) {
      offset += 1;
    }

    if (source[offset] !== "=") {
      attributes[name] = "";
      continue;
    }

    offset += 1;
    while (/\s/.test(source[offset] ?? "")) {
      offset += 1;
    }

    const quote = source[offset];
    if (quote !== '"' && quote !== "'") {
      throw new PackageError(`Invalid XML attribute ${name}`);
    }

    offset += 1;
    const valueStart = offset;
    while (offset < source.length && source[offset] !== quote) {
      offset += 1;
    }

    attributes[name] = decodeXml(source.slice(valueStart, offset));
    offset += 1;
  }

  return attributes;
}

export function escapeXmlText(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

export function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replaceAll('"', "&quot;").replaceAll("'", "&apos;");
}

export function decodeXml(value: string): string {
  return value
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&");
}

export function toLocalName(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function findTagEnd(xml: string, start: number): number {
  let quote: string | undefined;

  for (let offset = start + 1; offset < xml.length; offset += 1) {
    const char = xml[offset];

    if (quote !== undefined) {
      if (char === quote) {
        quote = undefined;
      }
      continue;
    }

    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }

    if (char === ">") {
      return offset;
    }
  }

  throw new PackageError(`Unterminated XML tag at byte ${start}`);
}

function findNameEnd(source: string): number {
  let offset = 0;

  while (offset < source.length && !/\s/.test(source[offset] ?? "")) {
    offset += 1;
  }

  return offset;
}

function isSpecialTag(xml: string, start: number): boolean {
  return xml.startsWith("</", start) || xml.startsWith("<?", start) || xml.startsWith("<!", start);
}
