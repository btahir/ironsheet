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

export type XmlToken =
  | { kind: "text"; start: number; end: number; text: string }
  | { kind: "start"; tag: XmlTag }
  | {
      kind: "end";
      name: string;
      localName: string;
      start: number;
      end: number;
      raw: string;
    }
  | { kind: "comment"; start: number; end: number; raw: string; text: string }
  | { kind: "cdata"; start: number; end: number; raw: string; text: string }
  | { kind: "processingInstruction"; start: number; end: number; raw: string }
  | { kind: "declaration"; start: number; end: number; raw: string };

export type XmlChunkTransform = (
  token: XmlToken
) => string | undefined | Promise<string | undefined>;

export function* tokenizeXml(xml: string, options: { start?: number } = {}): Generator<XmlToken> {
  let offset = options.start ?? 0;

  while (offset < xml.length) {
    const tagStart = xml.indexOf("<", offset);
    if (tagStart === -1) {
      if (offset < xml.length) {
        yield { kind: "text", start: offset, end: xml.length, text: xml.slice(offset) };
      }
      return;
    }

    if (tagStart > offset) {
      yield { kind: "text", start: offset, end: tagStart, text: xml.slice(offset, tagStart) };
    }

    if (xml.startsWith("<?", tagStart)) {
      const end = findDelimitedEnd(xml, "?>", tagStart, "processing instruction") + 1;
      yield {
        kind: "processingInstruction",
        start: tagStart,
        end,
        raw: xml.slice(tagStart, end)
      };
      offset = end;
      continue;
    }

    if (xml.startsWith("<!--", tagStart)) {
      const end = findDelimitedEnd(xml, "-->", tagStart, "comment") + 1;
      yield {
        kind: "comment",
        start: tagStart,
        end,
        raw: xml.slice(tagStart, end),
        text: xml.slice(tagStart + "<!--".length, end - "-->".length)
      };
      offset = end;
      continue;
    }

    if (xml.startsWith("<![CDATA[", tagStart)) {
      const end = findDelimitedEnd(xml, "]]>", tagStart, "CDATA section") + 1;
      yield {
        kind: "cdata",
        start: tagStart,
        end,
        raw: xml.slice(tagStart, end),
        text: xml.slice(tagStart + "<![CDATA[".length, end - "]]>".length)
      };
      offset = end;
      continue;
    }

    const tagEnd = findTagEnd(xml, tagStart) + 1;
    const raw = xml.slice(tagStart, tagEnd);

    if (raw.startsWith("</")) {
      const name = parseEndTagName(raw);
      if (name !== undefined) {
        yield {
          kind: "end",
          name,
          localName: toLocalName(name),
          start: tagStart,
          end: tagEnd,
          raw
        };
      }
      offset = tagEnd;
      continue;
    }

    if (raw.startsWith("<!")) {
      yield { kind: "declaration", start: tagStart, end: tagEnd, raw };
      offset = tagEnd;
      continue;
    }

    const tag = parseStartTag(raw, tagStart, tagEnd);
    if (tag !== undefined) {
      yield { kind: "start", tag };
    }

    offset = tagEnd;
  }
}

export async function* tokenizeXmlChunks(
  chunks: Iterable<string> | AsyncIterable<string>
): AsyncGenerator<XmlToken> {
  let buffer = "";
  let bufferStart = 0;

  function consume(length: number): void {
    buffer = buffer.slice(length);
    bufferStart += length;
  }

  function* drain(done: boolean): Generator<XmlToken> {
    while (buffer.length > 0) {
      const tagStart = buffer.indexOf("<");
      if (tagStart === -1) {
        yield {
          kind: "text",
          start: bufferStart,
          end: bufferStart + buffer.length,
          text: buffer
        };
        consume(buffer.length);
        return;
      }

      if (tagStart > 0) {
        yield {
          kind: "text",
          start: bufferStart,
          end: bufferStart + tagStart,
          text: buffer.slice(0, tagStart)
        };
        consume(tagStart);
        continue;
      }

      const parsed = parseBufferedToken(buffer, bufferStart, done);
      if (parsed === undefined) {
        return;
      }

      consume(parsed.consumed);
      if (parsed.token !== undefined) {
        yield parsed.token;
      }
    }
  }

  for await (const chunk of chunks) {
    if (chunk.length === 0) {
      continue;
    }

    buffer += chunk;
    yield* drain(false);
  }

  yield* drain(true);
}

export async function* transformXmlChunks(
  chunks: Iterable<string> | AsyncIterable<string>,
  transform: XmlChunkTransform
): AsyncGenerator<string> {
  for await (const token of tokenizeXmlChunks(chunks)) {
    yield (await transform(token)) ?? tokenRawText(token);
  }
}

export function findStartTags(xml: string, localName: string): XmlTag[] {
  const tags: XmlTag[] = [];

  for (const token of tokenizeXml(xml)) {
    if (token.kind === "start" && token.tag.localName === localName) {
      tags.push(token.tag);
    }
  }

  return tags;
}

export function findFirstStartTag(xml: string, localName: string): XmlTag | undefined {
  for (const token of tokenizeXml(xml)) {
    if (token.kind === "start" && token.tag.localName === localName) {
      return token.tag;
    }
  }

  return undefined;
}

export function findElementEnd(xml: string, tag: XmlTag): number {
  if (tag.selfClosing) {
    return tag.end;
  }

  return findElementCloseTag(xml, tag).end;
}

export function findElementCloseStart(xml: string, tag: XmlTag): number {
  if (tag.selfClosing) {
    return tag.end;
  }

  return findElementCloseTag(xml, tag).start;
}

function findElementCloseTag(xml: string, tag: XmlTag): { start: number; end: number } {
  let depth = 1;

  for (const token of tokenizeXml(xml, { start: tag.end })) {
    if (token.kind === "end" && token.localName === tag.localName) {
      depth -= 1;
      if (depth === 0) {
        return { start: token.start, end: token.end };
      }
      continue;
    }

    if (token.kind === "start" && token.tag.localName === tag.localName && !token.tag.selfClosing) {
      depth += 1;
    }
  }

  throw new PackageError(`Element ${tag.name} is missing a closing tag`);
}

function parseBufferedToken(
  source: string,
  absoluteStart: number,
  done: boolean
): { token: XmlToken | undefined; consumed: number } | undefined {
  if (source.startsWith("<?")) {
    const end = findBufferedDelimitedEnd(
      source,
      "?>",
      absoluteStart,
      "processing instruction",
      done
    );
    if (end === undefined) {
      return undefined;
    }

    return {
      consumed: end,
      token: {
        kind: "processingInstruction",
        start: absoluteStart,
        end: absoluteStart + end,
        raw: source.slice(0, end)
      }
    };
  }

  if (source.startsWith("<!--")) {
    const end = findBufferedDelimitedEnd(source, "-->", absoluteStart, "comment", done);
    if (end === undefined) {
      return undefined;
    }

    return {
      consumed: end,
      token: {
        kind: "comment",
        start: absoluteStart,
        end: absoluteStart + end,
        raw: source.slice(0, end),
        text: source.slice("<!--".length, end - "-->".length)
      }
    };
  }

  if (source.startsWith("<![CDATA[")) {
    const end = findBufferedDelimitedEnd(source, "]]>", absoluteStart, "CDATA section", done);
    if (end === undefined) {
      return undefined;
    }

    return {
      consumed: end,
      token: {
        kind: "cdata",
        start: absoluteStart,
        end: absoluteStart + end,
        raw: source.slice(0, end),
        text: source.slice("<![CDATA[".length, end - "]]>".length)
      }
    };
  }

  const tagEnd = findBufferedTagEnd(source, absoluteStart, done);
  if (tagEnd === undefined) {
    return undefined;
  }

  const raw = source.slice(0, tagEnd);
  if (raw.startsWith("</")) {
    const name = parseEndTagName(raw);
    return {
      consumed: tagEnd,
      token:
        name === undefined
          ? undefined
          : {
              kind: "end",
              name,
              localName: toLocalName(name),
              start: absoluteStart,
              end: absoluteStart + tagEnd,
              raw
            }
    };
  }

  if (raw.startsWith("<!")) {
    return {
      consumed: tagEnd,
      token: {
        kind: "declaration",
        start: absoluteStart,
        end: absoluteStart + tagEnd,
        raw
      }
    };
  }

  const tag = parseStartTag(raw, absoluteStart, absoluteStart + tagEnd);
  return {
    consumed: tagEnd,
    token: tag === undefined ? undefined : { kind: "start", tag }
  };
}

function findBufferedDelimitedEnd(
  source: string,
  delimiter: string,
  absoluteStart: number,
  description: string,
  done: boolean
): number | undefined {
  const end = source.indexOf(delimiter, delimiter.length);
  if (end === -1) {
    if (done) {
      throw new PackageError(`Unterminated XML ${description} at byte ${absoluteStart}`);
    }

    return undefined;
  }

  return end + delimiter.length;
}

function findBufferedTagEnd(
  source: string,
  absoluteStart: number,
  done: boolean
): number | undefined {
  let quote: string | undefined;

  for (let offset = 1; offset < source.length; offset += 1) {
    const char = source[offset];

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
      return offset + 1;
    }
  }

  if (done) {
    throw new PackageError(`Unterminated XML tag at byte ${absoluteStart}`);
  }

  return undefined;
}

function tokenRawText(token: XmlToken): string {
  switch (token.kind) {
    case "text":
      return token.text;
    case "start":
      return token.tag.raw;
    default:
      return token.raw;
  }
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

    if (source[offset] !== quote) {
      throw new PackageError(`Unterminated XML attribute ${name}`);
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

function parseEndTagName(raw: string): string | undefined {
  if (!raw.startsWith("</")) {
    return undefined;
  }

  const name = raw.slice(2, raw.endsWith(">") ? -1 : undefined).trim();
  return name.length === 0 ? undefined : name;
}

function findDelimitedEnd(
  xml: string,
  delimiter: string,
  start: number,
  description: string
): number {
  const end = xml.indexOf(delimiter, start + delimiter.length);
  if (end === -1) {
    throw new PackageError(`Unterminated XML ${description} at byte ${start}`);
  }

  return end + delimiter.length - 1;
}
