import { decodeXml, findStartTags } from "./xml.ts";

export function parseSharedStrings(xml: string): string[] {
  const items = findElementRanges(xml, "si");

  return items.map((item) => {
    const textRuns = findElementRanges(item, "t");
    return textRuns.map((run) => decodeXml(stripOuterTag(run))).join("");
  });
}

function findElementRanges(xml: string, localName: string): string[] {
  return findStartTags(xml, localName).map((tag) => {
    if (tag.selfClosing) {
      return tag.raw;
    }

    const closePattern = new RegExp(`</(?:[A-Za-z0-9_]+:)?${localName}>`, "g");
    closePattern.lastIndex = tag.end;
    const close = closePattern.exec(xml);

    if (close === null || close.index === undefined) {
      return tag.raw;
    }

    return xml.slice(tag.start, close.index + close[0].length);
  });
}

function stripOuterTag(xml: string): string {
  const start = xml.indexOf(">");
  const end = xml.lastIndexOf("<");

  if (start === -1 || end === -1 || end <= start) {
    return "";
  }

  return xml.slice(start + 1, end);
}
