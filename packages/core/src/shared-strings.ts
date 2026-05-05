import { decodeXml, findElementCloseStart, findElementEnd, findStartTags } from "./xml.ts";

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

    return xml.slice(tag.start, findElementEnd(xml, tag));
  });
}

function stripOuterTag(xml: string): string {
  const tag = findStartTags(xml, "t")[0];
  if (tag === undefined || tag.selfClosing) {
    return "";
  }

  return xml.slice(tag.end, findElementCloseStart(xml, tag));
}
