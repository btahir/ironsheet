import type { OoxmlPackage } from "./opc.ts";
import { escapeXmlAttribute, findStartTags } from "./xml.ts";

export type PivotCacheSourceRetarget = {
  from?: {
    ref?: string;
    sheet?: string;
  };
  to: {
    ref?: string;
    sheet?: string;
  };
};

export async function retargetWorkbookPivotCacheSources(
  pkg: OoxmlPackage,
  retargets: PivotCacheSourceRetarget[]
): Promise<number> {
  let changed = 0;

  for (const partName of pkg
    .listParts()
    .filter((part) => /^xl\/pivotCache\/pivotCacheDefinition.+\.xml$/.test(part))) {
    const xml = await pkg.readText(partName);
    const result = retargetPivotCacheSourceXml(xml, retargets);
    if (result.xml !== xml) {
      pkg.setText(partName, result.xml);
      changed += result.changed;
    }
  }

  return changed;
}

export function retargetPivotCacheSourceXml(
  xml: string,
  retargets: PivotCacheSourceRetarget[]
): { changed: number; xml: string } {
  let changed = 0;
  let nextXml = xml;
  const sources = findStartTags(xml, "worksheetSource").slice().reverse();

  for (const source of sources) {
    const retarget = retargets.find((candidate) => sourceMatches(source.attributes, candidate));
    if (retarget === undefined) {
      continue;
    }

    nextXml = `${nextXml.slice(0, source.start)}${upsertAttributes(source.raw, retarget.to)}${nextXml.slice(source.end)}`;
    changed += 1;
  }

  return { changed, xml: nextXml };
}

function sourceMatches(
  attributes: Record<string, string>,
  retarget: PivotCacheSourceRetarget
): boolean {
  return (
    (retarget.from?.ref === undefined || attributes.ref === retarget.from.ref) &&
    (retarget.from?.sheet === undefined || attributes.sheet === retarget.from.sheet)
  );
}

function upsertAttributes(rawTag: string, attributes: Record<string, string | undefined>): string {
  let nextTag = rawTag;

  for (const [name, value] of Object.entries(attributes)) {
    if (value === undefined) {
      continue;
    }

    const escapedValue = escapeXmlAttribute(value);
    const pattern = new RegExp(`\\s${name}=(["']).*?\\1`);
    if (pattern.test(nextTag)) {
      nextTag = nextTag.replace(pattern, ` ${name}="${escapedValue}"`);
      continue;
    }

    const closing = nextTag.endsWith("/>") ? "/>" : ">";
    nextTag = `${nextTag.slice(0, -closing.length)} ${name}="${escapedValue}"${closing}`;
  }

  return nextTag;
}
