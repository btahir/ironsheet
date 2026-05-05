import { WorkbookError } from "./errors.ts";
import { decodeXml, findStartTags, type XmlTag } from "./xml.ts";

export type WorkbookDefinedName = {
  name: string;
  text: string;
  comment?: string;
  hidden?: boolean;
  localSheetId?: string;
};

export function parseDefinedNames(workbookXml: string): WorkbookDefinedName[] {
  return findStartTags(workbookXml, "definedName").map((tag) => {
    const name = tag.attributes.name;
    if (name === undefined) {
      throw new WorkbookError("Defined name is missing name attribute");
    }

    return {
      name,
      text: decodeXml(workbookXml.slice(tag.end, findDefinedNameClose(workbookXml, tag))),
      ...(tag.attributes.comment === undefined ? {} : { comment: tag.attributes.comment }),
      ...(tag.attributes.hidden === undefined
        ? {}
        : { hidden: tag.attributes.hidden === "1" || tag.attributes.hidden === "true" }),
      ...(tag.attributes.localSheetId === undefined
        ? {}
        : { localSheetId: tag.attributes.localSheetId })
    };
  });
}

function findDefinedNameClose(xml: string, tag: XmlTag): number {
  const close = /<\/(?:[A-Za-z0-9_]+:)?definedName>/.exec(xml.slice(tag.end));
  if (close?.index === undefined) {
    throw new WorkbookError(`Defined name ${tag.attributes.name ?? ""} is missing a closing tag`);
  }

  return tag.end + close.index;
}
