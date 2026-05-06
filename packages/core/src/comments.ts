import { parseCellAddress } from "./address.ts";
import { decodeXml, findElementCloseStart, findElementEnd, findStartTags } from "./xml.ts";

export const worksheetCommentsRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

export type WorkbookComment = {
  sheetName: string;
  sheetPartName: string;
  commentPartName: string;
  relationshipId: string;
  ref: string;
  text: string;
  author?: string;
  authorId?: string;
  rawXml: string;
};

export type WorksheetComment = {
  ref: string;
  text: string;
  author?: string;
  authorId?: string;
  rawXml: string;
};

export function parseWorksheetComments(xml: string): WorksheetComment[] {
  const authors = parseCommentAuthors(xml);

  return findStartTags(xml, "comment").map((tag) => {
    const ref = parseCellAddress(tag.attributes.ref ?? "").address;
    const authorId = tag.attributes.authorId;
    const rawXml = xml.slice(tag.start, tag.selfClosing ? tag.end : findElementEnd(xml, tag));
    const author =
      authorId === undefined ? undefined : (authors[Number.parseInt(authorId, 10)] ?? undefined);

    return {
      ref,
      text: readCommentText(rawXml),
      ...(author === undefined ? {} : { author }),
      ...(authorId === undefined ? {} : { authorId }),
      rawXml
    };
  });
}

function parseCommentAuthors(xml: string): string[] {
  return findStartTags(xml, "author").map((tag) => {
    if (tag.selfClosing) {
      return "";
    }

    return decodeXml(xml.slice(tag.end, findElementCloseStart(xml, tag)));
  });
}

function readCommentText(commentXml: string): string {
  return findStartTags(commentXml, "t")
    .map((tag) =>
      tag.selfClosing
        ? ""
        : decodeXml(commentXml.slice(tag.end, findElementCloseStart(commentXml, tag)))
    )
    .join("");
}
