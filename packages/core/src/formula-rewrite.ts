import { decodeXml, escapeXmlText, findElementCloseStart, findStartTags } from "./xml.ts";
import type { XmlTag } from "./xml.ts";

const formulaElementNames = new Set([
  "calculatedColumnFormula",
  "definedName",
  "f",
  "formula",
  "formula1",
  "formula2",
  "totalsRowFormula"
]);

export function rewriteFormulaElements(
  xml: string,
  rewriteFormula: (formula: string) => string
): string {
  const formulaTags = findFormulaElementTags(xml);
  if (formulaTags.length === 0) {
    return xml;
  }

  let result = "";
  let offset = 0;
  let changed = false;

  for (const tag of formulaTags) {
    if (tag.selfClosing || tag.start < offset) {
      continue;
    }

    const textStart = tag.end;
    const textEnd = findElementCloseStart(xml, tag);
    const rawText = xml.slice(textStart, textEnd);
    if (rawText.includes("<")) {
      continue;
    }

    const formula = decodeXml(rawText);
    const rewritten = rewriteFormula(formula);
    if (rewritten === formula) {
      continue;
    }

    result += xml.slice(offset, textStart);
    result += escapeXmlText(rewritten);
    offset = textEnd;
    changed = true;
  }

  if (!changed) {
    return xml;
  }

  return result + xml.slice(offset);
}

function findFormulaElementTags(xml: string): XmlTag[] {
  return [...formulaElementNames]
    .flatMap((name) => findStartTags(xml, name))
    .sort((left, right) => left.start - right.start);
}
