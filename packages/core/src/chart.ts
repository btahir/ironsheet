import type { OoxmlPackage } from "./opc.ts";
import { decodeXml, escapeXmlText, findElementCloseStart, findStartTags } from "./xml.ts";

export type ChartFormulaRetarget = {
  from: string;
  to: string;
};

export type WorkbookChart = {
  partName: string;
  formulas: string[];
};

export async function listWorkbookCharts(pkg: OoxmlPackage): Promise<WorkbookChart[]> {
  const charts: WorkbookChart[] = [];

  for (const partName of pkg.listParts().filter((part) => /^xl\/charts\/.+\.xml$/.test(part))) {
    charts.push({
      partName,
      formulas: chartFormulaTexts(await pkg.readText(partName))
    });
  }

  return charts;
}

export async function retargetWorkbookChartFormulas(
  pkg: OoxmlPackage,
  retargets: ChartFormulaRetarget[]
): Promise<number> {
  let changed = 0;
  const replacementMap = new Map(retargets.map((retarget) => [retarget.from, retarget.to]));

  for (const partName of pkg.listParts().filter((part) => /^xl\/charts\/.+\.xml$/.test(part))) {
    const xml = await pkg.readText(partName);
    const result = retargetChartFormulaXml(xml, replacementMap);
    if (result.xml !== xml) {
      pkg.setText(partName, result.xml);
      changed += result.changed;
    }
  }

  return changed;
}

export function retargetChartFormulaXml(
  xml: string,
  replacementMap: Map<string, string>
): { changed: number; xml: string } {
  let changed = 0;
  let offset = 0;
  let result = "";

  for (const formula of findStartTags(xml, "f")) {
    if (formula.selfClosing || formula.start < offset) {
      continue;
    }

    const textStart = formula.end;
    const textEnd = findElementCloseStart(xml, formula);
    const text = decodeXml(xml.slice(textStart, textEnd));
    const replacement = replacementMap.get(text);
    if (replacement === undefined) {
      continue;
    }

    result += xml.slice(offset, textStart);
    result += escapeXmlText(replacement);
    offset = textEnd;
    changed += 1;
  }

  return {
    changed,
    xml: changed === 0 ? xml : result + xml.slice(offset)
  };
}

function chartFormulaTexts(xml: string): string[] {
  return findStartTags(xml, "f")
    .filter((formula) => !formula.selfClosing)
    .map((formula) => decodeXml(xml.slice(formula.end, findElementCloseStart(xml, formula))));
}
