export type FormulaSheetReference = {
  sheetName: string;
};

export function parseFormulaSheetReferences(formula: string): FormulaSheetReference[] {
  const references = new Set<string>();
  const scrubbed = stripDoubleQuotedStrings(formula);
  const pattern = /(?:^|[,( +\-*/^&=<>])((?:'(?:(?:'')|[^'])+'|[A-Za-z_][A-Za-z0-9_ .]*))!/g;

  for (const match of scrubbed.matchAll(pattern)) {
    const rawName = match[1];
    if (rawName === undefined || rawName.includes("[")) {
      continue;
    }

    references.add(unquoteSheetName(rawName));
  }

  return [...references].map((sheetName) => ({ sheetName }));
}

function stripDoubleQuotedStrings(formula: string): string {
  let result = "";
  let offset = 0;

  while (offset < formula.length) {
    const char = formula[offset];
    if (char !== '"') {
      result += char;
      offset += 1;
      continue;
    }

    result += " ";
    offset += 1;
    while (offset < formula.length) {
      if (formula[offset] === '"') {
        if (formula[offset + 1] === '"') {
          result += "  ";
          offset += 2;
          continue;
        }

        result += " ";
        offset += 1;
        break;
      }

      result += " ";
      offset += 1;
    }
  }

  return result;
}

function unquoteSheetName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replaceAll("''", "'");
  }

  return trimmed;
}
