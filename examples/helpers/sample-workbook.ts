import { crc32, writeZip, type ZipWriteEntry } from "@ironsheet/core";

const textEncoder = new TextEncoder();

export type SampleWorkbookOptions = {
  /** Adds a RevenueTable (Name/Amount) starting at A1 plus a RevenueRange defined name. */
  includeTable?: boolean;
  /** Table body rows. Defaults to a small revenue listing. */
  tableRows?: Array<[string, number]>;
};

/**
 * Builds a tiny but valid XLSX package in memory, the way an Excel author
 * would hand you a starting workbook. Examples use this instead of a
 * committed .xlsx fixture so they are fully self-contained.
 */
export async function createSampleWorkbook(
  options: SampleWorkbookOptions = {}
): Promise<Uint8Array> {
  const includeTable = options.includeTable === true;
  const tableRows: Array<[string, number]> = options.tableRows ?? [
    ["North", 1200],
    ["South", 980]
  ];
  const tableRef = `A1:B${tableRows.length + 1}`;

  const sheetRows = includeTable
    ? [
        '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>',
        ...tableRows.map(([name, amount], index) => {
          const row = index + 2;
          return `<row r="${row}"><c r="A${row}" t="inlineStr"><is><t>${name}</t></is></c><c r="B${row}"><v>${amount}</v></c></row>`;
        })
      ].join("")
    : '<row r="1"><c r="A1" t="inlineStr"><is><t>Original</t></is></c></row>';

  const entries: ZipWriteEntry[] = [
    textPart(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${includeTable ? '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' : ""}
</Types>`
    ),
    textPart(
      "_rels/.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    ),
    textPart(
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  ${includeTable ? '<definedNames><definedName name="RevenueRange" comment="Template output range">Sheet1!$A$1:$B$2</definedName></definedNames>' : ""}
</workbook>`
    ),
    textPart(
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
    ),
    textPart(
      "xl/styles.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`
    ),
    textPart(
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${includeTable ? tableRef : "A1:A1"}"/>
  <sheetData>
    ${sheetRows}
  </sheetData>
  ${includeTable ? '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>' : ""}
</worksheet>`
    )
  ];

  if (includeTable) {
    entries.push(
      textPart(
        "xl/worksheets/_rels/sheet1.xml.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`
      ),
      textPart(
        "xl/tables/table1.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="RevenueTable" displayName="RevenueTable" ref="${tableRef}" totalsRowShown="0">
  <autoFilter ref="${tableRef}"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Amount"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`
      )
    );
  }

  return writeZip(entries);
}

function textPart(name: string, text: string): ZipWriteEntry {
  const data = textEncoder.encode(text);
  return {
    name,
    data,
    compressionMethod: 0,
    crc32: crc32(data),
    uncompressedSize: data.byteLength
  };
}
