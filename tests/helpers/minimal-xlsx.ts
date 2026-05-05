import { crc32, writeZip, type ZipWriteEntry } from "../../packages/core/src/index.ts";

const textEncoder = new TextEncoder();

export type MinimalWorkbookOptions = {
  includeCalcChain?: boolean;
  includeMacro?: boolean;
  includeTable?: boolean;
  useSharedStrings?: boolean;
};

export async function createMinimalWorkbook(
  options: MinimalWorkbookOptions = {}
): Promise<Uint8Array> {
  const entries: ZipWriteEntry[] = [
    textPart(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="${options.includeMacro === true ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"}"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${options.includeTable === true ? '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' : ""}
  ${options.includeMacro === true ? '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>' : ""}
  ${options.useSharedStrings === true ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' : ""}
  ${options.includeCalcChain === true ? '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>' : ""}
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
</workbook>`
    ),
    textPart(
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  ${options.includeMacro === true ? '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>' : ""}
  ${options.useSharedStrings === true ? '<Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' : ""}
  ${options.includeCalcChain === true ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>' : ""}
</Relationships>`
    ),
    ...(options.includeTable === true
      ? [
          textPart(
            "xl/worksheets/_rels/sheet1.xml.rels",
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`
          )
        ]
      : []),
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
  <dimension ref="${options.includeTable === true ? "A1:B2" : "A1:A1"}"/>
  <sheetData>
    ${options.includeTable === true ? '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Old</t></is></c><c r="B2"><v>1</v></c></row>' : `<row r="1">${options.useSharedStrings === true ? '<c r="A1" s="1" t="s"><v>0</v></c>' : '<c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c>'}</row>`}
  </sheetData>
  ${options.includeTable === true ? '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>' : ""}
</worksheet>`
    )
  ];

  if (options.includeCalcChain === true) {
    entries.push(
      textPart(
        "xl/calcChain.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <c r="A1" i="1"/>
</calcChain>`
      )
    );
  }

  if (options.useSharedStrings === true) {
    entries.push(
      textPart(
        "xl/sharedStrings.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Original</t></si>
</sst>`
      )
    );
  }

  if (options.includeMacro === true) {
    entries.push({
      name: "xl/vbaProject.bin",
      data: new Uint8Array([0xca, 0xfe, 0xba, 0xbe, 0x00, 0x01]),
      compressionMethod: 0,
      crc32: crc32(new Uint8Array([0xca, 0xfe, 0xba, 0xbe, 0x00, 0x01])),
      uncompressedSize: 6
    });
  }

  if (options.includeTable === true) {
    entries.push(
      textPart(
        "xl/tables/table1.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="RevenueTable" displayName="RevenueTable" ref="A1:B2" totalsRowShown="0">
  <autoFilter ref="A1:B2"/>
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
