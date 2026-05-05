import { crc32, writeZip, type ZipWriteEntry } from "../../packages/core/src/index.ts";

const textEncoder = new TextEncoder();

export type MinimalWorkbookOptions = {
  includeCalcChain?: boolean;
  includeConditionalFormatting?: boolean;
  includeDataValidation?: boolean;
  includeDefinedName?: boolean;
  includeDrawing?: boolean;
  includeHiddenSheet?: boolean;
  includeHyperlink?: boolean;
  includeMacro?: boolean;
  includeMerge?: boolean;
  includeTable?: boolean;
  styledTableBody?: boolean;
  tableRows?: Array<[string, number]>;
  useSharedStrings?: boolean;
};

export async function createMinimalWorkbook(
  options: MinimalWorkbookOptions = {}
): Promise<Uint8Array> {
  const tableRows: Array<[string, number]> = options.tableRows ?? [["Old", 1]];
  const tableEndRow = 1 + tableRows.length;
  const tableRef = `A1:B${tableEndRow}`;
  const sheetRelationships = sheetRelationshipsXml(options);
  const entries: ZipWriteEntry[] = [
    textPart(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="${options.includeMacro === true ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"}"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  ${options.includeHiddenSheet === true ? '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' : ""}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${options.includeTable === true ? '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' : ""}
  ${options.includeDrawing === true ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : ""}
  ${options.includeDrawing === true ? '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' : ""}
  ${options.includeMacro === true ? '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>' : ""}
  ${options.useSharedStrings === true ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' : ""}
  ${options.includeCalcChain === true ? '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>' : ""}
  ${options.includeDrawing === true ? '<Default Extension="png" ContentType="image/png"/>' : ""}
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
    ${options.includeHiddenSheet === true ? '<sheet name="HiddenData" sheetId="2" state="hidden" r:id="rIdHidden"/>' : ""}
  </sheets>
  ${options.includeDefinedName === true ? '<definedNames><definedName name="RevenueRange" comment="Template output range">Sheet1!$A$1:$B$2</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0" hidden="1">Sheet1!$1:$1</definedName></definedNames>' : ""}
</workbook>`
    ),
    textPart(
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  ${options.includeHiddenSheet === true ? '<Relationship Id="rIdHidden" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' : ""}
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  ${options.includeMacro === true ? '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>' : ""}
  ${options.useSharedStrings === true ? '<Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' : ""}
  ${options.includeCalcChain === true ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>' : ""}
</Relationships>`
    ),
    ...(sheetRelationships.length > 0
      ? [
          textPart(
            "xl/worksheets/_rels/sheet1.xml.rels",
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheetRelationships}
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
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`
    ),
    textPart(
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${options.includeTable === true ? tableRef : "A1:A1"}"/>
  <sheetData>
    ${options.includeTable === true ? tableSheetRows(tableRows, options.styledTableBody === true) : `<row r="1">${options.useSharedStrings === true ? '<c r="A1" s="1" t="s"><v>0</v></c>' : '<c r="A1" s="1" t="inlineStr"><is><t>Original</t></is></c>'}</row>`}
  </sheetData>
  ${options.includeConditionalFormatting === true ? '<conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>10</formula></cfRule></conditionalFormatting>' : ""}
  ${options.includeDataValidation === true ? '<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="1" sqref="B2:B10"><formula1>0</formula1><formula2>100</formula2></dataValidation></dataValidations>' : ""}
  ${options.includeMerge === true ? '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>' : ""}
  ${options.includeHyperlink === true ? '<hyperlinks><hyperlink ref="A1" r:id="rIdHyperlink1"/></hyperlinks>' : ""}
  ${options.includeDrawing === true ? '<drawing r:id="rIdDrawing1"/>' : ""}
  ${options.includeTable === true ? '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>' : ""}
</worksheet>`
    )
  ];

  if (options.includeHiddenSheet === true) {
    entries.push(
      textPart(
        "xl/worksheets/sheet2.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Hidden</t></is></c></row></sheetData>
</worksheet>`
      )
    );
  }

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

  if (options.includeDrawing === true) {
    entries.push(
      textPart(
        "xl/drawings/drawing1.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:row>1</xdr:row></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:row>8</xdr:row></xdr:to>
    <xdr:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart1"/></a:graphicData></a:graphic></xdr:graphicFrame>
    <xdr:pic><xdr:blipFill><a:blip r:embed="rIdImage1"/></xdr:blipFill></xdr:pic>
  </xdr:twoCellAnchor>
</xdr:wsDr>`
      ),
      textPart(
        "xl/drawings/_rels/drawing1.xml.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`
      ),
      textPart(
        "xl/charts/chart1.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`
      ),
      {
        name: "xl/media/image1.png",
        data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        compressionMethod: 0,
        crc32: crc32(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])),
        uncompressedSize: 8
      }
    );
  }

  return writeZip(entries);
}

function sheetRelationshipsXml(options: MinimalWorkbookOptions): string {
  return [
    options.includeTable === true
      ? '<Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'
      : undefined,
    options.includeHyperlink === true
      ? '<Relationship Id="rIdHyperlink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>'
      : undefined,
    options.includeDrawing === true
      ? '<Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
      : undefined
  ]
    .filter((relationship): relationship is string => relationship !== undefined)
    .join("\n  ");
}

function tableSheetRows(rows: Array<[string, number]>, styledBody: boolean): string {
  const bodyRows = rows
    .map(([name, amount], index) => {
      const rowNumber = index + 2;
      const rowAttributes = styledBody ? ' s="2" customFormat="1"' : "";
      const nameStyle = styledBody ? ' s="3"' : "";
      const amountStyle = styledBody ? ' s="4"' : "";
      return `<row r="${rowNumber}"${rowAttributes}><c r="A${rowNumber}"${nameStyle} t="inlineStr"><is><t>${name}</t></is></c><c r="B${rowNumber}"${amountStyle}><v>${amount}</v></c></row>`;
    })
    .join("");

  return `<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>${bodyRows}`;
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
