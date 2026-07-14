---
title: Workbook
---

# Class: Workbook

Defined in: packages/core/src/workbook.ts:375

## Properties

### pkg

```ts
readonly pkg: OoxmlPackage;
```

Defined in: packages/core/src/workbook.ts:380

***

### workbookPart

```ts
readonly workbookPart: string;
```

Defined in: packages/core/src/workbook.ts:381

## Methods

### addSheet()

```ts
addSheet(name): Promise<WorkbookSheet>;
```

Defined in: packages/core/src/workbook.ts:1502

#### Parameters

##### name

`string`

#### Returns

`Promise`&lt;[`WorkbookSheet`](../type-aliases/WorkbookSheet.md)&gt;

***

### appendRows()

```ts
appendRows(
   sheetName, 
   rows, 
   options?): Promise<void>;
```

Defined in: packages/core/src/workbook.ts:549

#### Parameters

##### sheetName

`string`

##### rows

[`CellInput`](../type-aliases/CellInput.md)[][]

##### options?

###### startColumn?

`number`

#### Returns

`Promise`&lt;`void`&gt;

***

### appendTableColumn()

```ts
appendTableColumn(
   tableName, 
   columnName, 
   values?): Promise<WorkbookTable>;
```

Defined in: packages/core/src/workbook.ts:1476

#### Parameters

##### tableName

`string`

##### columnName

`string`

##### values?

[`CellInput`](../type-aliases/CellInput.md)[] = `[]`

#### Returns

`Promise`&lt;[`WorkbookTable`](../type-aliases/WorkbookTable.md)&gt;

***

### autoFilters()

```ts
autoFilters(sheetName?): Promise<WorkbookAutoFilter[]>;
```

Defined in: packages/core/src/workbook.ts:955

#### Parameters

##### sheetName?

`string`

#### Returns

`Promise`&lt;[`WorkbookAutoFilter`](../type-aliases/WorkbookAutoFilter.md)[]&gt;

***

### charts()

```ts
charts(): Promise<WorkbookChart[]>;
```

Defined in: packages/core/src/workbook.ts:1847

#### Returns

`Promise`&lt;[`WorkbookChart`](../type-aliases/WorkbookChart.md)[]&gt;

***

### clearCell()

```ts
clearCell(
   sheetName, 
   address, 
   options?): Promise<void>;
```

Defined in: packages/core/src/workbook.ts:480

#### Parameters

##### sheetName

`string`

##### address

`string`

##### options?

###### keepStyles?

`boolean`

#### Returns

`Promise`&lt;`void`&gt;

***

### clearRange()

```ts
clearRange(
   sheetName, 
   rangeRef, 
   options?): Promise<void>;
```

Defined in: packages/core/src/workbook.ts:488

#### Parameters

##### sheetName

`string`

##### rangeRef

`string`

##### options?

###### keepStyles?

`boolean`

#### Returns

`Promise`&lt;`void`&gt;

***

### comments()

```ts
comments(sheetName?): Promise<WorkbookComment[]>;
```

Defined in: packages/core/src/workbook.ts:973

#### Parameters

##### sheetName?

`string`

#### Returns

`Promise`&lt;[`WorkbookComment`](../type-aliases/WorkbookComment.md)[]&gt;

***

### conditionalFormats()

```ts
conditionalFormats(sheetName?): Promise<WorkbookConditionalFormat[]>;
```

Defined in: packages/core/src/workbook.ts:937

#### Parameters

##### sheetName?

`string`

#### Returns

`Promise`&lt;[`WorkbookConditionalFormat`](../type-aliases/WorkbookConditionalFormat.md)[]&gt;

***

### copySheet()

```ts
copySheet(sheetName, nextName): Promise<WorkbookSheet>;
```

Defined in: packages/core/src/workbook.ts:1515

#### Parameters

##### sheetName

`string`

##### nextName

`string`

#### Returns

`Promise`&lt;[`WorkbookSheet`](../type-aliases/WorkbookSheet.md)&gt;

***

### dataValidations()

```ts
dataValidations(sheetName?): Promise<WorkbookDataValidation[]>;
```

Defined in: packages/core/src/workbook.ts:919

#### Parameters

##### sheetName?

`string`

#### Returns

`Promise`&lt;[`WorkbookDataValidation`](../type-aliases/WorkbookDataValidation.md)[]&gt;

***

### definedNames()

```ts
definedNames(): Promise<WorkbookDefinedName[]>;
```

Defined in: packages/core/src/workbook.ts:1890

#### Returns

`Promise`&lt;[`WorkbookDefinedName`](../type-aliases/WorkbookDefinedName.md)[]&gt;

***

### deleteAutoFilter()

```ts
deleteAutoFilter(sheetName): Promise<boolean>;
```

Defined in: packages/core/src/workbook.ts:1232

#### Parameters

##### sheetName

`string`

#### Returns

`Promise`&lt;`boolean`&gt;

***

### deleteConditionalFormat()

```ts
deleteConditionalFormat(sheetName, sqref): Promise<boolean>;
```

Defined in: packages/core/src/workbook.ts:1243

#### Parameters

##### sheetName

`string`

##### sqref

`string`

#### Returns

`Promise`&lt;`boolean`&gt;

***

### deleteDataValidation()

```ts
deleteDataValidation(sheetName, sqref): Promise<boolean>;
```

Defined in: packages/core/src/workbook.ts:1221

#### Parameters

##### sheetName

`string`

##### sqref

`string`

#### Returns

`Promise`&lt;`boolean`&gt;

***

### deleteDefinedName()

```ts
deleteDefinedName(name, options?): Promise<boolean>;
```

Defined in: packages/core/src/workbook.ts:1919

#### Parameters

##### name

`string`

##### options?

###### sheetName?

`string`

#### Returns

`Promise`&lt;`boolean`&gt;

***

### deleteHyperlink()

```ts
deleteHyperlink(sheetName, ref): Promise<boolean>;
```

Defined in: packages/core/src/workbook.ts:1320

#### Parameters

##### sheetName

`string`

##### ref

`string`

#### Returns

`Promise`&lt;`boolean`&gt;

***

### deleteRows()

```ts
deleteRows(
   sheetName, 
   startRow, 
   count?): Promise<void>;
```

Defined in: packages/core/src/workbook.ts:584

#### Parameters

##### sheetName

`string`

##### startRow

`number`

##### count?

`number` = `1`

#### Returns

`Promise`&lt;`void`&gt;

***

### deleteSheet()

```ts
deleteSheet(sheetName): Promise<void>;
```

Defined in: packages/core/src/workbook.ts:1562

#### Parameters

##### sheetName

`string`

#### Returns

`Promise`&lt;`void`&gt;

***

### diagnostics()

```ts
diagnostics(): Diagnostic[];
```

Defined in: packages/core/src/workbook.ts:2216

#### Returns

[`Diagnostic`](../type-aliases/Diagnostic.md)[]

***

### ensureCellStyle()

```ts
ensureCellStyle(style): Promise<string>;
```

Defined in: packages/core/src/workbook.ts:1335

#### Parameters

##### style

[`WorkbookCellStyleInput`](../type-aliases/WorkbookCellStyleInput.md)

#### Returns

`Promise`&lt;`string`&gt;

***

### formulas()

```ts
formulas(): Promise<WorkbookFormula[]>;
```

Defined in: packages/core/src/workbook.ts:2169

#### Returns

`Promise`&lt;[`WorkbookFormula`](../type-aliases/WorkbookFormula.md)[]&gt;

***

### hideSheet()

```ts
hideSheet(sheetName, state?): Promise<WorkbookSheet>;
```

Defined in: packages/core/src/workbook.ts:1831

#### Parameters

##### sheetName

`string`

##### state?

[`WorkbookSheetState`](../type-aliases/WorkbookSheetState.md) = `"hidden"`

#### Returns

`Promise`&lt;[`WorkbookSheet`](../type-aliases/WorkbookSheet.md)&gt;

***

### hyperlinks()

```ts
hyperlinks(sheetName?): Promise<WorkbookHyperlink[]>;
```

Defined in: packages/core/src/workbook.ts:862

#### Parameters

##### sheetName?

`string`

#### Returns

`Promise`&lt;[`WorkbookHyperlink`](../type-aliases/WorkbookHyperlink.md)[]&gt;

***

### images()

```ts
images(sheetName?): Promise<WorkbookImage[]>;
```

Defined in: packages/core/src/workbook.ts:1004

#### Parameters

##### sheetName?

`string`

#### Returns

`Promise`&lt;[`WorkbookImage`](../type-aliases/WorkbookImage.md)[]&gt;

***

### insertImage()

```ts
insertImage(
   sheetName, 
   data, 
   options?): Promise<WorkbookImage>;
```

Defined in: packages/core/src/workbook.ts:1099

#### Parameters

##### sheetName

`string`

##### data

`Uint8Array`

##### options?

[`WorkbookInsertImageOptions`](../type-aliases/WorkbookInsertImageOptions.md) = `{}`

#### Returns

`Promise`&lt;[`WorkbookImage`](../type-aliases/WorkbookImage.md)&gt;

***

### insertRows()

```ts
insertRows(
   sheetName, 
   beforeRow, 
   count?): Promise<void>;
```

Defined in: packages/core/src/workbook.ts:575

#### Parameters

##### sheetName

`string`

##### beforeRow

`number`

##### count?

`number` = `1`

#### Returns

`Promise`&lt;`void`&gt;

***

### inspect()

```ts
inspect(): Promise<WorkbookInspectResult>;
```

Defined in: packages/core/src/workbook.ts:1859

#### Returns

`Promise`&lt;[`WorkbookInspectResult`](../type-aliases/WorkbookInspectResult.md)&gt;

***

### mergeCells()

```ts
mergeCells(sheetName, ref): Promise<WorkbookMergedCell>;
```

Defined in: packages/core/src/workbook.ts:1254

#### Parameters

##### sheetName

`string`

##### ref

`string`

#### Returns

`Promise`&lt;[`WorkbookMergedCell`](../type-aliases/WorkbookMergedCell.md)&gt;

***

### mergedCells()

```ts
mergedCells(sheetName?): Promise<WorkbookMergedCell[]>;
```

Defined in: packages/core/src/workbook.ts:901

#### Parameters

##### sheetName?

`string`

#### Returns

`Promise`&lt;[`WorkbookMergedCell`](../type-aliases/WorkbookMergedCell.md)[]&gt;

***

### namedRanges()

```ts
namedRanges(name?): Promise<WorkbookNamedRange[]>;
```

Defined in: packages/core/src/workbook.ts:803

#### Parameters

##### name?

`string`

#### Returns

`Promise`&lt;[`WorkbookNamedRange`](../type-aliases/WorkbookNamedRange.md)[]&gt;

***

### patchCell()

```ts
patchCell(
   sheetName, 
   address, 
   value): Promise<void>;
```

Defined in: packages/core/src/workbook.ts:407

#### Parameters

##### sheetName

`string`

##### address

`string`

##### value

[`CellInput`](../type-aliases/CellInput.md)

#### Returns

`Promise`&lt;`void`&gt;

***

### patchCells()

```ts
patchCells(sheetName, patches): Promise<void>;
```

Defined in: packages/core/src/workbook.ts:429

#### Parameters

##### sheetName

`string`

##### patches

[`CellPatch`](../type-aliases/CellPatch.md)[]

#### Returns

`Promise`&lt;`void`&gt;

***

### patchNamedRange()

```ts
patchNamedRange(
   name, 
   values, 
   options?): Promise<WorkbookNamedRange>;
```

Defined in: packages/core/src/workbook.ts:841

#### Parameters

##### name

`string`

##### values

[`CellInput`](../type-aliases/CellInput.md)[][]

##### options?

[`WorkbookNamedRangePatchOptions`](../type-aliases/WorkbookNamedRangePatchOptions.md) = `{}`

#### Returns

`Promise`&lt;[`WorkbookNamedRange`](../type-aliases/WorkbookNamedRange.md)&gt;

***

### patchRange()

```ts
patchRange(
   sheetName, 
   startAddress, 
   values): Promise<void>;
```

Defined in: packages/core/src/workbook.ts:451

#### Parameters

##### sheetName

`string`

##### startAddress

`string`

##### values

[`CellInput`](../type-aliases/CellInput.md)[][]

#### Returns

`Promise`&lt;`void`&gt;

***

### pivotCacheSources()

```ts
pivotCacheSources(): Promise<WorkbookPivotCacheSource[]>;
```

Defined in: packages/core/src/workbook.ts:1851

#### Returns

`Promise`&lt;[`WorkbookPivotCacheSource`](../type-aliases/WorkbookPivotCacheSource.md)[]&gt;

***

### preflightTemplatePatch()

```ts
preflightTemplatePatch(patch): Promise<WorkbookTemplatePreflightResult>;
```

Defined in: packages/core/src/workbook.ts:2061

#### Parameters

##### patch

[`WorkbookTemplatePatch`](../type-aliases/WorkbookTemplatePatch.md)

#### Returns

`Promise`&lt;[`WorkbookTemplatePreflightResult`](../type-aliases/WorkbookTemplatePreflightResult.md)&gt;

***

### readCell()

```ts
readCell(sheetName, address): Promise<ReadCellResult | undefined>;
```

Defined in: packages/core/src/workbook.ts:791

#### Parameters

##### sheetName

`string`

##### address

`string`

#### Returns

`Promise`&lt;[`ReadCellResult`](../type-aliases/ReadCellResult.md) \| `undefined`&gt;

***

### readNamedRange()

```ts
readNamedRange(name, options?): Promise<ReadRangeResult>;
```

Defined in: packages/core/src/workbook.ts:833

#### Parameters

##### name

`string`

##### options?

###### sheetName?

`string`

#### Returns

`Promise`&lt;[`ReadRangeResult`](../type-aliases/ReadRangeResult.md)&gt;

***

### readRange()

```ts
readRange(sheetName, rangeRef): Promise<ReadRangeResult>;
```

Defined in: packages/core/src/workbook.ts:797

#### Parameters

##### sheetName

`string`

##### rangeRef

`string`

#### Returns

`Promise`&lt;[`ReadRangeResult`](../type-aliases/ReadRangeResult.md)&gt;

***

### readSheetCells()

```ts
readSheetCells(sheetName): Promise<ReadCellResult[]>;
```

Defined in: packages/core/src/workbook.ts:473

#### Parameters

##### sheetName

`string`

#### Returns

`Promise`&lt;[`ReadCellResult`](../type-aliases/ReadCellResult.md)[]&gt;

***

### removeRightmostTableColumn()

```ts
removeRightmostTableColumn(tableName, columnName): Promise<WorkbookTable>;
```

Defined in: packages/core/src/workbook.ts:1491

#### Parameters

##### tableName

`string`

##### columnName

`string`

#### Returns

`Promise`&lt;[`WorkbookTable`](../type-aliases/WorkbookTable.md)&gt;

***

### renameSheet()

```ts
renameSheet(sheetName, nextName): Promise<WorkbookSheet>;
```

Defined in: packages/core/src/workbook.ts:1746

#### Parameters

##### sheetName

`string`

##### nextName

`string`

#### Returns

`Promise`&lt;[`WorkbookSheet`](../type-aliases/WorkbookSheet.md)&gt;

***

### renameTable()

```ts
renameTable(tableName, nextName): Promise<WorkbookTable>;
```

Defined in: packages/core/src/workbook.ts:1450

#### Parameters

##### tableName

`string`

##### nextName

`string`

#### Returns

`Promise`&lt;[`WorkbookTable`](../type-aliases/WorkbookTable.md)&gt;

***

### renameTableColumn()

```ts
renameTableColumn(
   tableName, 
   columnName, 
   nextName): Promise<WorkbookTable>;
```

Defined in: packages/core/src/workbook.ts:1461

#### Parameters

##### tableName

`string`

##### columnName

`string`

##### nextName

`string`

#### Returns

`Promise`&lt;[`WorkbookTable`](../type-aliases/WorkbookTable.md)&gt;

***

### renderTemplate()

```ts
renderTemplate(patch): Promise<WorkbookTemplateRenderResult>;
```

Defined in: packages/core/src/workbook.ts:740

#### Parameters

##### patch

[`WorkbookTemplatePatch`](../type-aliases/WorkbookTemplatePatch.md)

#### Returns

`Promise`&lt;[`WorkbookTemplateRenderResult`](../type-aliases/WorkbookTemplateRenderResult.md)&gt;

***

### replaceImage()

```ts
replaceImage(imagePartName, data): Promise<WorkbookImage>;
```

Defined in: packages/core/src/workbook.ts:1075

#### Parameters

##### imagePartName

`string`

##### data

`Uint8Array`

#### Returns

`Promise`&lt;[`WorkbookImage`](../type-aliases/WorkbookImage.md)&gt;

***

### replaceTableRows()

```ts
replaceTableRows(tableName, rows): Promise<WorkbookTable>;
```

Defined in: packages/core/src/workbook.ts:1438

#### Parameters

##### tableName

`string`

##### rows

[`CellInput`](../type-aliases/CellInput.md)[][]

#### Returns

`Promise`&lt;[`WorkbookTable`](../type-aliases/WorkbookTable.md)&gt;

***

### resolveNamedRange()

```ts
resolveNamedRange(name, options?): Promise<WorkbookNamedRange>;
```

Defined in: packages/core/src/workbook.ts:820

#### Parameters

##### name

`string`

##### options?

###### sheetName?

`string`

#### Returns

`Promise`&lt;[`WorkbookNamedRange`](../type-aliases/WorkbookNamedRange.md)&gt;

***

### retargetChartFormulas()

```ts
retargetChartFormulas(retargets): Promise<number>;
```

Defined in: packages/core/src/workbook.ts:1839

#### Parameters

##### retargets

[`ChartFormulaRetarget`](../type-aliases/ChartFormulaRetarget.md)[]

#### Returns

`Promise`&lt;`number`&gt;

***

### retargetPivotCacheSources()

```ts
retargetPivotCacheSources(retargets): Promise<number>;
```

Defined in: packages/core/src/workbook.ts:1843

#### Parameters

##### retargets

[`PivotCacheSourceRetarget`](../type-aliases/PivotCacheSourceRetarget.md)[]

#### Returns

`Promise`&lt;`number`&gt;

***

### setAutoFilter()

```ts
setAutoFilter(sheetName, autoFilter): Promise<WorkbookAutoFilter>;
```

Defined in: packages/core/src/workbook.ts:1188

#### Parameters

##### sheetName

`string`

##### autoFilter

[`WorksheetAutoFilter`](../type-aliases/WorksheetAutoFilter.md)

#### Returns

`Promise`&lt;[`WorkbookAutoFilter`](../type-aliases/WorkbookAutoFilter.md)&gt;

***

### setConditionalFormat()

```ts
setConditionalFormat(sheetName, conditionalFormat): Promise<WorkbookConditionalFormat>;
```

Defined in: packages/core/src/workbook.ts:1203

#### Parameters

##### sheetName

`string`

##### conditionalFormat

[`WorksheetConditionalFormat`](../type-aliases/WorksheetConditionalFormat.md)

#### Returns

`Promise`&lt;[`WorkbookConditionalFormat`](../type-aliases/WorkbookConditionalFormat.md)&gt;

***

### setDataValidation()

```ts
setDataValidation(sheetName, dataValidation): Promise<WorkbookDataValidation>;
```

Defined in: packages/core/src/workbook.ts:1170

#### Parameters

##### sheetName

`string`

##### dataValidation

[`WorksheetDataValidation`](../type-aliases/WorksheetDataValidation.md)

#### Returns

`Promise`&lt;[`WorkbookDataValidation`](../type-aliases/WorkbookDataValidation.md)&gt;

***

### setDefinedName()

```ts
setDefinedName(
   name, 
   text, 
   options?): Promise<WorkbookDefinedName>;
```

Defined in: packages/core/src/workbook.ts:1894

#### Parameters

##### name

`string`

##### text

`string`

##### options?

###### comment?

`string`

###### hidden?

`boolean`

###### sheetName?

`string`

#### Returns

`Promise`&lt;[`WorkbookDefinedName`](../type-aliases/WorkbookDefinedName.md)&gt;

***

### setHyperlink()

```ts
setHyperlink(
   sheetName, 
   ref, 
   target, 
   options?): Promise<WorkbookHyperlink>;
```

Defined in: packages/core/src/workbook.ts:1279

#### Parameters

##### sheetName

`string`

##### ref

`string`

##### target

`string`

##### options?

###### display?

`string`

###### tooltip?

`string`

#### Returns

`Promise`&lt;[`WorkbookHyperlink`](../type-aliases/WorkbookHyperlink.md)&gt;

***

### setSheetState()

```ts
setSheetState(sheetName, state): Promise<WorkbookSheet>;
```

Defined in: packages/core/src/workbook.ts:1787

#### Parameters

##### sheetName

`string`

##### state

  \| [`WorkbookSheetState`](../type-aliases/WorkbookSheetState.md)
  \| `undefined`

#### Returns

`Promise`&lt;[`WorkbookSheet`](../type-aliases/WorkbookSheet.md)&gt;

***

### sheet()

```ts
sheet(name): WorkbookSheet;
```

Defined in: packages/core/src/workbook.ts:398

#### Parameters

##### name

`string`

#### Returns

[`WorkbookSheet`](../type-aliases/WorkbookSheet.md)

***

### sheets()

```ts
sheets(): WorkbookSheet[];
```

Defined in: packages/core/src/workbook.ts:394

#### Returns

[`WorkbookSheet`](../type-aliases/WorkbookSheet.md)[]

***

### showSheet()

```ts
showSheet(sheetName): Promise<WorkbookSheet>;
```

Defined in: packages/core/src/workbook.ts:1835

#### Parameters

##### sheetName

`string`

#### Returns

`Promise`&lt;[`WorkbookSheet`](../type-aliases/WorkbookSheet.md)&gt;

***

### styleCell()

```ts
styleCell(
   sheetName, 
   address, 
   style): Promise<string>;
```

Defined in: packages/core/src/workbook.ts:1354

#### Parameters

##### sheetName

`string`

##### address

`string`

##### style

[`WorkbookCellStyleInput`](../type-aliases/WorkbookCellStyleInput.md)

#### Returns

`Promise`&lt;`string`&gt;

***

### styleRange()

```ts
styleRange(
   sheetName, 
   rangeRef, 
   style): Promise<string[]>;
```

Defined in: packages/core/src/workbook.ts:1372

#### Parameters

##### sheetName

`string`

##### rangeRef

`string`

##### style

[`WorkbookCellStyleInput`](../type-aliases/WorkbookCellStyleInput.md)

#### Returns

`Promise`&lt;`string`[]&gt;

***

### styles()

```ts
styles(): Promise<WorkbookStyles>;
```

Defined in: packages/core/src/workbook.ts:1935

#### Returns

`Promise`&lt;[`WorkbookStyles`](../type-aliases/WorkbookStyles.md)&gt;

***

### tables()

```ts
tables(): Promise<WorkbookTable[]>;
```

Defined in: packages/core/src/workbook.ts:1855

#### Returns

`Promise`&lt;[`WorkbookTable`](../type-aliases/WorkbookTable.md)[]&gt;

***

### templateManifest()

```ts
templateManifest(): Promise<WorkbookTemplateManifest>;
```

Defined in: packages/core/src/workbook.ts:1871

#### Returns

`Promise`&lt;[`WorkbookTemplateManifest`](../type-aliases/WorkbookTemplateManifest.md)&gt;

***

### unmergeCells()

```ts
unmergeCells(sheetName, ref): Promise<boolean>;
```

Defined in: packages/core/src/workbook.ts:1268

#### Parameters

##### sheetName

`string`

##### ref

`string`

#### Returns

`Promise`&lt;`boolean`&gt;

***

### validate()

```ts
validate(): Promise<ValidationReport>;
```

Defined in: packages/core/src/workbook.ts:1886

#### Returns

`Promise`&lt;[`ValidationReport`](../type-aliases/ValidationReport.md)&gt;

***

### write()

```ts
write(): Promise<Uint8Array<ArrayBufferLike>>;
```

Defined in: packages/core/src/workbook.ts:2220

#### Returns

`Promise`&lt;`Uint8Array`&lt;`ArrayBufferLike`&gt;&gt;

***

### fromPackage()

```ts
static fromPackage(pkg): Promise<Workbook>;
```

Defined in: packages/core/src/workbook.ts:385

#### Parameters

##### pkg

[`OoxmlPackage`](OoxmlPackage.md)

#### Returns

`Promise`&lt;`Workbook`&gt;
