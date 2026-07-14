---
title: writeWorkbookSafely
---

# Function: writeWorkbookSafely()

```ts
function writeWorkbookSafely(
   workbook, 
   outputPath, 
   beforeData, 
   options?): Promise<WorkbookSafeWriteReport>;
```

Defined in: node/src/index.ts:62

## Parameters

### workbook

[`Workbook`](../../core/classes/Workbook.md)

### outputPath

`string`

### beforeData

`Uint8Array`

### options?

[`WorkbookSafeWriteOptions`](../type-aliases/WorkbookSafeWriteOptions.md) = `{}`

## Returns

`Promise`&lt;[`WorkbookSafeWriteReport`](../type-aliases/WorkbookSafeWriteReport.md)&gt;
