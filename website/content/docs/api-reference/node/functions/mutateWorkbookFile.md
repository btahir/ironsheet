---
title: mutateWorkbookFile
---

# Function: mutateWorkbookFile()

```ts
function mutateWorkbookFile(
   inputPath, 
   outputPath, 
   mutate, 
   options?): Promise<WorkbookSafeWriteReport>;
```

Defined in: node/src/index.ts:84

## Parameters

### inputPath

`string`

### outputPath

`string`

### mutate

(`workbook`) => `void` \| `Promise`&lt;`void`&gt;

### options?

[`WorkbookSafeWriteOptions`](../type-aliases/WorkbookSafeWriteOptions.md) = `{}`

## Returns

`Promise`&lt;[`WorkbookSafeWriteReport`](../type-aliases/WorkbookSafeWriteReport.md)&gt;
