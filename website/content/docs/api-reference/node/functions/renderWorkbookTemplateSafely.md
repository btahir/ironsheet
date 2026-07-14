---
title: renderWorkbookTemplateSafely
---

# Function: renderWorkbookTemplateSafely()

```ts
function renderWorkbookTemplateSafely(
   inputPath, 
   outputPath, 
   patch, 
   options?): Promise<WorkbookSafeTemplateRenderReport>;
```

Defined in: node/src/index.ts:268

## Parameters

### inputPath

`string`

### outputPath

`string`

### patch

[`WorkbookTemplatePatch`](../../core/type-aliases/WorkbookTemplatePatch.md)

### options?

[`WorkbookSafeWriteOptions`](../type-aliases/WorkbookSafeWriteOptions.md) = `{}`

## Returns

`Promise`&lt;[`WorkbookSafeTemplateRenderReport`](../type-aliases/WorkbookSafeTemplateRenderReport.md)&gt;
