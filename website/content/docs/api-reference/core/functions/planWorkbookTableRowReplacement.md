---
title: planWorkbookTableRowReplacement
---

# Function: planWorkbookTableRowReplacement()

```ts
function planWorkbookTableRowReplacement(
   pkg, 
   tableName, 
   rows): Promise<WorkbookTableRowReplacementPlan>;
```

Defined in: packages/core/src/table.ts:118

## Parameters

### pkg

[`OoxmlPackage`](../classes/OoxmlPackage.md)

### tableName

`string`

### rows

[`CellInput`](../type-aliases/CellInput.md)[][]

## Returns

`Promise`&lt;[`WorkbookTableRowReplacementPlan`](../type-aliases/WorkbookTableRowReplacementPlan.md)&gt;
