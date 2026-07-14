---
title: appendWorkbookTableColumn
---

# Function: appendWorkbookTableColumn()

```ts
function appendWorkbookTableColumn(
   pkg, 
   tableName, 
   columnName, 
   values?): Promise<WorkbookTable>;
```

Defined in: packages/core/src/table.ts:235

## Parameters

### pkg

[`OoxmlPackage`](../classes/OoxmlPackage.md)

### tableName

`string`

### columnName

`string`

### values?

[`CellInput`](../type-aliases/CellInput.md)[] = `[]`

## Returns

`Promise`&lt;[`WorkbookTable`](../type-aliases/WorkbookTable.md)&gt;
