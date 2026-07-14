---
title: WorkbookSemanticDiff
---

# Type Alias: WorkbookSemanticDiff

```ts
type WorkbookSemanticDiff = object;
```

Defined in: packages/core/src/workbook-diff.ts:27

## Properties

### cells

```ts
cells: WorkbookCellDiff[];
```

Defined in: packages/core/src/workbook-diff.ts:28

***

### definedNames

```ts
definedNames: WorkbookNameListDiff;
```

Defined in: packages/core/src/workbook-diff.ts:29

***

### sheets

```ts
sheets: object;
```

Defined in: packages/core/src/workbook-diff.ts:30

#### added

```ts
added: string[];
```

#### removed

```ts
removed: string[];
```

***

### summary

```ts
summary: object;
```

Defined in: packages/core/src/workbook-diff.ts:31

#### addedCells

```ts
addedCells: number;
```

#### changedCells

```ts
changedCells: number;
```

#### removedCells

```ts
removedCells: number;
```

#### truncated

```ts
truncated: boolean;
```

***

### tables

```ts
tables: WorkbookNameListDiff;
```

Defined in: packages/core/src/workbook-diff.ts:37
