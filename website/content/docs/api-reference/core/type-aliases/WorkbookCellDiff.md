---
title: WorkbookCellDiff
---

# Type Alias: WorkbookCellDiff

```ts
type WorkbookCellDiff = object;
```

Defined in: packages/core/src/workbook-diff.ts:12

## Properties

### address

```ts
address: string;
```

Defined in: packages/core/src/workbook-diff.ts:13

***

### after?

```ts
optional after?: WorkbookCellDiffSide;
```

Defined in: packages/core/src/workbook-diff.ts:14

***

### before?

```ts
optional before?: WorkbookCellDiffSide;
```

Defined in: packages/core/src/workbook-diff.ts:15

***

### changed?

```ts
optional changed?: ("formula" | "style" | "value")[];
```

Defined in: packages/core/src/workbook-diff.ts:16

***

### kind

```ts
kind: WorkbookCellDiffKind;
```

Defined in: packages/core/src/workbook-diff.ts:17

***

### sheetName

```ts
sheetName: string;
```

Defined in: packages/core/src/workbook-diff.ts:18
