---
title: WorkbookCellStyleInput
---

# Type Alias: WorkbookCellStyleInput

```ts
type WorkbookCellStyleInput = Partial<Omit<WorkbookCellFormat, "alignment">> & object;
```

Defined in: packages/core/src/styles.ts:115

## Type Declaration

### alignment?

```ts
optional alignment?: 
  | WorkbookAlignmentInput
  | WorkbookCellAlignment;
```

### border?

```ts
optional border?: WorkbookBorderInput;
```

### fill?

```ts
optional fill?: WorkbookFillInput;
```

### font?

```ts
optional font?: WorkbookFontInput;
```

### numberFormat?

```ts
optional numberFormat?: string;
```
