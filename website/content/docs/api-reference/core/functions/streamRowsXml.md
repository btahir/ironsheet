---
title: streamRowsXml
---

# Function: streamRowsXml()

```ts
function streamRowsXml(rows, options): AsyncGenerator<string>;
```

Defined in: packages/core/src/worksheet.ts:605

## Parameters

### rows

  \| `Iterable`&lt;[`CellInput`](../type-aliases/CellInput.md)[], `any`, `any`&gt;
  \| `AsyncIterable`&lt;[`CellInput`](../type-aliases/CellInput.md)[], `any`, `any`&gt;

### options

#### startColumn?

`number`

#### startRow

`number`

## Returns

`AsyncGenerator`&lt;`string`&gt;
