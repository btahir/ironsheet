---
title: streamReplaceWorksheetRowsXml
---

# Function: streamReplaceWorksheetRowsXml()

```ts
function streamReplaceWorksheetRowsXml(chunks, replacements): AsyncGenerator<string>;
```

Defined in: packages/core/src/worksheet.ts:633

## Parameters

### chunks

  \| `Iterable`&lt;`string`, `any`, `any`&gt;
  \| `AsyncIterable`&lt;`string`, `any`, `any`&gt;

### replacements

`Iterable`&lt;[`WorksheetRowReplacement`](../type-aliases/WorksheetRowReplacement.md)&gt;

## Returns

`AsyncGenerator`&lt;`string`&gt;
