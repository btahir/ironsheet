---
title: transformXmlChunks
---

# Function: transformXmlChunks()

```ts
function transformXmlChunks(chunks, transform): AsyncGenerator<string>;
```

Defined in: packages/core/src/xml.ts:188

## Parameters

### chunks

  \| `Iterable`&lt;`string`, `any`, `any`&gt;
  \| `AsyncIterable`&lt;`string`, `any`, `any`&gt;

### transform

[`XmlChunkTransform`](../type-aliases/XmlChunkTransform.md)

## Returns

`AsyncGenerator`&lt;`string`&gt;
