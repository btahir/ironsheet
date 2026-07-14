---
title: streamXmlElements
---

# Function: streamXmlElements()

```ts
function streamXmlElements(chunks, localName): AsyncGenerator<XmlElementChunk>;
```

Defined in: packages/core/src/xml.ts:197

## Parameters

### chunks

  \| `Iterable`&lt;`string`, `any`, `any`&gt;
  \| `AsyncIterable`&lt;`string`, `any`, `any`&gt;

### localName

`string`

## Returns

`AsyncGenerator`&lt;[`XmlElementChunk`](../type-aliases/XmlElementChunk.md)&gt;
