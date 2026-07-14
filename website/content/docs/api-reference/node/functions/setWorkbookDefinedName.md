---
title: setWorkbookDefinedName
---

# Function: setWorkbookDefinedName()

```ts
function setWorkbookDefinedName(
   inputPath, 
   outputPath, 
   name, 
   text, 
   options?): Promise<void>;
```

Defined in: node/src/index.ts:417

## Parameters

### inputPath

`string`

### outputPath

`string`

### name

`string`

### text

`string`

### options?

  \| \{
  `comment?`: `string`;
  `hidden?`: `boolean`;
  `sheetName?`: `string`;
\}
  \| `undefined`

## Returns

`Promise`&lt;`void`&gt;
