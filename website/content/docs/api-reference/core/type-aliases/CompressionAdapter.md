---
title: CompressionAdapter
---

# Type Alias: CompressionAdapter

```ts
type CompressionAdapter = object;
```

Defined in: packages/core/src/zip/index.ts:29

## Methods

### deflateRaw()

```ts
deflateRaw(data): 
  | Uint8Array<ArrayBufferLike>
  | Promise<Uint8Array<ArrayBufferLike>>;
```

Defined in: packages/core/src/zip/index.ts:31

#### Parameters

##### data

`Uint8Array`

#### Returns

  \| `Uint8Array`&lt;`ArrayBufferLike`&gt;
  \| `Promise`&lt;`Uint8Array`&lt;`ArrayBufferLike`&gt;&gt;

***

### inflateRaw()

```ts
inflateRaw(data): 
  | Uint8Array<ArrayBufferLike>
  | Promise<Uint8Array<ArrayBufferLike>>;
```

Defined in: packages/core/src/zip/index.ts:30

#### Parameters

##### data

`Uint8Array`

#### Returns

  \| `Uint8Array`&lt;`ArrayBufferLike`&gt;
  \| `Promise`&lt;`Uint8Array`&lt;`ArrayBufferLike`&gt;&gt;
