---
title: OoxmlPackage
---

# Class: OoxmlPackage

Defined in: packages/core/src/opc.ts:33

## Properties

### parts

```ts
readonly parts: Map<string, PackagePart>;
```

Defined in: packages/core/src/opc.ts:34

## Methods

### addPart()

```ts
addPart(partName, data): void;
```

Defined in: packages/core/src/opc.ts:117

#### Parameters

##### partName

`string`

##### data

`Uint8Array`

#### Returns

`void`

***

### addTextPart()

```ts
addTextPart(partName, text): void;
```

Defined in: packages/core/src/opc.ts:105

#### Parameters

##### partName

`string`

##### text

`string`

#### Returns

`void`

***

### deletePart()

```ts
deletePart(partName): boolean;
```

Defined in: packages/core/src/opc.ts:129

#### Parameters

##### partName

`string`

#### Returns

`boolean`

***

### hasPart()

```ts
hasPart(partName): boolean;
```

Defined in: packages/core/src/opc.ts:56

#### Parameters

##### partName

`string`

#### Returns

`boolean`

***

### inspect()

```ts
inspect(): Promise<PackageInspectResult>;
```

Defined in: packages/core/src/opc.ts:353

#### Returns

`Promise`&lt;[`PackageInspectResult`](../type-aliases/PackageInspectResult.md)&gt;

***

### listParts()

```ts
listParts(): string[];
```

Defined in: packages/core/src/opc.ts:52

#### Returns

`string`[]

***

### nextRelationshipId()

```ts
nextRelationshipId(partName, prefix?): Promise<string>;
```

Defined in: packages/core/src/opc.ts:177

#### Parameters

##### partName

`string`

##### prefix?

`string` = `"rId"`

#### Returns

`Promise`&lt;`string`&gt;

***

### readPart()

```ts
readPart(partName): Promise<Uint8Array<ArrayBufferLike>>;
```

Defined in: packages/core/src/opc.ts:60

#### Parameters

##### partName

`string`

#### Returns

`Promise`&lt;`Uint8Array`&lt;`ArrayBufferLike`&gt;&gt;

***

### readText()

```ts
readText(partName): Promise<string>;
```

Defined in: packages/core/src/opc.ts:77

#### Parameters

##### partName

`string`

#### Returns

`Promise`&lt;`string`&gt;

***

### relationshipsFor()

```ts
relationshipsFor(partName): Promise<Relationship[]>;
```

Defined in: packages/core/src/opc.ts:133

#### Parameters

##### partName

`string`

#### Returns

`Promise`&lt;[`Relationship`](../type-aliases/Relationship.md)[]&gt;

***

### removeContentTypeOverride()

```ts
removeContentTypeOverride(partName): Promise<boolean>;
```

Defined in: packages/core/src/opc.ts:250

#### Parameters

##### partName

`string`

#### Returns

`Promise`&lt;`boolean`&gt;

***

### removeRelationships()

```ts
removeRelationships(partName, predicate): Promise<number>;
```

Defined in: packages/core/src/opc.ts:144

#### Parameters

##### partName

`string`

##### predicate

(`relationship`) => `boolean`

#### Returns

`Promise`&lt;`number`&gt;

***

### rootRelationships()

```ts
rootRelationships(): Promise<Relationship[]>;
```

Defined in: packages/core/src/opc.ts:345

#### Returns

`Promise`&lt;[`Relationship`](../type-aliases/Relationship.md)[]&gt;

***

### setPart()

```ts
setPart(partName, data): void;
```

Defined in: packages/core/src/opc.ts:93

#### Parameters

##### partName

`string`

##### data

`Uint8Array`

#### Returns

`void`

***

### setText()

```ts
setText(partName, text): void;
```

Defined in: packages/core/src/opc.ts:81

#### Parameters

##### partName

`string`

##### text

`string`

#### Returns

`void`

***

### upsertContentTypeDefault()

```ts
upsertContentTypeDefault(extension, contentType): Promise<void>;
```

Defined in: packages/core/src/opc.ts:270

#### Parameters

##### extension

`string`

##### contentType

`string`

#### Returns

`Promise`&lt;`void`&gt;

***

### upsertContentTypeOverride()

```ts
upsertContentTypeOverride(partName, contentType): Promise<void>;
```

Defined in: packages/core/src/opc.ts:308

#### Parameters

##### partName

`string`

##### contentType

`string`

#### Returns

`Promise`&lt;`void`&gt;

***

### upsertRelationship()

```ts
upsertRelationship(partName, relationship): Promise<void>;
```

Defined in: packages/core/src/opc.ts:198

#### Parameters

##### partName

`string`

##### relationship

[`Relationship`](../type-aliases/Relationship.md)

#### Returns

`Promise`&lt;`void`&gt;

***

### write()

```ts
write(): Promise<Uint8Array<ArrayBufferLike>>;
```

Defined in: packages/core/src/opc.ts:377

#### Returns

`Promise`&lt;`Uint8Array`&lt;`ArrayBufferLike`&gt;&gt;

***

### open()

```ts
static open(data, compression): OoxmlPackage;
```

Defined in: packages/core/src/opc.ts:48

#### Parameters

##### data

`Uint8Array`

##### compression

[`CompressionAdapter`](../type-aliases/CompressionAdapter.md)

#### Returns

`OoxmlPackage`
