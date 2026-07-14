---
title: XmlToken
---

# Type Alias: XmlToken

```ts
type XmlToken = 
  | {
  end: number;
  kind: "text";
  start: number;
  text: string;
}
  | {
  kind: "start";
  tag: XmlTag;
}
  | {
  end: number;
  kind: "end";
  localName: string;
  name: string;
  raw: string;
  start: number;
}
  | {
  end: number;
  kind: "comment";
  raw: string;
  start: number;
  text: string;
}
  | {
  end: number;
  kind: "cdata";
  raw: string;
  start: number;
  text: string;
}
  | {
  end: number;
  kind: "processingInstruction";
  raw: string;
  start: number;
}
  | {
  end: number;
  kind: "declaration";
  raw: string;
  start: number;
};
```

Defined in: packages/core/src/xml.ts:13
