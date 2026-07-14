---
title: WorkbookImageAnchor
---

# Type Alias: WorkbookImageAnchor

```ts
type WorkbookImageAnchor = 
  | {
  ext: WorkbookImageExtent;
  from: WorkbookImageAnchorMarker;
  kind: "oneCell";
}
  | {
  editAs?: "twoCell" | "oneCell" | "absolute";
  from: WorkbookImageAnchorMarker;
  kind: "twoCell";
  to: WorkbookImageAnchorMarker;
};
```

Defined in: packages/core/src/images.ts:42
