import assert from "node:assert/strict";
import test from "node:test";
import { openPackage } from "../packages/node/src/index.ts";
import { createMinimalWorkbook } from "./helpers/minimal-xlsx.ts";

test("OPC packages can read newly added binary parts before write", async () => {
  const pkg = await openPackage(await createMinimalWorkbook());
  const bytes = new Uint8Array([1, 2, 3, 4]);

  pkg.addPart("xl/media/new-image.png", bytes);

  assert.deepEqual(Array.from(await pkg.readPart("xl/media/new-image.png")), [...bytes]);
});
