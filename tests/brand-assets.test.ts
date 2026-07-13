import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import process from "node:process";
import test from "node:test";

const brandDirectory = join(process.cwd(), "docs", "assets", "brand");
const faviconDirectory = join(brandDirectory, "favicons");

async function readPngDimensions(path: string): Promise<{ width: number; height: number }> {
  const png = await readFile(path);
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  assert.deepEqual(png.subarray(0, signature.length), signature);

  return {
    width: png.readUInt32BE(16),
    height: png.readUInt32BE(20)
  };
}

test("brand images use their intended production dimensions", async () => {
  assert.deepEqual(await readPngDimensions(join(brandDirectory, "ironsheet-icon.png")), {
    width: 1024,
    height: 1024
  });
  assert.deepEqual(await readPngDimensions(join(brandDirectory, "ironsheet-opengraph.png")), {
    width: 1200,
    height: 630
  });
  assert.deepEqual(await readPngDimensions(join(faviconDirectory, "favicon-16x16.png")), {
    width: 16,
    height: 16
  });
  assert.deepEqual(await readPngDimensions(join(faviconDirectory, "apple-touch-icon.png")), {
    width: 180,
    height: 180
  });
});

test("favicon.ico includes the requested multi-size icon directory", async () => {
  const ico = await readFile(join(faviconDirectory, "favicon.ico"));

  assert.equal(ico.readUInt16LE(0), 0);
  assert.equal(ico.readUInt16LE(2), 1);
  assert.equal(ico.readUInt16LE(4), 4);
});

test("favicon manifest carries the Ironsheet brand metadata", async () => {
  const manifest = JSON.parse(
    await readFile(join(faviconDirectory, "site.webmanifest"), "utf8")
  ) as {
    name: string;
    short_name: string;
    theme_color: string;
    background_color: string;
    icons: Array<{ src: string; sizes: string; type: string }>;
  };

  assert.equal(manifest.name, "Ironsheet");
  assert.equal(manifest.short_name, "Ironsheet");
  assert.equal(manifest.theme_color, "#17324d");
  assert.equal(manifest.background_color, "#f7fafc");
  assert.deepEqual(
    manifest.icons.map((icon) => icon.sizes),
    ["192x192", "512x512"]
  );
});
