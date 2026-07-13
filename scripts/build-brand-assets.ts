#!/usr/bin/env tsx
import { mkdir, rm, stat } from "node:fs/promises";
import { join, relative } from "node:path";
import process from "node:process";
import { compressFile, createFaviconPack } from "favipack";

const brandDirectory = join(process.cwd(), "docs", "assets", "brand");
const faviconDirectory = join(brandDirectory, "favicons");
const iconSource = join(brandDirectory, "ironsheet-icon-source.png");
const openGraphSource = join(brandDirectory, "ironsheet-opengraph-source.png");
const iconOutput = join(brandDirectory, "ironsheet-icon.png");
const openGraphPng = join(brandDirectory, "ironsheet-opengraph.png");
const openGraphWebp = join(brandDirectory, "ironsheet-opengraph.webp");

async function describeOutput(path: string): Promise<string> {
  const { size } = await stat(path);
  return `${relative(process.cwd(), path)} (${Math.ceil(size / 1024)} KiB)`;
}

await mkdir(brandDirectory, { recursive: true });

await Promise.all([
  compressFile(iconSource, iconOutput, {
    format: "png",
    resize: {
      width: 1024,
      height: 1024,
      fit: "contain",
      withoutEnlargement: true,
      background: { r: 0, g: 0, b: 0, alpha: 0 }
    },
    png: {
      compressionLevel: 9,
      adaptiveFiltering: true,
      palette: true,
      quality: 100,
      colors: 256,
      effort: 10,
      dither: 1
    }
  }),
  compressFile(openGraphSource, openGraphPng, {
    format: "png",
    resize: {
      width: 1200,
      height: 630,
      fit: "cover"
    },
    png: {
      compressionLevel: 9,
      adaptiveFiltering: true,
      palette: true,
      quality: 100,
      colors: 256,
      effort: 10,
      dither: 1
    }
  }),
  compressFile(openGraphSource, openGraphWebp, {
    format: "webp",
    resize: {
      width: 1200,
      height: 630,
      fit: "cover"
    },
    webp: {
      quality: 88,
      alphaQuality: 100,
      smartSubsample: true,
      effort: 6
    }
  })
]);

await rm(faviconDirectory, { recursive: true, force: true });
const faviconPack = await createFaviconPack(iconOutput, faviconDirectory, {
  appName: "Ironsheet",
  shortName: "Ironsheet",
  themeColor: "#17324d",
  backgroundColor: "#f7fafc",
  display: "standalone",
  fit: "contain",
  pathPrefix: "/",
  ico: {
    sizes: [16, 32, 48, 256],
    format: "auto"
  },
  png: {
    compressionLevel: 9,
    adaptiveFiltering: true,
    effort: 10,
    palette: true
  }
});

const outputs = [
  iconOutput,
  openGraphPng,
  openGraphWebp,
  ...faviconPack.files.map((file) => file.path)
];

console.log("brand assets: generated");
for (const output of outputs) {
  console.log(`- ${await describeOutput(output)}`);
}
