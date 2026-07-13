# Ironsheet brand assets

The master icon and OpenGraph artwork were generated with OpenAI image generation, then normalized and compressed locally with [Favipack](https://www.npmjs.com/package/favipack).

## Production files

- `ironsheet-icon.png`: transparent 1024x1024 brand mark.
- `ironsheet-opengraph.png`: 1200x630 PNG for GitHub and OpenGraph previews.
- `ironsheet-opengraph.webp`: compressed 1200x630 WebP for websites that support it.
- `favicons/`: multi-size ICO, PNG icons, Apple touch icon, Android icons, and web manifest.

The `*-source.png` files are the image-generation masters. Rebuild all deterministic derivatives with:

```bash
npm run brand:assets
```

The exact production prompts are recorded in [`PROMPTS.md`](PROMPTS.md).

For a future website, copy the contents of `favicons/` into its public root and add:

```html
<link rel="icon" href="/favicon.ico" sizes="any">
<link rel="icon" type="image/png" href="/favicon-32x32.png" sizes="32x32">
<link rel="apple-touch-icon" href="/apple-touch-icon.png">
<link rel="manifest" href="/site.webmanifest">
<meta property="og:image" content="https://example.com/ironsheet-opengraph.png">
<meta name="twitter:card" content="summary_large_image">
```
