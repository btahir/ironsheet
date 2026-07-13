# Ironsheet image-generation prompts

These production masters were created with the built-in OpenAI image-generation tool. The icon used a removable chroma-key background; the OpenGraph card used the generated icon as a visual reference.

## Icon master

```text
Use case: logo-brand
Asset type: master app icon and favicon source for Ironsheet, a preservation-first TypeScript OOXML workbook engine
Primary request: create an original emblem combining a rigid iron shield/frame with a simplified spreadsheet grid, subtly suggesting that a workbook is protected during precise edits
Style/medium: crisp flat vector-like brand mark, geometric, minimal, production-ready
Composition/framing: one centered square emblem with an exceptionally clear silhouette, generous even padding, readable at 16 pixels
Color palette: deep navy #17324D and safety teal #0F766E with a small warm amber #B7791F accent; no green close to the background key
Scene/backdrop: perfectly flat solid #FF00FF chroma-key background for removal
Constraints: no text, no letters, no numbers, no gradients, no 3D, no mockup, no watermark, no drop shadow, no cast shadow, no reflection; the background must be one uniform color with no texture or lighting variation; keep all subject edges crisp and fully separated from the background; do not use #FF00FF anywhere in the emblem
Avoid: generic Excel logo, Microsoft branding, rounded app-tile background, thin details, tiny cells, ornate metal texture
```

## OpenGraph master

```text
Use case: ads-marketing
Asset type: OpenGraph and social preview card for the Ironsheet TypeScript library, landscape 1200x630 target ratio (1.91:1)
Input images: Image 1 is the exact brand emblem reference; preserve its recognizable shield, spreadsheet grid, navy, teal, white, and amber language
Primary request: create a polished editorial-style social card for a preservation-first OOXML workbook engine
Scene/backdrop: deep navy technical backdrop with a subtle oversized spreadsheet grid and restrained package-diff lines; clean, premium developer-tool aesthetic
Subject: place a simplified faithful rendering of the Image 1 emblem as the main visual anchor, balanced with product copy
Composition/framing: very wide landscape card, safe margins of at least 64px, large emblem on one side and a clean text block on the other, strong hierarchy, nothing near crop edges
Lighting/mood: confident, precise, trustworthy, quietly industrial
Color palette: deep navy #17324D, safety teal #0F766E, off-white #F7FAFC, warm amber #B7791F
Text (verbatim): "Ironsheet" and "Move fast and break no spreadsheets."
Typography: bold modern sans-serif product name; smaller high-contrast tagline; render each line exactly once
Constraints: exact spelling I-r-o-n-s-h-e-e-t; no extra copy, no tiny labels, no logos besides the supplied original emblem, no Microsoft or Excel logo, no watermark, no photographic mockup
Avoid: busy code screenshots, glossy 3D, gradients that muddy contrast, generic stock imagery
```
