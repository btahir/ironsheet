# Ironsheet Strategy Notes

Date: 2026-07-11. Distilled from ecosystem research across GitHub issues, security advisories, pricing pages, and funding news. This document guides prioritization; the technical spec stays in `IRONSHEET_SPEC.md`.

## 1. The window is open

The two incumbent JavaScript Excel libraries are decaying at the same time:

- **SheetJS (`xlsx`)**: the npm package has been frozen at 0.18.5 since 2022; security fixes ship only from their own CDN, so `npm audit` and enterprise scanners permanently flag it. Styling, template editing, and formula evaluation are Pro-only with quote-only pricing. ~7-8M weekly downloads of a package compliance teams are actively migrating away from.
- **ExcelJS**: no release since 4.4.0 (2023), no maintainer activity, open "intent to fork" discussions, and fragmented community forks. Its most-upvoted issues are exactly Ironsheet's wedge: charts deleted on round-trip, pivot tables destroyed on save, files Excel flags as corrupt.
- **xlsx-populate** proved the "patch, don't re-model" philosophy has demand, then went dormant.
- **excelize** (Go, 20k+ stars) shows what the category default looks like when someone maintains a high-compatibility engine. No JavaScript equivalent exists.

The dominant complaint themes across all of them: (1) charts/pivots/validations vanish after save, (2) Excel repair prompts, (3) styling paywalled or lost, (4) OOM on large files, (5) "is this maintained?", (6) unfixable `npm audit` failures.

Competitive watch item: `@office-kit/xlsx` (MIT, pre-1.0, ~14 stars) claims the same lossless wedge. Ironsheet's answer is a faster stable 1.0 plus the layers a pure library will not fund: semantic diff, validation gates, template service, MCP server.

## 2. Positioning

> Patch real Excel files without breaking them — and prove it.

The "prove it" is the differentiator no one else markets: validation before write, package diffs, semantic cell diffs, and a public fidelity corpus. Corruption fear is the purchase trigger in this market; trust artifacts are the product.

Counter-position directly against incumbent weaknesses:

- npm-first releases with published security policy (vs SheetJS CDN distribution).
- Transparent pricing from day one (vs quote-only Pro tiers).
- Active release cadence as a first-class marketing signal.

## 3. Open-core split

Free (Apache-2.0, this repo):

- The whole lossless mutation engine: cells, ranges, tables, names, styles, rows, sheets, images.
- Validation, diagnostics, package diff, semantic workbook diff, CLI.
- Basic template rendering (named ranges, tables, images).
- Browser/Node adapters, compatibility harness.

Paid layers (proven pricing comps in parentheses):

1. **Hosted template-fill + validation API** — JSON in, verified workbook + diff report out, per-document metering with a free tier (carbone: EUR 29-295/mo; APITemplate: from $35/mo). Validation-on-every-render is the differentiator carbone lacks.
2. **Template DSL modules** — loops, conditional blocks, chart/image injection into real Excel templates (docxtemplater sells 19 modules at EUR 250-400/module/yr; PRO packs EUR 1,250/yr; OEM EUR 6,500-15,000/yr).
3. **Enterprise** — SLA, LTS branches, security attestation, private support (the SheetJS npm fiasco makes "guaranteed npm-published patches" itself sellable). Formula recalculation as a later paid module (SheetJS Pro and Univer Pro both gate it).

Weak idea, deprioritized: standalone Excel-diff CI SaaS. Existing market is fragmented desktop tools; diff works better as a free adoption hook and as a feature inside the render API.

## 4. Distribution levers (ranked)

1. **Fidelity proof page**: round-trip a public corpus of real workbooks; publish per-library results (charts destroyed in X% by ExcelJS, validations dropped by SheetJS CE). More differentiating than any speed benchmark.
2. **Migration pages**: "ExcelJS alternative", "SheetJS migration" — capture the abandonment-driven search intent that already exists.
3. **Official MCP server** exposing read/patch/diff/validate tools. Every AI spreadsheet agent (Endex raised $14M from OpenAI's fund; Shortcut raised $33M) needs a safe headless mutation engine; existing Excel MCP servers are built on fragile libraries and already market "no corruption" as the differentiator. This is both product and funnel.
4. **Browser playground**: drop an xlsx, patch it live, see the semantic diff. Zero-install proof of the promise.
5. **AI-legible docs** (llms.txt, agent-friendly API reference): library selection increasingly happens inside coding agents.

## 5. Formula evaluation stance

Do not build evaluation into core. Preserve formulas, rewrite references, mark recalculation (done). Offer a pluggable recalc adapter interface so HyperFormula (GPL/commercial) stays isolated behind an optional integration, and keep native recalculation as a possible paid module later.

## 6. Next engineering priorities after this pass

1. Column insert/delete (mirror of row edits, same reference-rewrite machinery).
2. Real Excel-authored fixtures for the 5 pending corpus slots; the fidelity claim needs Excel-touched files, not only synthetic OOXML.
3. Streaming workbook writer (the spec's remaining headline bet) for multi-hundred-MB exports.
4. Comment/note authoring (threaded comments first; VML notes are legacy).
5. Table insert-rows-inside support so row edits and tables compose instead of refusing.
6. MCP server package wrapping the node adapter.
