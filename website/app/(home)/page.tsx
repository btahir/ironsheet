import {
  DiffIcon,
  FileCheck2,
  GitFork,
  Globe,
  Layers,
  ShieldCheck,
  Terminal,
} from 'lucide-react';
import Link from 'next/link';
import { appDescription, appTagline, gitConfig } from '@/lib/shared';

const githubUrl = `https://github.com/${gitConfig.user}/${gitConfig.repo}`;

const installCommand = 'npm install @ironsheet/node';

const codeSnippet = `import { mutateWorkbookFile } from "@ironsheet/node";

const report = await mutateWorkbookFile("template.xlsx", "report.xlsx", async (workbook) => {
  await workbook.patchCell("Summary", "B2", "Q1");
  await workbook.patchNamedRange("RevenueRange", [["North", 42000]]);
  await workbook.replaceTableRows("RevenueTable", [["North", 42000]]);
});

if (!report.wrote) {
  throw new Error("Ironsheet refused to write an invalid workbook");
}`;

const features = [
  {
    icon: ShieldCheck,
    title: 'Preservation-first',
    body: 'Untouched ZIP entries and unknown XML — charts, styles, pivots, drawings — are preserved by default. Ironsheet patches only what you target.',
  },
  {
    icon: Layers,
    title: 'Transactional template fills',
    body: 'Template rendering resolves anchors, validates resize plans, then applies one transaction. Fills update named ranges and tables without shifting layout.',
  },
  {
    icon: FileCheck2,
    title: 'Validation-gated writes',
    body: 'Safe writes validate the OOXML package before committing. Invalid output is refused and nothing touches your workbook.',
  },
  {
    icon: DiffIcon,
    title: 'Package + cell diffs',
    body: 'Every safe mutation returns diagnostics, validation results, and content-vs-container package diffs — proof of exactly what changed.',
  },
  {
    icon: Terminal,
    title: 'Macro-safe XLSM',
    body: 'Macro parts are preserved byte-for-byte. Generate and edit macro-enabled workbooks without breaking the VBA project.',
  },
  {
    icon: Globe,
    title: 'Runtime-neutral core',
    body: 'A dependency-free, browser-compatible engine. Runtime-specific IO and compression live in thin Node and browser adapters.',
  },
];

export default function HomePage() {
  return (
    <main className="flex flex-1 flex-col">
      {/* Hero */}
      <section className="mx-auto flex w-full max-w-5xl flex-col items-center px-6 pt-20 pb-12 text-center sm:pt-28">
        <span className="mb-6 inline-flex items-center rounded-full border border-fd-primary/30 bg-fd-primary/10 px-3 py-1 text-sm font-medium text-fd-primary">
          Lossless XLSX &amp; XLSM editing for TypeScript
        </span>
        <h1 className="text-balance text-4xl font-bold tracking-tight sm:text-6xl">
          {appTagline}
        </h1>
        <p className="mt-6 max-w-2xl text-balance text-lg text-fd-muted-foreground">
          {appDescription}
        </p>

        <div className="mt-8 flex flex-wrap items-center justify-center gap-3">
          <Link
            href="/docs"
            className="inline-flex h-11 items-center justify-center rounded-lg bg-fd-primary px-6 text-sm font-semibold text-fd-primary-foreground transition-opacity hover:opacity-90"
          >
            Get started
          </Link>
          <a
            href={githubUrl}
            target="_blank"
            rel="noreferrer"
            className="inline-flex h-11 items-center justify-center gap-2 rounded-lg border border-fd-border bg-fd-card px-6 text-sm font-semibold transition-colors hover:bg-fd-accent"
          >
            <GitFork className="size-4" />
            GitHub
          </a>
        </div>

        <div className="mt-8 w-full max-w-md">
          <div className="flex items-center gap-3 rounded-lg border border-fd-border bg-fd-card px-4 py-3 text-left font-mono text-sm">
            <span className="select-none text-fd-muted-foreground">$</span>
            <code className="flex-1 text-fd-foreground">{installCommand}</code>
          </div>
        </div>
      </section>

      {/* Code snippet */}
      <section className="mx-auto w-full max-w-3xl px-6 pb-16">
        <div className="overflow-hidden rounded-xl border border-fd-border bg-fd-card">
          <div className="flex items-center gap-2 border-b border-fd-border px-4 py-2.5">
            <span className="size-3 rounded-full bg-red-400/70" />
            <span className="size-3 rounded-full bg-yellow-400/70" />
            <span className="size-3 rounded-full bg-green-400/70" />
            <span className="ml-2 font-mono text-xs text-fd-muted-foreground">
              report.ts
            </span>
          </div>
          <pre className="overflow-x-auto p-4 text-sm leading-relaxed">
            <code className="font-mono text-fd-foreground">{codeSnippet}</code>
          </pre>
        </div>
        <p className="mt-3 text-center text-sm text-fd-muted-foreground">
          Open an existing workbook, make a narrow mutation, and write only if it
          still validates.
        </p>
      </section>

      {/* Feature grid */}
      <section className="mx-auto w-full max-w-5xl px-6 pb-24">
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {features.map((feature) => (
            <div
              key={feature.title}
              className="flex flex-col rounded-xl border border-fd-border bg-fd-card p-5 transition-colors hover:border-fd-primary/40"
            >
              <feature.icon className="size-6 text-fd-primary" />
              <h3 className="mt-4 font-semibold">{feature.title}</h3>
              <p className="mt-2 text-sm text-fd-muted-foreground">
                {feature.body}
              </p>
            </div>
          ))}
        </div>
      </section>
    </main>
  );
}
