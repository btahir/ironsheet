import type { Metadata } from 'next';
import Link from 'next/link';
import {
  ArrowRight,
  Braces,
  CheckCircle2,
  FileCheck2,
  FileDown,
  FileLock2,
  GitFork,
  MonitorCog,
  ShieldCheck,
  TableProperties,
} from 'lucide-react';
import { ExcelRefresher } from './excel-refresher';
import { gitConfig, siteUrl } from '@/lib/shared';

const canonicalPath = '/tools/update-excel-from-csv';
const canonicalUrl = new URL(canonicalPath, siteUrl).toString();

export const metadata: Metadata = {
  title: 'Update Excel from CSV Without Breaking Formatting',
  description:
    'Replace an Excel table with fresh CSV data and download the updated XLSX or XLSM. Runs locally in your browser—no file upload, account, or server.',
  alternates: { canonical: canonicalPath },
  openGraph: {
    title: 'Update Excel from CSV Without Breaking Your Workbook',
    description:
      'Refresh a real Excel table, validate the result, and download a new copy. Your files stay in your browser.',
    type: 'website',
    url: canonicalPath,
    images: ['/ironsheet-tool-og.png'],
  },
  twitter: {
    card: 'summary_large_image',
    title: 'Update Excel from CSV — Private, In-Browser',
    description: 'Replace table data without rebuilding or uploading your workbook.',
    images: ['/ironsheet-tool-og.png'],
  },
};

const githubUrl = `https://github.com/${gitConfig.user}/${gitConfig.repo}`;

const faqs = [
  {
    question: 'Does Ironsheet upload my workbook?',
    answer:
      'No. Vercel serves this page and its code, but your workbook and CSV are read, changed, and written inside a browser worker on your device. There is no file-upload endpoint or saved result URL.',
  },
  {
    question: 'What happens when I close the tab?',
    answer:
      'The in-memory workbook, CSV, and temporary download disappear with the browser session. The tool does not store them in a database or browser storage.',
  },
  {
    question: 'Which files work?',
    answer:
      'This first version accepts XLSX and XLSM workbooks containing a real Excel table, plus a CSV with a header row. Legacy XLS, XLSB, encrypted workbooks, and arbitrary worksheet ranges are not supported yet.',
  },
  {
    question: 'What happens to formulas, charts, pivots, and macros?',
    answer:
      'Ironsheet edits the targeted table inside the existing OOXML package instead of rebuilding the workbook. Untouched parts are preserved, macros are never executed, and Excel is asked to recalculate formulas when the result opens. Cached chart or pivot data may refresh in Excel.',
  },
  {
    question: 'How large can the files be?',
    answer:
      'The current browser guardrails allow workbooks up to 25 MB compressed and CSV updates up to 50,000 rows or 250,000 cells. Actual speed depends on workbook complexity and your device.',
  },
  {
    question: 'Does this replace Excel or Google Sheets?',
    answer:
      'No. This is a focused workbook refresh tool, not a general spreadsheet editor. It is designed for the recurring job of putting fresh exported data into an existing Excel report without recreating that report.',
  },
];

export default function UpdateExcelFromCsvPage() {
  const structuredData = {
    '@context': 'https://schema.org',
    '@type': 'WebApplication',
    name: 'Ironsheet Excel Table Refresher',
    applicationCategory: 'BusinessApplication',
    browserRequirements: 'Requires a modern browser with CompressionStream support',
    description:
      'A private, browser-only tool for replacing an Excel table with fresh CSV data while preserving the existing workbook.',
    isAccessibleForFree: true,
    operatingSystem: 'Any modern desktop operating system',
    offers: { '@type': 'Offer', price: 0, priceCurrency: 'USD' },
    url: canonicalUrl,
  };

  return (
    <main className="relative flex flex-1 flex-col overflow-hidden">
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(structuredData).replace(/</g, '\\u003c') }}
      />

      <div className="pointer-events-none absolute inset-x-0 top-0 -z-10 h-[52rem] bg-[radial-gradient(circle_at_18%_15%,rgba(20,184,166,0.16),transparent_34%),radial-gradient(circle_at_85%_12%,rgba(15,118,110,0.12),transparent_28%)]" />
      <section className="mx-auto grid w-full max-w-6xl gap-12 px-6 pb-16 pt-16 lg:grid-cols-[0.82fr_1.18fr] lg:items-start lg:pt-24">
        <div className="lg:sticky lg:top-24">
          <div className="inline-flex items-center gap-2 rounded-full border border-teal-600/25 bg-teal-500/10 px-3 py-1.5 text-xs font-semibold text-teal-800 dark:text-teal-200">
            <ShieldCheck className="size-4" /> Free · private · browser-only
          </div>
          <h1 className="mt-6 text-balance text-4xl font-bold leading-[1.05] tracking-[-0.045em] sm:text-5xl lg:text-[3.6rem]">
            Update your Excel report without breaking the workbook.
          </h1>
          <p className="mt-6 max-w-xl text-pretty text-lg leading-8 text-fd-muted-foreground">
            Choose an existing Excel workbook and a fresh CSV. Ironsheet replaces the data table, validates the result, and gives you back a new copy—with the rest of your workbook intact.
          </p>
          <div className="mt-7 grid gap-3 text-sm sm:grid-cols-2 lg:grid-cols-1 xl:grid-cols-2">
            {[
              'Files never leave your browser',
              'No account or installation',
              'Original file stays untouched',
              'XLSX and XLSM supported',
            ].map((item) => (
              <div key={item} className="flex items-center gap-2.5">
                <CheckCircle2 className="size-4 shrink-0 text-teal-700 dark:text-teal-300" />
                <span>{item}</span>
              </div>
            ))}
          </div>
          <a href="#how-it-works" className="mt-8 inline-flex items-center gap-2 text-sm font-semibold text-teal-700 hover:gap-3 dark:text-teal-300">
            See exactly what happens <ArrowRight className="size-4 transition-all" />
          </a>
        </div>

        <ExcelRefresher />
      </section>

      <section id="how-it-works" className="border-y border-fd-border bg-fd-muted/20">
        <div className="mx-auto w-full max-w-6xl px-6 py-20">
          <div className="max-w-2xl">
            <p className="text-sm font-semibold uppercase tracking-[0.18em] text-teal-700 dark:text-teal-300">A narrow edit, not a rebuild</p>
            <h2 className="mt-3 text-balance text-3xl font-bold tracking-tight sm:text-4xl">Your workbook is more than a grid of cells.</h2>
            <p className="mt-4 text-lg leading-8 text-fd-muted-foreground">
              Real reports contain formulas, styles, charts, pivots, named ranges, macros, relationships, and application-specific XML. Recreating a workbook from extracted values can quietly discard those parts. Ironsheet changes the table you selected and preserves untouched package entries.
            </p>
          </div>
          <div className="mt-10 grid gap-4 md:grid-cols-3">
            {[
              { icon: TableProperties, number: '01', title: 'Choose the target', body: 'Ironsheet inspects the workbook locally and shows the Excel tables it can safely update.' },
              { icon: MonitorCog, number: '02', title: 'Match fresh data', body: 'Review the CSV headers and explicitly map every destination column before anything changes.' },
              { icon: FileDown, number: '03', title: 'Validate and download', body: 'The updated package is checked before a temporary local download is created.' },
            ].map((step) => (
              <article key={step.number} className="rounded-2xl border border-fd-border bg-fd-card p-6">
                <div className="flex items-center justify-between">
                  <step.icon className="size-6 text-teal-700 dark:text-teal-300" />
                  <span className="font-mono text-xs text-fd-muted-foreground">{step.number}</span>
                </div>
                <h3 className="mt-8 text-lg font-semibold">{step.title}</h3>
                <p className="mt-2 text-sm leading-6 text-fd-muted-foreground">{step.body}</p>
              </article>
            ))}
          </div>
        </div>
      </section>

      <section className="mx-auto grid w-full max-w-6xl gap-12 px-6 py-20 lg:grid-cols-2 lg:items-center">
        <div className="rounded-[1.75rem] border border-teal-600/20 bg-gradient-to-br from-teal-950 to-slate-950 p-7 text-white shadow-2xl shadow-teal-950/15 sm:p-9">
          <div className="flex items-center gap-2 text-sm font-semibold text-teal-300"><FileLock2 className="size-5" /> Private by architecture</div>
          <h2 className="mt-4 text-3xl font-bold tracking-tight">No upload disguised as an upload.</h2>
          <p className="mt-4 leading-7 text-slate-300">
            The page code comes from Vercel. Your files do not go back. The browser reads them into an isolated worker, Ironsheet performs the change in memory, and the browser creates the download.
          </p>
          <div className="mt-7 grid gap-3 sm:grid-cols-2">
            {['No file API route', 'No database', 'No saved result URL', 'No workbook analytics'].map((item) => (
              <div key={item} className="flex items-center gap-2 rounded-lg border border-white/10 bg-white/[0.06] px-3 py-2.5 text-sm text-slate-200">
                <CheckCircle2 className="size-4 text-teal-300" /> {item}
              </div>
            ))}
          </div>
        </div>
        <div>
          <p className="text-sm font-semibold uppercase tracking-[0.18em] text-teal-700 dark:text-teal-300">Powered by the open-source engine</p>
          <h2 className="mt-3 text-3xl font-bold tracking-tight sm:text-4xl">The demo and the library improve each other.</h2>
          <p className="mt-4 text-lg leading-8 text-fd-muted-foreground">
            This tool runs through the same public browser package available to developers. The archive checks and validation-gated browser write were built for this workflow, then promoted into Ironsheet so other applications can use them too.
          </p>
          <div className="mt-7 flex flex-wrap gap-3">
            <Link href="/docs/guides/browser" className="inline-flex h-11 items-center gap-2 rounded-xl bg-fd-primary px-5 text-sm font-semibold text-fd-primary-foreground hover:opacity-90">
              <Braces className="size-4" /> Browser package
            </Link>
            <a href={githubUrl} target="_blank" rel="noreferrer" className="inline-flex h-11 items-center gap-2 rounded-xl border border-fd-border bg-fd-card px-5 text-sm font-semibold hover:bg-fd-accent">
              <GitFork className="size-4" /> View source
            </a>
          </div>
        </div>
      </section>

      <section className="border-y border-fd-border bg-fd-muted/20">
        <div className="mx-auto w-full max-w-6xl px-6 py-20">
          <div className="mx-auto max-w-3xl text-center">
            <FileCheck2 className="mx-auto size-7 text-teal-700 dark:text-teal-300" />
            <h2 className="mt-4 text-3xl font-bold tracking-tight sm:text-4xl">Built to fail safely.</h2>
            <p className="mt-4 text-lg leading-8 text-fd-muted-foreground">The tool checks archive size before opening, refuses unsafe table expansion, validates the finished workbook, and withholds the download if structural errors are found.</p>
          </div>
        </div>
      </section>

      <section className="mx-auto w-full max-w-4xl px-6 py-20">
        <div className="text-center">
          <p className="text-sm font-semibold uppercase tracking-[0.18em] text-teal-700 dark:text-teal-300">Questions, answered plainly</p>
          <h2 className="mt-3 text-3xl font-bold tracking-tight sm:text-4xl">Frequently asked questions</h2>
        </div>
        <div className="mt-10 divide-y divide-fd-border rounded-2xl border border-fd-border bg-fd-card px-5 sm:px-7">
          {faqs.map((faq) => (
            <details key={faq.question} className="group py-5">
              <summary className="flex cursor-pointer list-none items-center justify-between gap-4 font-semibold">
                {faq.question}
                <span className="text-xl font-normal text-fd-muted-foreground transition group-open:rotate-45">+</span>
              </summary>
              <p className="mt-3 max-w-3xl pr-8 text-sm leading-7 text-fd-muted-foreground">{faq.answer}</p>
            </details>
          ))}
        </div>
      </section>
    </main>
  );
}
