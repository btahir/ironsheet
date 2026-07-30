'use client';

import {
  ArrowRight,
  Check,
  CheckCircle2,
  ChevronDown,
  CircleAlert,
  Download,
  FileArchive,
  FileSpreadsheet,
  LoaderCircle,
  LockKeyhole,
  RefreshCw,
  RotateCcw,
  ShieldCheck,
  Sparkles,
  Table2,
  Upload,
} from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { normalizeHeader } from './csv';
import type {
  CsvInspectionResult,
  RefreshResult,
  WorkbookInspectionResult,
  WorkbookTableSummary,
  WorkerRequest,
  WorkerResponse,
} from './worker-protocol';

type BusyStep = 'csv' | 'refresh' | 'sample' | 'workbook' | undefined;

export function ExcelRefresher() {
  const workerRef = useRef<Worker | null>(null);
  const workbookRef = useRef<WorkbookInspectionResult | undefined>(undefined);
  const selectedTableRef = useRef('');
  const pendingSampleCsvRef = useRef<string | undefined>(undefined);
  const [workerVersion, setWorkerVersion] = useState(0);
  const [workbook, setWorkbook] = useState<WorkbookInspectionResult>();
  const [csv, setCsv] = useState<CsvInspectionResult>();
  const [csvFileName, setCsvFileName] = useState('');
  const [selectedTableName, setSelectedTableName] = useState('');
  const [mapping, setMapping] = useState<number[]>([]);
  const [busy, setBusy] = useState<BusyStep>();
  const [progress, setProgress] = useState('');
  const [error, setError] = useState<{ code: string; message: string }>();
  const [result, setResult] = useState<RefreshResult>();
  const [downloadUrl, setDownloadUrl] = useState('');

  useEffect(() => {
    const worker = new Worker(new URL('./workbook.worker.ts', import.meta.url), { type: 'module' });
    workerRef.current = worker;
    worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
      const message = event.data;
      if (message.type === 'progress') {
        setProgress(message.label);
        return;
      }
      if (message.type === 'error') {
        setBusy(undefined);
        setProgress('');
        setError({ code: message.code, message: message.message });
        return;
      }
      if (message.type === 'workbook-inspected') {
        const nextWorkbook = message.result;
        const firstTable = nextWorkbook.tables[0]?.name ?? '';
        workbookRef.current = nextWorkbook;
        selectedTableRef.current = firstTable;
        setWorkbook(nextWorkbook);
        setSelectedTableName(firstTable);
        setBusy(undefined);
        setProgress('');
        if (pendingSampleCsvRef.current !== undefined && nextWorkbook.tables.length > 0) {
          const sampleCsv = pendingSampleCsvRef.current;
          pendingSampleCsvRef.current = undefined;
          setBusy('csv');
          setCsvFileName('fresh-sales.csv');
          postToWorker({ type: 'parse-csv', text: sampleCsv });
        }
        return;
      }
      if (message.type === 'csv-parsed') {
        setCsv(message.result);
        const selected = selectedTable(
          workbookRef.current,
          selectedTableRef.current,
        );
        setMapping(autoMapColumns(selected?.columns ?? [], message.result.headers));
        setBusy(undefined);
        setProgress('');
        return;
      }
      setResult(message.result);
      setBusy(undefined);
      setProgress('');
    };
    worker.onerror = () => {
      setBusy(undefined);
      setProgress('');
      setError({
        code: 'WORKER_FAILED',
        message: 'The browser stopped the workbook worker. Try a smaller file or close other tabs.',
      });
    };

    return () => {
      worker.terminate();
      if (workerRef.current === worker) workerRef.current = null;
    };
  }, [workerVersion]);

  useEffect(() => {
    if (result === undefined) {
      setDownloadUrl('');
      return;
    }
    const url = URL.createObjectURL(result.blob);
    setDownloadUrl(url);
    return () => URL.revokeObjectURL(url);
  }, [result]);

  const table = useMemo(
    () => selectedTable(workbook, selectedTableName),
    [selectedTableName, workbook],
  );
  const everyColumnMapped =
    table !== undefined && mapping.length === table.columns.length && mapping.every((index) => index >= 0);
  const canRefresh =
    workbook !== undefined &&
    workbook.validation.errors === 0 &&
    csv !== undefined &&
    table !== undefined &&
    everyColumnMapped &&
    busy === undefined;

  async function chooseWorkbook(file: File): Promise<void> {
    if (!/\.(?:xlsx|xlsm)$/i.test(file.name)) {
      setError({
        code: 'WORKBOOK_FORMAT',
        message: 'Choose an .xlsx or .xlsm workbook. Legacy .xls and .xlsb files are not supported.',
      });
      return;
    }
    setError(undefined);
    setResult(undefined);
    setCsv(undefined);
    setCsvFileName('');
    setMapping([]);
    setBusy('workbook');
    const bytes = await file.arrayBuffer();
    postToWorker({ type: 'inspect-workbook', bytes, fileName: file.name }, [bytes]);
  }

  async function chooseCsv(file: File): Promise<void> {
    if (!/\.csv$/i.test(file.name)) {
      setError({ code: 'CSV_FORMAT', message: 'Choose a CSV file with a header row.' });
      return;
    }
    setError(undefined);
    setResult(undefined);
    setBusy('csv');
    setCsvFileName(file.name);
    postToWorker({ type: 'parse-csv', text: await file.text() });
  }

  async function trySample(): Promise<void> {
    setBusy('sample');
    setError(undefined);
    try {
      const [workbookResponse, csvResponse] = await Promise.all([
        fetch('/samples/sales-report.xlsx'),
        fetch('/samples/fresh-sales.csv'),
      ]);
      if (!workbookResponse.ok || !csvResponse.ok) {
        throw new Error('The sample files could not be loaded.');
      }
      const workbookBlob = await workbookResponse.blob();
      pendingSampleCsvRef.current = await csvResponse.text();
      await chooseWorkbook(
        new File([workbookBlob], 'sales-report.xlsx', {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
      );
    } catch (sampleError) {
      setBusy(undefined);
      setError({
        code: 'SAMPLE_FAILED',
        message: sampleError instanceof Error ? sampleError.message : 'The sample could not be loaded.',
      });
    }
  }

  function changeTable(name: string): void {
    selectedTableRef.current = name;
    setSelectedTableName(name);
    setResult(undefined);
    const nextTable = selectedTable(workbook, name);
    setMapping(autoMapColumns(nextTable?.columns ?? [], csv?.headers ?? []));
  }

  function refreshWorkbook(): void {
    if (!canRefresh || table === undefined) return;
    setError(undefined);
    setResult(undefined);
    setBusy('refresh');
    postToWorker({ type: 'refresh', mapping, tableName: table.name });
  }

  function reset(): void {
    workbookRef.current = undefined;
    selectedTableRef.current = '';
    pendingSampleCsvRef.current = undefined;
    setWorkbook(undefined);
    setCsv(undefined);
    setCsvFileName('');
    setSelectedTableName('');
    setMapping([]);
    setBusy(undefined);
    setProgress('');
    setError(undefined);
    setResult(undefined);
    setWorkerVersion((value) => value + 1);
  }

  function postToWorker(request: WorkerRequest, transfer: Transferable[] = []): void {
    workerRef.current?.postMessage(request, transfer);
  }

  return (
    <section
      id="refresher"
      className="relative overflow-hidden rounded-[1.75rem] border border-fd-border bg-fd-card shadow-2xl shadow-teal-950/10"
      aria-labelledby="tool-title"
    >
      <div className="flex flex-col gap-3 border-b border-fd-border bg-gradient-to-r from-teal-950 to-slate-950 px-5 py-5 text-white sm:flex-row sm:items-center sm:justify-between sm:px-7">
        <div>
          <div className="mb-1 flex items-center gap-2 text-xs font-semibold uppercase tracking-[0.18em] text-teal-300">
            <ShieldCheck className="size-4" /> Local workbook tool
          </div>
          <h2 id="tool-title" className="text-xl font-semibold tracking-tight sm:text-2xl">
            Refresh an Excel table
          </h2>
        </div>
        <div className="flex items-center gap-2 rounded-full border border-white/15 bg-white/10 px-3 py-1.5 text-xs text-slate-200">
          <LockKeyhole className="size-3.5 text-teal-300" /> Files stay in this browser
        </div>
      </div>

      <div className="space-y-6 p-5 sm:p-7">
        {busy !== undefined && (
          <div className="flex items-center gap-3 rounded-xl border border-teal-500/20 bg-teal-500/10 px-4 py-3 text-sm" role="status" aria-live="polite">
            <LoaderCircle className="size-5 animate-spin text-teal-600 dark:text-teal-300" />
            <span>{progress || (busy === 'sample' ? 'Loading the sample files' : 'Preparing your file')}</span>
          </div>
        )}

        {error !== undefined && (
          <div className="rounded-xl border border-red-500/30 bg-red-500/10 px-4 py-3" role="alert">
            <div className="flex gap-3">
              <CircleAlert className="mt-0.5 size-5 shrink-0 text-red-600 dark:text-red-300" />
              <div>
                <p className="font-medium">We couldn&apos;t complete that step</p>
                <p className="mt-1 text-sm text-fd-muted-foreground">{error.message}</p>
                <p className="mt-2 font-mono text-[11px] text-fd-muted-foreground">{error.code}</p>
              </div>
            </div>
          </div>
        )}

        <ToolStep number="1" title="Choose your existing workbook" complete={workbook !== undefined}>
          {workbook === undefined ? (
            <FilePicker
              accept=".xlsx,.xlsm"
              description="XLSX or XLSM · processed locally · up to 25 MB compressed"
              icon={<FileSpreadsheet className="size-7" />}
              label="Choose Excel workbook"
              onFile={(file) => void chooseWorkbook(file)}
            />
          ) : (
            <div className="rounded-xl border border-fd-border bg-fd-background/60 p-4">
              <div className="flex items-start justify-between gap-4">
                <div className="flex min-w-0 gap-3">
                  <div className="flex size-10 shrink-0 items-center justify-center rounded-lg bg-teal-500/10 text-teal-700 dark:text-teal-300">
                    <FileArchive className="size-5" />
                  </div>
                  <div className="min-w-0">
                    <p className="truncate font-medium">{workbook.fileName}</p>
                    <p className="mt-0.5 text-xs text-fd-muted-foreground">
                      {formatBytes(workbook.archive.compressedBytes)} · {workbook.tables.length}{' '}
                      {workbook.tables.length === 1 ? 'table' : 'tables'} · {workbook.validation.errors} validation errors
                    </p>
                  </div>
                </div>
                <button type="button" onClick={reset} className="text-xs font-medium text-fd-muted-foreground hover:text-fd-foreground">
                  Change
                </button>
              </div>
              <FeatureChips features={workbook.features} />
            </div>
          )}
          {workbook !== undefined && workbook.tables.length === 0 && (
            <p className="mt-3 rounded-lg border border-amber-500/30 bg-amber-500/10 px-3 py-2 text-sm">
              No Excel table was found. In Excel, select the data and choose <strong>Insert → Table</strong>, then save the workbook and try again.
            </p>
          )}
          {workbook !== undefined && workbook.validation.errors > 0 && (
            <p className="mt-3 rounded-lg border border-red-500/30 bg-red-500/10 px-3 py-2 text-sm">
              This workbook already has structural validation errors, so Ironsheet will not write an updated copy.
            </p>
          )}
        </ToolStep>

        {workbook !== undefined && workbook.tables.length > 0 && (
          <ToolStep number="2" title="Choose the Excel table" complete={table !== undefined}>
            <label className="block">
              <span className="mb-2 block text-sm font-medium">Table to refresh</span>
              <div className="relative">
                <select
                  value={selectedTableName}
                  onChange={(event) => changeTable(event.target.value)}
                  className="h-12 w-full appearance-none rounded-xl border border-fd-border bg-fd-background px-4 pr-10 text-sm outline-none ring-teal-600 transition focus:ring-2"
                >
                  {workbook.tables.map((candidate) => (
                    <option key={candidate.name} value={candidate.name}>
                      {candidate.displayName} — {candidate.sheetName} ({candidate.rowCount.toLocaleString()} rows)
                    </option>
                  ))}
                </select>
                <ChevronDown className="pointer-events-none absolute right-4 top-1/2 size-4 -translate-y-1/2 text-fd-muted-foreground" />
              </div>
            </label>
          </ToolStep>
        )}

        {table !== undefined && (
          <ToolStep number="3" title="Choose the fresh CSV" complete={csv !== undefined}>
            {csv === undefined ? (
              <FilePicker
                accept=".csv,text/csv"
                description={`Header row required · up to 50,000 rows or 250,000 cells`}
                icon={<Table2 className="size-7" />}
                label="Choose fresh CSV"
                onFile={(file) => void chooseCsv(file)}
              />
            ) : (
              <div className="flex items-center justify-between gap-4 rounded-xl border border-fd-border bg-fd-background/60 p-4">
                <div className="min-w-0">
                  <p className="truncate font-medium">{csvFileName}</p>
                  <p className="mt-0.5 text-xs text-fd-muted-foreground">
                    {csv.rowCount.toLocaleString()} rows · {csv.headers.length} columns · {csv.delimiter} delimited
                  </p>
                </div>
                <label className="cursor-pointer text-xs font-medium text-fd-muted-foreground hover:text-fd-foreground">
                  Change
                  <input
                    type="file"
                    accept=".csv,text/csv"
                    className="sr-only"
                    onChange={(event) => {
                      const file = event.target.files?.[0];
                      if (file !== undefined) void chooseCsv(file);
                    }}
                  />
                </label>
              </div>
            )}
          </ToolStep>
        )}

        {table !== undefined && csv !== undefined && (
          <ToolStep number="4" title="Review the column match" complete={everyColumnMapped}>
            <div className="overflow-hidden rounded-xl border border-fd-border">
              <div className="grid grid-cols-[minmax(0,1fr)_2rem_minmax(0,1fr)] bg-fd-muted/40 px-4 py-2 text-[11px] font-semibold uppercase tracking-wider text-fd-muted-foreground">
                <span>Excel table</span><span /><span>CSV source</span>
              </div>
              {table.columns.map((column, index) => (
                <div key={`${column}-${index}`} className="grid grid-cols-[minmax(0,1fr)_2rem_minmax(0,1fr)] items-center border-t border-fd-border px-4 py-3 text-sm">
                  <span className="truncate font-medium">{column}</span>
                  <ArrowRight className="size-4 text-fd-muted-foreground" />
                  <select
                    aria-label={`CSV column for ${column}`}
                    value={mapping[index] ?? -1}
                    onChange={(event) => {
                      const next = [...mapping];
                      next[index] = Number(event.target.value);
                      setMapping(next);
                      setResult(undefined);
                    }}
                    className="min-w-0 rounded-lg border border-fd-border bg-fd-background px-3 py-2 outline-none ring-teal-600 focus:ring-2"
                  >
                    <option value={-1}>Select a CSV column</option>
                    {csv.headers.map((header, headerIndex) => (
                      <option key={`${header}-${headerIndex}`} value={headerIndex}>{header}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>
            {csv.warnings.map((warning) => (
              <p key={warning} className="mt-2 text-xs text-amber-700 dark:text-amber-300">{warning}</p>
            ))}
            <div className="mt-4 grid gap-3 sm:grid-cols-2">
              <PreviewTable title={`Existing · ${table.rowCount.toLocaleString()} rows`} headers={table.columns} rows={table.preview} />
              <PreviewTable title={`Fresh CSV · ${csv.rowCount.toLocaleString()} rows`} headers={csv.headers} rows={csv.preview} />
            </div>
          </ToolStep>
        )}

        {table !== undefined && csv !== undefined && (
          <div className="rounded-2xl border border-teal-600/20 bg-teal-500/[0.06] p-4 sm:p-5">
            <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <p className="font-medium">Ready to create an updated copy</p>
                <p className="mt-1 text-sm text-fd-muted-foreground">
                  Replace {table.rowCount.toLocaleString()} rows with {csv.rowCount.toLocaleString()} rows. Your original file stays untouched.
                </p>
              </div>
              <button
                type="button"
                onClick={refreshWorkbook}
                disabled={!canRefresh}
                className="inline-flex h-11 shrink-0 items-center justify-center gap-2 rounded-xl bg-teal-700 px-5 text-sm font-semibold text-white shadow-lg shadow-teal-900/15 transition hover:bg-teal-800 disabled:cursor-not-allowed disabled:opacity-40 dark:bg-teal-500 dark:text-teal-950 dark:hover:bg-teal-400"
              >
                <RefreshCw className="size-4" /> Create updated workbook
              </button>
            </div>
          </div>
        )}

        {result !== undefined && (
          <div className="rounded-2xl border border-emerald-500/30 bg-emerald-500/10 p-5" role="status" aria-live="polite">
            <div className="flex items-start gap-3">
              <CheckCircle2 className="mt-0.5 size-6 shrink-0 text-emerald-600 dark:text-emerald-300" />
              <div className="min-w-0 flex-1">
                <h3 className="text-lg font-semibold">Your updated workbook is ready</h3>
                <p className="mt-1 text-sm text-fd-muted-foreground">
                  Ironsheet replaced {result.oldRowCount.toLocaleString()} rows with {result.newRowCount.toLocaleString()} rows in {result.tableName} and found {result.validation.errors} validation errors.
                </p>
                <div className="mt-4 flex flex-wrap gap-3">
                  <a href={downloadUrl} download={result.fileName} className="inline-flex h-11 items-center gap-2 rounded-xl bg-emerald-700 px-5 text-sm font-semibold text-white transition hover:bg-emerald-800">
                    <Download className="size-4" /> Download {result.fileName}
                  </a>
                  <button type="button" onClick={reset} className="inline-flex h-11 items-center gap-2 rounded-xl border border-fd-border bg-fd-card px-4 text-sm font-medium hover:bg-fd-accent">
                    <RotateCcw className="size-4" /> Start over
                  </button>
                </div>
                {result.diagnostics.length > 0 && (
                  <details className="mt-4 text-sm">
                    <summary className="cursor-pointer font-medium">Technical notes ({result.diagnostics.length})</summary>
                    <ul className="mt-2 space-y-2 text-fd-muted-foreground">
                      {result.diagnostics.map((diagnostic, index) => (
                        <li key={`${diagnostic.code}-${index}`}><code className="text-xs">{diagnostic.code}</code> — {diagnostic.message}</li>
                      ))}
                    </ul>
                  </details>
                )}
              </div>
            </div>
          </div>
        )}

        {workbook === undefined && busy === undefined && (
          <div className="flex flex-col items-center justify-between gap-3 border-t border-fd-border pt-5 text-center sm:flex-row sm:text-left">
            <div className="flex items-center gap-2 text-sm text-fd-muted-foreground">
              <Sparkles className="size-4 text-teal-600 dark:text-teal-300" /> No workbook handy? Use our safe sample.
            </div>
            <button type="button" onClick={() => void trySample()} className="inline-flex items-center gap-2 rounded-lg px-3 py-2 text-sm font-semibold text-teal-700 hover:bg-teal-500/10 dark:text-teal-300">
              Try sample files <ArrowRight className="size-4" />
            </button>
          </div>
        )}
      </div>
    </section>
  );
}

function ToolStep({ children, complete, number, title }: { children: React.ReactNode; complete: boolean; number: string; title: string }) {
  return (
    <div>
      <div className="mb-3 flex items-center gap-3">
        <span className={`flex size-7 items-center justify-center rounded-full text-xs font-bold ${complete ? 'bg-teal-700 text-white dark:bg-teal-400 dark:text-teal-950' : 'border border-fd-border bg-fd-background text-fd-muted-foreground'}`}>
          {complete ? <Check className="size-4" /> : number}
        </span>
        <h3 className="font-semibold">{title}</h3>
      </div>
      <div className="pl-0 sm:pl-10">{children}</div>
    </div>
  );
}

function FilePicker({ accept, description, icon, label, onFile }: { accept: string; description: string; icon: React.ReactNode; label: string; onFile(file: File): void }) {
  const [dragging, setDragging] = useState(false);
  return (
    <label
      className={`group flex cursor-pointer flex-col items-center justify-center rounded-xl border border-dashed px-5 py-7 text-center transition ${dragging ? 'border-teal-500 bg-teal-500/10' : 'border-fd-border bg-fd-background/40 hover:border-teal-500/60 hover:bg-teal-500/[0.04]'}`}
      onDragEnter={(event) => { event.preventDefault(); setDragging(true); }}
      onDragOver={(event) => event.preventDefault()}
      onDragLeave={() => setDragging(false)}
      onDrop={(event) => {
        event.preventDefault();
        setDragging(false);
        const file = event.dataTransfer.files[0];
        if (file !== undefined) onFile(file);
      }}
    >
      <span className="mb-3 flex size-12 items-center justify-center rounded-xl bg-teal-500/10 text-teal-700 transition group-hover:scale-105 dark:text-teal-300">{icon}</span>
      <span className="inline-flex items-center gap-2 font-semibold"><Upload className="size-4" /> {label}</span>
      <span className="mt-1 text-xs text-fd-muted-foreground">{description}</span>
      <input type="file" accept={accept} className="sr-only" onChange={(event) => {
        const file = event.target.files?.[0];
        if (file !== undefined) onFile(file);
      }} />
    </label>
  );
}

function FeatureChips({ features }: { features: WorkbookInspectionResult['features'] }) {
  const labels: Array<[keyof typeof features, string, string]> = [
    ['formulaCells', 'formula', 'formulas'],
    ['charts', 'chart', 'charts'],
    ['pivotTables', 'pivot', 'pivots'],
    ['macros', 'macro', 'macros'],
    ['drawings', 'drawing', 'drawings'],
    ['dataValidations', 'validation', 'validations'],
    ['hiddenSheets', 'hidden sheet', 'hidden sheets'],
  ];
  const visible = labels.filter(([key]) => features[key] > 0);
  if (visible.length === 0) return null;
  return <div className="mt-3 flex flex-wrap gap-1.5">{visible.map(([key, singular, plural]) => <span key={String(key)} className="rounded-full border border-fd-border bg-fd-card px-2 py-1 text-[11px] text-fd-muted-foreground">{features[key]} {features[key] === 1 ? singular : plural}</span>)}</div>;
}

function PreviewTable({ headers, rows, title }: { headers: string[]; rows: string[][]; title: string }) {
  const visibleHeaders = headers.slice(0, 3);
  return (
    <div className="min-w-0 overflow-hidden rounded-xl border border-fd-border bg-fd-background/50">
      <p className="border-b border-fd-border px-3 py-2 text-xs font-semibold text-fd-muted-foreground">{title}</p>
      <div className="overflow-x-auto">
        <table className="w-full min-w-64 text-left text-[11px]">
          <thead className="bg-fd-muted/30"><tr>{visibleHeaders.map((header, index) => <th key={`${header}-${index}`} className="max-w-32 truncate px-3 py-2 font-semibold">{header}</th>)}</tr></thead>
          <tbody>{rows.slice(0, 3).map((row, rowIndex) => <tr key={rowIndex} className="border-t border-fd-border">{visibleHeaders.map((_, columnIndex) => <td key={columnIndex} className="max-w-32 truncate px-3 py-2 text-fd-muted-foreground">{row[columnIndex] || '—'}</td>)}</tr>)}</tbody>
        </table>
      </div>
    </div>
  );
}

function selectedTable(workbook: WorkbookInspectionResult | undefined, name: string): WorkbookTableSummary | undefined {
  return workbook?.tables.find((table) => table.name === name);
}

function autoMapColumns(tableColumns: string[], csvHeaders: string[]): number[] {
  const normalizedCsv = csvHeaders.map(normalizeHeader);
  return tableColumns.map((column) => normalizedCsv.indexOf(normalizeHeader(column)));
}

function formatBytes(bytes: number): string {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
