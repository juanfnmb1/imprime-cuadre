import { useMemo, useRef, useState } from 'react';
import { formatMoney, parseWorkbook, sumDaysTotals } from './parser';
import {
  buildDayPdf,
  buildGrandTotalPdf,
  dayFilename,
  downloadAllPdfs,
  downloadSelectionZip,
} from './pdf';
import type { DayTotals, ParsedWorkbook, Transaction, TransactionIssue } from './types';

export function App() {
  const [data, setData] = useState<ParsedWorkbook | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const [error, setError] = useState<string>('');
  const [busy, setBusy] = useState<boolean>(false);
  const inputRef = useRef<HTMLInputElement>(null);

  async function handleFile(file: File) {
    setError('');
    setBusy(true);
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseWorkbook(buf);
      setData(parsed);
      setFileName(file.name);
    } catch (e) {
      setError(e instanceof Error ? e.message : 'No se pudo leer el archivo.');
      setData(null);
    } finally {
      setBusy(false);
    }
  }

  function onPickClick() {
    inputRef.current?.click();
  }

  function reset() {
    setData(null);
    setFileName('');
    setError('');
    if (inputRef.current) inputRef.current.value = '';
  }

  return (
    <div className="max-w-6xl mx-auto px-margin py-xl space-y-xl">
      <div className="flex flex-col items-center justify-center text-center space-y-xs mb-lg">
        <span className="text-h1 font-h1 font-bold text-secondary tracking-tighter">
          IMPRIME CUADRE
        </span>
        <p className="text-label-sm text-on-surface-variant uppercase tracking-[0.2em]">
          Procesador de Cuadres Diarios
        </p>
      </div>

      <Uploader
        fileName={fileName}
        busy={busy}
        error={error}
        onPickClick={onPickClick}
        onDrop={handleFile}
        onReset={reset}
      />

      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void handleFile(f);
        }}
        hidden
      />

      {data && <Preview data={data} />}

      <footer className="pt-xl text-center">
        <p className="text-label-sm text-on-surface-variant opacity-50 uppercase tracking-widest">
          Imprime Cuadre · Vista previa y descarga de PDFs
        </p>
      </footer>
    </div>
  );
}

function Uploader({
  fileName,
  busy,
  error,
  onPickClick,
  onDrop,
  onReset,
}: {
  fileName: string;
  busy: boolean;
  error: string;
  onPickClick: () => void;
  onDrop: (f: File) => void;
  onReset: () => void;
}) {
  return (
    <section className="w-full">
      <div
        className="bg-surface-container border border-outline-variant p-xl flex flex-col items-center justify-center text-center space-y-md relative overflow-hidden group cursor-pointer hover:border-secondary transition-colors"
        onClick={onPickClick}
        onDragOver={(e) => {
          e.preventDefault();
        }}
        onDrop={(e) => {
          e.preventDefault();
          const f = e.dataTransfer.files?.[0];
          if (f) onDrop(f);
        }}
      >
        <div className="absolute inset-0 opacity-10 pointer-events-none">
          <div className="absolute -right-10 -top-10 w-64 h-64 bg-secondary blur-[100px]" />
        </div>

        <div className="relative z-10">
          <div className="w-16 h-16 bg-surface-container-high rounded-full flex items-center justify-center mb-md border border-outline-variant mx-auto group-hover:scale-110 transition-transform">
            <span className="material-symbols-outlined text-secondary text-[32px]">
              upload_file
            </span>
          </div>
          <h1 className="text-h1 font-h1 text-on-surface mb-xs">
            {fileName ? fileName : 'Subir Excel del Cuadre'}
          </h1>
          <p className="text-body-md text-on-surface-variant max-w-md mx-auto">
            {busy
              ? 'Procesando archivo...'
              : 'Arrastra y suelta tu archivo Excel aquí, o haz clic para buscarlo.'}
          </p>
        </div>

        <div className="flex gap-sm pt-md relative z-10">
          <button
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              onPickClick();
            }}
            className="border border-secondary text-secondary px-lg py-base font-label-sm hover:bg-secondary/10 transition-colors"
          >
            Buscar Archivo
          </button>
          {fileName && (
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                onReset();
              }}
              className="bg-primary text-on-primary px-lg py-base font-label-sm hover:opacity-90 transition-opacity"
            >
              Cambiar Archivo
            </button>
          )}
        </div>
      </div>

      {error && (
        <div className="mt-md bg-error-container/40 border border-error text-on-error-container px-md py-sm text-body-md">
          {error}
        </div>
      )}
    </section>
  );
}

function Preview({ data }: { data: ParsedWorkbook }) {
  const { days, transactionsByDay, grandTotals, issues } = data;
  const [downloadingAll, setDownloadingAll] = useState(false);
  const [downloadStatus, setDownloadStatus] = useState<string>('');

  const [selectionMode, setSelectionMode] = useState(false);
  const [selectedKeys, setSelectedKeys] = useState<Set<string>>(new Set());
  const [selectionDownloading, setSelectionDownloading] = useState(false);
  const [sortAscending, setSortAscending] = useState(true);

  const sortedDays = useMemo(() => {
    const asc = [...days].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
    return sortAscending ? asc : asc.reverse();
  }, [days, sortAscending]);

  const totalTx = useMemo(
    () => Object.values(transactionsByDay).reduce((acc, list) => acc + list.length, 0),
    [transactionsByDay],
  );

  const period = useMemo(() => {
    if (days.length === 0) return null;
    const sorted = [...days].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
    return { first: sorted[0].rawDate, last: sorted[sorted.length - 1].rawDate };
  }, [days]);

  const selectedDays = useMemo(
    () => days.filter((d) => selectedKeys.has(d.dateKey)),
    [days, selectedKeys],
  );

  function toggleDay(dateKey: string) {
    setSelectedKeys((prev) => {
      const next = new Set(prev);
      if (next.has(dateKey)) next.delete(dateKey);
      else next.add(dateKey);
      return next;
    });
  }

  const allSelected = days.length > 0 && selectedKeys.size === days.length;

  function toggleAll() {
    setSelectedKeys((prev) =>
      prev.size === days.length ? new Set() : new Set(days.map((d) => d.dateKey)),
    );
  }

  function startSelection() {
    setSelectionMode(true);
    setSelectedKeys(new Set());
    setDownloadStatus('');
  }

  function exitSelection() {
    setSelectionMode(false);
    setSelectedKeys(new Set());
  }

  async function printSelection() {
    if (selectedDays.length === 0) return;
    setSelectionDownloading(true);
    try {
      const result = await downloadSelectionZip(selectedDays, transactionsByDay);
      setDownloadStatus(
        `Descargado: ${result.folderName}.zip (${result.count} PDFs adentro, incluyendo el resumen).`,
      );
      exitSelection();
    } catch (e) {
      setDownloadStatus(e instanceof Error ? `Error: ${e.message}` : 'Error al generar el zip.');
    } finally {
      setSelectionDownloading(false);
    }
  }

  async function downloadAll() {
    setDownloadingAll(true);
    setDownloadStatus('');
    try {
      const result = await downloadAllPdfs(days, transactionsByDay, grandTotals);
      setDownloadStatus(
        `Descargado: ${result.folderName}.zip (${result.count} PDFs adentro). Doble clic para abrir.`,
      );
    } catch (e) {
      setDownloadStatus(e instanceof Error ? `Error: ${e.message}` : 'Error desconocido.');
    } finally {
      setDownloadingAll(false);
    }
  }

  function downloadGrand() {
    const doc = buildGrandTotalPdf(grandTotals, period ?? undefined);
    doc.save('Totales Generales.pdf');
  }

  return (
    <>
      {issues.length > 0 && <IssuesBanner issues={issues} />}

      <button
        onClick={downloadAll}
        disabled={downloadingAll}
        className="w-full bg-secondary text-on-secondary px-lg py-md font-label-sm font-bold uppercase tracking-wider active:scale-[0.99] hover:opacity-90 transition-all flex items-center justify-center gap-sm disabled:opacity-60 text-[14px]"
      >
        <span className="material-symbols-outlined text-[22px]">folder_zip</span>
        {downloadingAll
          ? 'Descargando...'
          : `Descargar Todos los PDFs (${days.length + 1})`}
      </button>

      {downloadStatus && (
        <p className="text-label-sm text-on-surface-variant -mt-sm">{downloadStatus}</p>
      )}

      {period && (
        <div className="border-l-4 border-secondary pl-md">
          <p className="text-label-sm text-on-surface-variant uppercase tracking-widest">
            Periodo
          </p>
          <p className="text-h2 font-h2 text-on-surface mt-xs">
            De {period.first} a {period.last}
          </p>
        </div>
      )}

      <GrandTotalCard rows={grandTotals} onDownload={downloadGrand} />

      <section className="space-y-md">
        <div className="flex justify-between items-center border-b border-outline-variant pb-base flex-wrap gap-sm">
          <h2 className="text-h2 font-h2 text-on-surface">Por Día</h2>
          <div className="flex items-center gap-sm flex-wrap">
            <span className="text-label-sm text-on-surface-variant">
              {days.length} {days.length === 1 ? 'día' : 'días'} · {totalTx} transacciones
            </span>
            <button
              onClick={() => setSortAscending((v) => !v)}
              className="border border-outline-variant text-on-surface px-md py-xs font-label-sm font-bold hover:bg-surface-container-high transition-colors flex items-center gap-xs"
              title="Cambiar orden"
            >
              <span className="material-symbols-outlined text-[18px]">
                {sortAscending ? 'arrow_upward' : 'arrow_downward'}
              </span>
              {sortAscending ? 'Más antiguos' : 'Más recientes'}
            </button>
            {!selectionMode && (
              <button
                onClick={startSelection}
                className="border border-secondary text-secondary px-md py-xs font-label-sm font-bold hover:bg-secondary/10 transition-colors flex items-center gap-xs"
              >
                <span className="material-symbols-outlined text-[18px]">tune</span>
                Modificar Cuadre
              </button>
            )}
          </div>
        </div>

        {selectionMode && (
          <SelectionPanel
            selectedDays={selectedDays}
            totalDays={days.length}
            allSelected={allSelected}
            downloading={selectionDownloading}
            onToggleAll={toggleAll}
            onPrint={printSelection}
            onCancel={exitSelection}
          />
        )}

        <div className="file-grid">
          {sortedDays.map((day) => (
            <DayCard
              key={day.dateKey + day.rawDate}
              day={day}
              transactions={transactionsByDay[day.dateKey] ?? []}
              selectionMode={selectionMode}
              selected={selectedKeys.has(day.dateKey)}
              onToggleSelect={() => toggleDay(day.dateKey)}
            />
          ))}
        </div>
      </section>
    </>
  );
}

function SelectionPanel({
  selectedDays,
  totalDays,
  allSelected,
  downloading,
  onToggleAll,
  onPrint,
  onCancel,
}: {
  selectedDays: DayTotals[];
  totalDays: number;
  allSelected: boolean;
  downloading: boolean;
  onToggleAll: () => void;
  onPrint: () => void;
  onCancel: () => void;
}) {
  const { methods, grand } = useMemo(() => sumDaysTotals(selectedDays), [selectedDays]);

  return (
    <section className="bg-surface-container border-2 border-secondary p-md space-y-md sticky top-2 z-10 shadow-lg">
      <div className="flex items-start justify-between flex-wrap gap-sm border-b border-outline-variant pb-base">
        <div>
          <h3 className="text-h2 font-h2 text-on-surface">Modificar Cuadre</h3>
          <div className="flex items-center gap-sm mt-xs flex-wrap">
            <p className="text-label-sm text-on-surface-variant uppercase tracking-wider">
              {selectedDays.length === 0
                ? 'Selecciona los días que quieres incluir'
                : `${selectedDays.length} de ${totalDays} ${
                    totalDays === 1 ? 'día seleccionado' : 'días seleccionados'
                  }`}
            </p>
            <button
              onClick={onToggleAll}
              disabled={downloading || totalDays === 0}
              className="text-label-sm text-secondary hover:underline font-bold uppercase tracking-wider disabled:opacity-50"
            >
              {allSelected ? 'Quitar Todos' : 'Seleccionar Todos'}
            </button>
          </div>
        </div>
        <div className="flex gap-sm">
          <button
            onClick={onCancel}
            disabled={downloading}
            className="border border-outline-variant text-on-surface px-md py-xs font-label-sm font-bold hover:bg-surface-container-high transition-colors disabled:opacity-50"
          >
            Cancelar
          </button>
          <button
            onClick={onPrint}
            disabled={selectedDays.length === 0 || downloading}
            className="bg-secondary text-on-secondary px-md py-xs font-label-sm font-bold disabled:opacity-50 flex items-center gap-xs hover:opacity-90 transition-opacity"
          >
            <span className="material-symbols-outlined text-[18px]">picture_as_pdf</span>
            {downloading ? 'Descargando...' : 'Descargar Resumen'}
          </button>
        </div>
      </div>

      <ul className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-sm">
        <li className="bg-secondary text-on-secondary px-sm py-xs flex flex-col gap-xs col-span-2 sm:col-span-4 lg:col-span-1">
          <span className="text-label-sm uppercase tracking-wider font-bold">Total</span>
          <span className="text-data-mono font-data-mono text-[16px] font-bold">
            {formatMoney(grand)}
          </span>
        </li>
        {methods.map((m) => (
          <li
            key={m.method}
            className="bg-surface-container-lowest border border-outline-variant px-sm py-xs flex flex-col gap-xs"
          >
            <span className="text-label-sm text-on-surface-variant uppercase tracking-wider">
              {m.method}
            </span>
            <span className="text-data-mono font-data-mono text-on-surface text-[16px]">
              {formatMoney(m.amount)}
            </span>
          </li>
        ))}
      </ul>
    </section>
  );
}

function IssuesBanner({ issues }: { issues: TransactionIssue[] }) {
  return (
    <section className="border border-error bg-error-container/20 p-md space-y-sm">
      <div className="flex items-center gap-sm">
        <span className="material-symbols-outlined text-error text-[24px]">error</span>
        <h2 className="text-h2 font-h2 text-error">
          {issues.length} {issues.length === 1 ? 'transacción' : 'transacciones'} con problemas
        </h2>
      </div>
      <p className="text-label-sm text-on-surface-variant">
        No se incluirán en los PDFs. Revisa la hoja "Transacciones Organizadas" del Excel.
      </p>
      <ul className="space-y-xs max-h-64 overflow-y-auto pr-sm">
        {issues.map((issue, i) => (
          <li
            key={i}
            className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-xs bg-surface-container-lowest border border-outline-variant px-sm py-xs"
          >
            <div className="flex flex-col text-data-mono font-data-mono">
              <span className="text-on-surface">
                {issue.transaction.rawDate || '(sin fecha)'} ·{' '}
                {issue.transaction.servicio || '(sin servicio)'}
              </span>
              <span className="text-on-surface-variant">
                {issue.transaction.metodo || '(sin método)'} ·{' '}
                {formatMoney(issue.transaction.total)}
              </span>
            </div>
            <span className="text-label-sm text-error sm:text-right uppercase tracking-wide">
              {issue.reason}
            </span>
          </li>
        ))}
      </ul>
    </section>
  );
}

function GrandTotalCard({
  rows,
  onDownload,
}: {
  rows: ParsedWorkbook['grandTotals'];
  onDownload: () => void;
}) {
  return (
    <section className="bg-surface-container border border-outline-variant p-md">
      <div className="flex items-start justify-between border-b border-outline-variant pb-base mb-md flex-wrap gap-sm">
        <h2 className="text-h2 font-h2 text-on-surface">Totales Generales</h2>
        <button
          onClick={onDownload}
          className="bg-secondary text-on-secondary px-md py-xs font-label-sm font-bold active:scale-95 transition-all flex items-center gap-xs"
        >
          <span className="material-symbols-outlined text-[18px]">download</span>
          Descargar PDF
        </button>
      </div>

      <ul className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-sm">
        {rows.map((r) => (
          <li
            key={r.label}
            className="bg-surface-container-lowest border border-outline-variant p-sm flex flex-col gap-xs"
          >
            <span className="text-label-sm text-on-surface-variant uppercase tracking-wider">
              {r.label}
            </span>
            <span className="text-data-mono font-data-mono text-secondary text-[18px]">
              {r.amount}
            </span>
          </li>
        ))}
      </ul>
    </section>
  );
}

function DayCard({
  day,
  transactions,
  selectionMode = false,
  selected = false,
  onToggleSelect,
}: {
  day: DayTotals;
  transactions: Transaction[];
  selectionMode?: boolean;
  selected?: boolean;
  onToggleSelect?: () => void;
}) {
  const [open, setOpen] = useState(false);

  function downloadDay() {
    const doc = buildDayPdf(day, transactions);
    doc.save(dayFilename(day));
  }

  const selectionClasses = selectionMode
    ? `cursor-pointer ${
        selected
          ? 'border-secondary border-2 bg-secondary/5'
          : 'border-outline-variant hover:border-on-surface-variant'
      }`
    : 'border border-outline-variant hover:bg-surface-container-high';

  return (
    <article
      className={`bg-surface-container p-md transition-colors group relative flex flex-col ${selectionClasses}`}
      onClick={selectionMode ? onToggleSelect : undefined}
      role={selectionMode ? 'button' : undefined}
      aria-pressed={selectionMode ? selected : undefined}
    >
      <div className="flex items-start justify-between mb-md">
        <div className="p-sm bg-surface-container-lowest border border-outline-variant rounded">
          <span className="material-symbols-outlined text-on-surface">description</span>
        </div>
        {selectionMode ? (
          <div
            className={`w-7 h-7 border-2 flex items-center justify-center transition-colors ${
              selected ? 'bg-secondary border-secondary' : 'border-outline-variant bg-surface-container-lowest'
            }`}
            aria-hidden
          >
            {selected && (
              <span className="material-symbols-outlined text-on-secondary text-[20px]">
                check
              </span>
            )}
          </div>
        ) : (
          <button
            onClick={downloadDay}
            aria-label="Descargar PDF del día"
            className="text-secondary hover:opacity-80 transition-opacity p-xs"
          >
            <span className="material-symbols-outlined text-[28px]">download</span>
          </button>
        )}
      </div>

      <div>
        <p className="text-body-md font-semibold text-on-surface truncate">
          Cuadre {day.rawDate}
        </p>
        <p className="text-label-sm text-on-surface-variant">
          Total diario · {formatMoney(day.totalDiario)}
        </p>
      </div>

      <div className="grid grid-cols-2 gap-xs mt-md">
        {day.totals.map((t) => (
          <div
            key={t.method}
            className="bg-surface-container-lowest border border-outline-variant px-sm py-xs flex flex-col gap-xs"
          >
            <span className="text-label-sm text-on-surface-variant uppercase tracking-wider">
              {t.method}
            </span>
            <span className="text-data-mono font-data-mono text-on-surface">
              {formatMoney(t.amount)}
            </span>
          </div>
        ))}
      </div>

      {!selectionMode && (
        <>
          <div className="mt-md pt-sm border-t border-outline-variant flex justify-between items-center">
            <span className="text-label-sm font-bold text-emerald-500 uppercase">Listo</span>
            <button
              onClick={() => setOpen((v) => !v)}
              className="text-label-sm text-secondary hover:underline flex items-center gap-xs font-bold"
            >
              {open ? 'Ocultar' : `Ver ${transactions.length}`}
              <span className="material-symbols-outlined text-[14px]">
                {open ? 'expand_less' : 'expand_more'}
              </span>
            </button>
          </div>

          {open && (
            <div className="mt-sm border-t border-outline-variant pt-sm">
              {transactions.length === 0 ? (
                <p className="text-label-sm text-on-surface-variant text-center py-md">
                  Sin transacciones para este día.
                </p>
              ) : (
                <table className="w-full text-data-mono font-data-mono">
                  <thead>
                    <tr className="text-left text-label-sm text-on-surface-variant uppercase tracking-wider">
                      <th className="py-xs pr-xs font-bold">Servicio</th>
                      <th className="py-xs px-xs font-bold">Método</th>
                      <th className="py-xs pl-xs font-bold text-right">Monto</th>
                    </tr>
                  </thead>
                  <tbody>
                    {transactions.map((t, i) => (
                      <tr key={i} className="border-t border-outline-variant">
                        <td className="py-xs pr-xs text-on-surface align-top">{t.servicio}</td>
                        <td className="py-xs px-xs text-on-surface-variant align-top">{t.metodo}</td>
                        <td className="py-xs pl-xs text-right text-on-surface align-top">
                          {formatMoney(t.total)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}
        </>
      )}
    </article>
  );
}
