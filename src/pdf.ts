import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import JSZip from 'jszip';
import { applyDeducibles, formatMoney, sumDaysTotals } from './parser';
import type { DayTotals, Deducible, GrandTotalRow, Transaction } from './types';

// Black & white only — no fills, just thin black borders. Saves ink.
const BW_TABLE_DEFAULTS = {
  theme: 'grid' as const,
  headStyles: {
    fillColor: false as const,
    textColor: 0,
    fontStyle: 'bold' as const,
    lineColor: 0,
    lineWidth: 0.5,
  },
  bodyStyles: {
    fillColor: false as const,
    textColor: 0,
    lineColor: 0,
    lineWidth: 0.25,
  },
  styles: {
    lineColor: 0,
    lineWidth: 0.25,
    textColor: 0,
  },
};

function newDoc() {
  return new jsPDF({ unit: 'pt', format: 'letter' });
}

function header(doc: jsPDF, title: string, subtitle?: string) {
  doc.setTextColor(0);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(20);
  doc.text(title, 40, 50);
  if (subtitle) {
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(12);
    doc.text(subtitle, 40, 70);
  }
}

export function buildDayPdf(day: DayTotals, transactions: Transaction[]): jsPDF {
  const doc = newDoc();
  header(doc, `Cuadre del ${day.rawDate}`, `Total diario: ${formatMoney(day.totalDiario)}`);

  autoTable(doc, {
    ...BW_TABLE_DEFAULTS,
    startY: 90,
    head: [['Método', 'Monto']],
    body: day.totals.map((t) => [t.method, formatMoney(t.amount)]),
    styles: { ...BW_TABLE_DEFAULTS.styles, fontSize: 11, cellPadding: 6 },
    columnStyles: { 1: { halign: 'right' } },
    margin: { left: 40, right: 40 },
    tableWidth: 280,
  });

  const afterTotalsY = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY;

  autoTable(doc, {
    ...BW_TABLE_DEFAULTS,
    startY: afterTotalsY + 24,
    head: [['Servicio', 'Método', 'Monto']],
    body: transactions.length
      ? transactions.map((t) => [t.servicio, t.metodo, formatMoney(t.total)])
      : [[{ content: 'Sin transacciones para este día.', colSpan: 3, styles: { halign: 'center' } }]],
    styles: { ...BW_TABLE_DEFAULTS.styles, fontSize: 10, cellPadding: 5 },
    columnStyles: { 2: { halign: 'right', cellWidth: 80 } },
    margin: { left: 40, right: 40 },
  });

  return doc;
}

export function buildGrandTotalPdf(
  grandTotals: GrandTotalRow[],
  period?: { first: string; last: string },
): jsPDF {
  const doc = newDoc();
  const subtitle = period ? `Periodo: ${period.first} a ${period.last}` : undefined;
  header(doc, 'Totales Generales', subtitle);

  autoTable(doc, {
    ...BW_TABLE_DEFAULTS,
    startY: 90,
    head: [['Categoría', 'Monto']],
    body: grandTotals.map((g) => [g.label, g.amount]),
    styles: { ...BW_TABLE_DEFAULTS.styles, fontSize: 12, cellPadding: 8 },
    columnStyles: { 1: { halign: 'right' } },
    margin: { left: 40, right: 40 },
    tableWidth: 360,
  });

  return doc;
}

export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function dayFilename(day: DayTotals): string {
  return `Cuadre ${day.dateKey || day.rawDate.replace(/\//g, '-')}.pdf`;
}

function shortMonthDay(rawDate: string): string {
  // "4/27/2026" -> "4-27"; falls back to whatever's there.
  const parts = rawDate.split('/');
  if (parts.length < 2) return rawDate.replace(/[\\/:*?"<>|]/g, '-');
  return `${parts[0]}-${parts[1]}`;
}

export function rangeFolderName(days: DayTotals[]): string {
  if (days.length === 0) return 'cuadre';
  const sorted = [...days].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  const first = shortMonthDay(sorted[0].rawDate);
  const last = shortMonthDay(sorted[sorted.length - 1].rawDate);
  return first === last ? `cuadre-${first}` : `cuadre-${first}-a-${last}`;
}

export function buildCustomSummaryPdf(
  selectedDays: DayTotals[],
  deducibles: Deducible[] = [],
): jsPDF {
  const sorted = [...selectedDays].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  const { methods, grand } = sumDaysTotals(sorted);
  const after = applyDeducibles(methods, grand, deducibles);

  const subtitle =
    sorted.length === 0
      ? 'Sin días seleccionados'
      : sorted.length === 1
        ? `Día: ${sorted[0].rawDate}`
        : `Periodo: ${sorted[0].rawDate} a ${sorted[sorted.length - 1].rawDate} · ${sorted.length} días`;

  const doc = newDoc();
  header(doc, 'Resumen del Cuadre', subtitle);

  const boldFootStyles = {
    fillColor: false as const,
    textColor: 0,
    fontStyle: 'bold' as const,
    lineColor: 0,
    lineWidth: 0.5,
  };

  autoTable(doc, {
    ...BW_TABLE_DEFAULTS,
    startY: 90,
    head: [['Método', 'Monto']],
    body: methods.map((m) => [m.method, formatMoney(m.amount)]),
    foot: [['Total General', formatMoney(grand)]],
    footStyles: boldFootStyles,
    columnStyles: { 1: { halign: 'right' } },
    margin: { left: 40, right: 40 },
    tableWidth: 280,
    styles: { ...BW_TABLE_DEFAULTS.styles, fontSize: 11, cellPadding: 6 },
  });

  const lastY = () =>
    (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY;

  if (deducibles.length > 0) {
    autoTable(doc, {
      ...BW_TABLE_DEFAULTS,
      startY: lastY() + 24,
      head: [['Deducible', 'Método', 'Monto']],
      body: deducibles.map((d) => [d.name || '(sin nombre)', d.method, formatMoney(d.amount)]),
      columnStyles: { 2: { halign: 'right', cellWidth: 80 } },
      margin: { left: 40, right: 40 },
      styles: { ...BW_TABLE_DEFAULTS.styles, fontSize: 10, cellPadding: 5 },
    });

    autoTable(doc, {
      ...BW_TABLE_DEFAULTS,
      startY: lastY() + 24,
      head: [['Total Después de Deducibles', 'Monto']],
      body: after.methods.map((m) => [m.method, formatMoney(m.amount)]),
      foot: [['Total General', formatMoney(after.grand)]],
      footStyles: boldFootStyles,
      columnStyles: { 1: { halign: 'right' } },
      margin: { left: 40, right: 40 },
      tableWidth: 280,
      styles: { ...BW_TABLE_DEFAULTS.styles, fontSize: 11, cellPadding: 6 },
    });
  }

  if (sorted.length > 0) {
    autoTable(doc, {
      ...BW_TABLE_DEFAULTS,
      startY: lastY() + 24,
      head: [['Día', 'Total Diario']],
      body: sorted.map((d) => [d.rawDate, formatMoney(d.totalDiario)]),
      columnStyles: { 1: { halign: 'right' } },
      margin: { left: 40, right: 40 },
      styles: { ...BW_TABLE_DEFAULTS.styles, fontSize: 10, cellPadding: 5 },
    });
  }

  return doc;
}

function summaryFilename(sortedDays: DayTotals[]): string {
  if (sortedDays.length === 0) return 'Resumen Cuadre.pdf';
  const first = shortMonthDay(sortedDays[0].rawDate);
  const last = shortMonthDay(sortedDays[sortedDays.length - 1].rawDate);
  return first === last
    ? `Resumen Cuadre ${first}.pdf`
    : `Resumen Cuadre ${first} a ${last}.pdf`;
}

// Downloads just the summary PDF for the user-selected days (no zip, no per-day PDFs).
export function downloadSelectionSummary(
  selectedDays: DayTotals[],
  deducibles: Deducible[] = [],
): { filename: string } {
  const sorted = [...selectedDays].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  const filename = summaryFilename(sorted);
  buildCustomSummaryPdf(sorted, deducibles).save(filename);
  return { filename };
}

// Bundles the user-selected days into a zip with a custom summary PDF + each
// selected day's full PDF. Folder name uses the same range pattern.
export async function downloadSelectionZip(
  selectedDays: DayTotals[],
  transactionsByDay: Record<string, Transaction[]>,
  deducibles: Deducible[] = [],
): Promise<{ folderName: string; count: number }> {
  const sorted = [...selectedDays].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  const folderName = rangeFolderName(sorted);

  const zip = new JSZip();
  const folder = zip.folder(folderName);
  if (!folder) throw new Error('No se pudo crear la carpeta dentro del zip.');

  folder.file(summaryFilename(sorted), buildCustomSummaryPdf(sorted, deducibles).output('blob'));
  for (const day of sorted) {
    const txs = transactionsByDay[day.dateKey] ?? [];
    folder.file(dayFilename(day), buildDayPdf(day, txs).output('blob'));
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  downloadBlob(zipBlob, `${folderName}.zip`);

  return { folderName, count: sorted.length + 1 };
}

// Bundles every PDF into a single zip and downloads it to the user's Downloads
// folder. The zip contains a folder so extracting yields cuadre-{first}-a-{last}/
// with all PDFs inside.
export async function downloadAllPdfs(
  days: DayTotals[],
  transactionsByDay: Record<string, Transaction[]>,
  grandTotals: GrandTotalRow[],
): Promise<{ folderName: string; count: number }> {
  const sorted = [...days].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  const folderName = rangeFolderName(sorted);
  const period =
    sorted.length > 0
      ? { first: sorted[0].rawDate, last: sorted[sorted.length - 1].rawDate }
      : undefined;

  const zip = new JSZip();
  const folder = zip.folder(folderName);
  if (!folder) throw new Error('No se pudo crear la carpeta dentro del zip.');

  folder.file(summaryFilename(sorted), buildGrandTotalPdf(grandTotals, period).output('blob'));
  for (const day of sorted) {
    const txs = transactionsByDay[day.dateKey] ?? [];
    folder.file(dayFilename(day), buildDayPdf(day, txs).output('blob'));
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  downloadBlob(zipBlob, `${folderName}.zip`);

  return { folderName, count: sorted.length + 1 };
}
