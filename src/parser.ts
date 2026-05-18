import * as XLSX from 'xlsx';
import type {
  DayTotals,
  Deducible,
  GrandTotalRow,
  ParsedWorkbook,
  PaymentMethod,
  Transaction,
  TransactionIssue,
} from './types';

const TOTALES_SHEET = 'Totales';
const TRANSACCIONES_SHEET = 'Transacciones Organizadas';

export const PAYMENT_METHODS: PaymentMethod[] = [
  'Zelle',
  'Cash',
  'Clover',
  'Venmo',
  'Paypal',
  'Cash App',
];

const METHOD_COLUMNS = PAYMENT_METHODS;

export function parseWorkbook(buf: ArrayBuffer): ParsedWorkbook {
  const wb = XLSX.read(buf, { cellDates: true });

  const totales = wb.Sheets[TOTALES_SHEET];
  const transacciones = wb.Sheets[TRANSACCIONES_SHEET];
  if (!totales) throw new Error(`Falta la hoja "${TOTALES_SHEET}" en el Excel.`);
  if (!transacciones) throw new Error(`Falta la hoja "${TRANSACCIONES_SHEET}" en el Excel.`);

  const totalesRows = XLSX.utils.sheet_to_json<unknown[]>(totales, {
    header: 1,
    raw: false,
    defval: '',
  });
  const transRows = XLSX.utils.sheet_to_json<unknown[]>(transacciones, {
    header: 1,
    raw: false,
    defval: '',
  });

  const grandTotals = parseGrandTotals(totalesRows);
  const days = parseDays(totalesRows);
  const transactions = parseTransactions(transRows);

  const dayKeys = new Set(days.map((d) => d.dateKey));
  const transactionsByDay: Record<string, Transaction[]> = {};
  const issues: TransactionIssue[] = [];

  for (const t of transactions) {
    if (!t.rawDate) {
      issues.push({ transaction: t, reason: 'Sin fecha' });
      continue;
    }
    if (!dayKeys.has(t.dateKey)) {
      issues.push({
        transaction: t,
        reason: `La fecha "${t.rawDate}" no aparece en la hoja Totales`,
      });
      continue;
    }
    (transactionsByDay[t.dateKey] ??= []).push(t);
  }

  return { days, transactionsByDay, grandTotals, issues };
}

function parseGrandTotals(rows: unknown[][]): GrandTotalRow[] {
  const result: GrandTotalRow[] = [];
  for (let i = 1; i < rows.length; i++) {
    const cell = String(rows[i]?.[0] ?? '').trim();
    if (!cell) continue;
    const colonIdx = cell.indexOf(':');
    if (colonIdx === -1) {
      result.push({ label: cell, amount: '' });
      continue;
    }
    const label = cell.slice(0, colonIdx).trim();
    const amount = cell.slice(colonIdx + 1).trim();
    result.push({ label, amount });
  }
  return result;
}

function parseDays(rows: unknown[][]): DayTotals[] {
  const days: DayTotals[] = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] ?? [];
    const rawDate = rawDateString(row[2]);
    if (!rawDate) continue;

    const totals = METHOD_COLUMNS.map((method, idx) => ({
      method,
      amount: extractAmount(row[3 + idx]),
    }));
    const totalDiario = extractAmount(row[9]);

    days.push({
      rawDate,
      dateKey: normalizeDateKey(rawDate),
      totals,
      totalDiario,
    });
  }
  return days;
}

function parseTransactions(rows: unknown[][]): Transaction[] {
  const list: Transaction[] = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] ?? [];
    const rawDate = rawDateString(row[0]);
    const servicio = String(row[1] ?? '').trim();
    const totalRaw = row[2];
    const metodo = normalizeMethod(String(row[3] ?? ''));
    const diaPuesto = String(row[4] ?? '').trim();

    if (!rawDate && !servicio) continue;

    list.push({
      rawDate,
      dateKey: normalizeDateKey(rawDate),
      servicio,
      total: extractAmount(totalRaw),
      metodo,
      diaPuesto,
    });
  }
  return list;
}

function extractAmount(value: unknown): number {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  // Strip thousands separators (commas) so "1,234.56" parses correctly.
  const cleaned = String(value).replace(/,/g, '');
  const match = cleaned.match(/-?\d+(?:\.\d+)?/);
  return match ? parseFloat(match[0]) : 0;
}

function rawDateString(value: unknown): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    return `${value.getMonth() + 1}/${value.getDate()}/${value.getFullYear()}`;
  }
  return String(value).trim();
}

const METHOD_ALIASES: Record<string, string> = {
  zelle: 'Zelle',
  cash: 'Cash',
  efectivo: 'Cash',
  clover: 'Clover',
  venmo: 'Venmo',
  paypal: 'Paypal',
  'pay pal': 'Paypal',
  'cash app': 'Cash App',
  cashapp: 'Cash App',
  'cash-app': 'Cash App',
};

function normalizeMethod(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) return '';
  return METHOD_ALIASES[trimmed.toLowerCase()] ?? trimmed;
}

export function normalizeDateKey(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) return '';
  // Accept /, -, or . as separators. Drop the time portion if present.
  const dateOnly = trimmed.split(/\s+/)[0];
  const parts = dateOnly.split(/[\/\-.]/).filter(Boolean);
  if (parts.length !== 3) return trimmed;

  let y: number;
  let m: number;
  let d: number;
  if (parts[0].length === 4) {
    // ISO: YYYY-MM-DD
    y = parseInt(parts[0], 10);
    m = parseInt(parts[1], 10);
    d = parseInt(parts[2], 10);
  } else {
    // US: M/D/YY or M/D/YYYY
    m = parseInt(parts[0], 10);
    d = parseInt(parts[1], 10);
    y = parseInt(parts[2], 10);
  }

  if (Number.isNaN(m) || Number.isNaN(d) || Number.isNaN(y)) return trimmed;
  if (y < 100) y += 2000;
  return `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

export function formatMoney(n: number): string {
  return `$${n.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

// Aggregates the per-day totals across the given days into per-method amounts
// and an overall grand total. Source values come straight from the Excel
// (each day.totals[i].amount and day.totalDiario), so this just re-totals
// numbers the spreadsheet itself produced.
export function sumDaysTotals(days: DayTotals[]): {
  methods: { method: PaymentMethod; amount: number }[];
  grand: number;
} {
  const byMethod: Record<string, number> = {};
  let grand = 0;
  for (const day of days) {
    for (const t of day.totals) {
      byMethod[t.method] = (byMethod[t.method] ?? 0) + t.amount;
    }
    grand += day.totalDiario;
  }
  const methods = METHOD_COLUMNS.map((method) => ({
    method,
    amount: byMethod[method] ?? 0,
  }));
  return { methods, grand };
}

// Subtracts each deducible from its matching method's amount and from the
// grand total. Amounts can go negative if a deducible exceeds the method
// total — that's intentional, the user trusts their own input.
export function applyDeducibles(
  methods: { method: PaymentMethod; amount: number }[],
  grand: number,
  deducibles: Deducible[],
): { methods: { method: PaymentMethod; amount: number }[]; grand: number } {
  const byMethod: Record<string, number> = {};
  let totalDeduc = 0;
  for (const d of deducibles) {
    const amt = Number.isFinite(d.amount) ? d.amount : 0;
    byMethod[d.method] = (byMethod[d.method] ?? 0) + amt;
    totalDeduc += amt;
  }
  return {
    methods: methods.map((m) => ({
      method: m.method,
      amount: m.amount - (byMethod[m.method] ?? 0),
    })),
    grand: grand - totalDeduc,
  };
}
