export type DayTotals = {
  rawDate: string;
  dateKey: string;
  totals: { method: PaymentMethod; amount: number }[];
  totalDiario: number;
};

export type Transaction = {
  rawDate: string;
  dateKey: string;
  servicio: string;
  total: number;
  metodo: string;
  diaPuesto: string;
};

export type GrandTotalRow = {
  label: string;
  amount: string;
};

export type TransactionIssue = {
  transaction: Transaction;
  reason: string;
};

export type ParsedWorkbook = {
  days: DayTotals[];
  transactionsByDay: Record<string, Transaction[]>;
  grandTotals: GrandTotalRow[];
  issues: TransactionIssue[];
};

export type PaymentMethod = 'Zelle' | 'Cash' | 'Clover' | 'Venmo' | 'Paypal' | 'Cash App';
