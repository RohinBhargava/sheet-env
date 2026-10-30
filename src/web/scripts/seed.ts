// Builds the "fund-admin" scenario: fictional fund-administration workbooks for Halverson Growth Partners III,
// L.P., written with the xlsx engine (formulas carry cached values), plus tasks.json.
//
//   pnpm --dir src/web seed
//
// Output: src/workbooks/seed/fund-admin/files/<folder>/<name>.xlsx and src/workbooks/seed/fund-admin/tasks.json
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { Book, FMT, serial } from './xl.ts';
import { ADMIN, CALL4, CALL4_TOTAL, FEE, FEE_DAYS, FUND, GP, LPS, M, PRIOR_CALLS, TOTAL_M, feeFor, rollForward } from './fund-data.ts';

const ROOT = resolve(import.meta.dirname, '../../workbooks/seed/fund-admin');
const FILES = join(ROOT, 'files');

const GREEN = '#1F6F43';
const HEAD_FILL = '#E3F2E9';
const INPUT_FILL = '#FFF4CE';

function title(b: Book, sheet: string, line2: string, line3?: string) {
  b.set(sheet, 'A1', FUND);
  b.fmt(sheet, 'A1', { font: { b: true, sz: 14, color: GREEN } });
  b.set(sheet, 'A2', line2);
  b.fmt(sheet, 'A2', { font: { b: true, sz: 12 } });
  if (line3) { b.set(sheet, 'A3', line3); b.fmt(sheet, 'A3', { font: { i: true, color: '#595959' } }); }
}

function header(b: Book, sheet: string, range: string, wrap = false) {
  b.fmt(sheet, range, { font: { b: true }, fill: HEAD_FILL, border: 'bottom', align: { wrap, v: 'center' } });
}

function totalRow(b: Book, sheet: string, range: string) {
  b.fmt(sheet, range, { font: { b: true }, border: 'top' });
}

const round2 = (x: number) => Math.round(x * 100) / 100;
const written: string[] = [];
function write(rel: string, b: Book) {
  const out = join(FILES, rel);
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, b.save());
  written.push(rel);
}

// ---------------------------------------------------------------- 1. LP commitments register
function commitmentsRegister() {
  const b = new Book();
  const S = 'Commitments', T = 'Fund Terms';
  b.sheet(S); b.sheet(T);
  title(b, S, 'LP Commitments Register', 'As of September 30, 2026 (USD)');
  b.row(S, 'A5', ['LP ID', 'Investor', 'Investor Type', 'Commitment', 'Commitment %', 'Called to Date', 'Unfunded Commitment', '% Called']);
  header(b, S, 'A5:H5');
  LPS.forEach((lp, i) => {
    const r = 6 + i;
    b.row(S, `A${r}`, [lp.id, lp.name, lp.type, lp.commitmentM * M, `=D${r}/$D$18`, `=ROUND(D${r}*'Fund Terms'!$C$17,2)`, `=D${r}-F${r}`, `=IF(D${r}=0,0,F${r}/D${r})`]);
  });
  b.row(S, 'A18', ['Total', null, null, '=SUM(D6:D17)', '=SUM(E6:E17)', '=SUM(F6:F17)', '=SUM(G6:G17)', '=IF(D18=0,0,F18/D18)']);
  totalRow(b, S, 'A18:H18');
  b.fmt(S, 'D6:D18', { numFmt: FMT.usd }); b.fmt(S, 'F6:G18', { numFmt: FMT.usd });
  b.fmt(S, 'E6:E18', { numFmt: FMT.pct4 }); b.fmt(S, 'H6:H18', { numFmt: FMT.pct });
  b.widths(S, { A: 9, B: 46, C: 18, D: 16, E: 13, F: 16, G: 19, H: 10 });
  b.freeze(S, 5, 2);

  title(b, T, 'Fund Terms');
  const terms: [string, string | number, string?][] = [
    ['Fund', FUND], ['General partner', GP], ['Administrator', ADMIN], ['Vintage year', 2024],
    ['Final close', serial(2025, 3, 31), FMT.date], ['Investment period ends', serial(2030, 3, 31), FMT.date],
    ['Management fee (annual, on commitments during the investment period)', FEE.rate, FMT.pct], ['Fee day count', 'Actual/365'],
    ['Fee-paying partners', 'All limited partners (the general partner pays no fee)'],
    ['Total commitments', '=Commitments!D18', FMT.usd],
  ];
  terms.forEach(([k, v, f], i) => { b.set(T, `A${4 + i}`, k); b.set(T, `C${4 + i}`, v); if (f) b.fmt(T, `C${4 + i}`, { numFmt: f }); });
  b.row(T, 'A15', ['Capital calls to date', null, '% of commitments', 'Notice date']);
  header(b, T, 'A15:D15');
  // calls 1-3 total in C17 (Commitments!F uses it)
  b.row(T, 'A16', ['Calls 1-3 (see below)']);
  b.set(T, 'C17', '=SUM(C19:C21)'); b.set(T, 'A17', 'Cumulative called'); b.fmt(T, 'A17:C17', { font: { b: true } });
  const dates = [serial(2025, 6, 16), serial(2026, 1, 20), serial(2026, 7, 15)];
  PRIOR_CALLS.forEach((p, i) => { b.row(T, `A${19 + i}`, [`Capital Call No. ${i + 1}`, null, p, dates[i]]); });
  b.fmt(T, 'C17:C21', { numFmt: FMT.pct }); b.fmt(T, 'D19:D21', { numFmt: FMT.date });
  b.widths(T, { A: 62, B: 2, C: 26, D: 14 });
  write('Fund Records/LP Commitments Register.xlsx', b);
}

// ---------------------------------------------------------------- 2. Capital call no. 4
function capitalCall4() {
  const b = new Book();
  const S = 'Call Summary', A = 'Allocation';
  b.sheet(S); b.sheet(A);
  title(b, S, 'Capital Call Notice No. 4');
  b.row(S, 'A4', ['Notice date', null, serial(...CALL4.notice)]);
  b.row(S, 'A5', ['Due date', null, serial(...CALL4.due)]);
  b.row(S, 'A6', ['Wire to', null, `${FUND} — Operating Account (x4417)`]);
  b.fmt(S, 'C4:C5', { numFmt: FMT.date, align: { h: 'left' } });
  b.row(S, 'A8', ['Use of proceeds', null, 'Amount']);
  header(b, S, 'A8:C8');
  CALL4.items.forEach((it, i) => b.row(S, `A${9 + i}`, [it.label, null, it.amount]));
  b.row(S, 'A12', ['Total capital call', null, '=SUM(C9:C11)']);
  totalRow(b, S, 'A12:C12');
  b.row(S, 'A14', ['Total commitments', null, '=Allocation!C18']);
  b.row(S, 'A15', ['Call as % of commitments', null, '=IF(C14=0,0,C12/C14)']);
  b.fmt(S, 'C9:C14', { numFmt: FMT.usd }); b.fmt(S, 'C15', { numFmt: FMT.pct });
  b.widths(S, { A: 54, B: 2, C: 22 });

  title(b, A, 'Capital Call No. 4 — allocation by commitment', 'Each partner\'s call amount = Commitment % × Total capital call, rounded to the cent.');
  b.row(A, 'A5', ['LP ID', 'Investor', 'Commitment', 'Commitment %', 'Call Amount']);
  header(b, A, 'A5:E5');
  LPS.forEach((lp, i) => {
    const r = 6 + i;
    b.row(A, `A${r}`, [lp.id, lp.name, lp.commitmentM * M, `=C${r}/$C$18`]);
  });
  b.fmt(A, 'E6:E17', { fill: INPUT_FILL });
  b.row(A, 'A18', ['Total', null, '=SUM(C6:C17)', '=SUM(D6:D17)', '=SUM(E6:E17)']);
  totalRow(b, A, 'A18:E18');
  b.row(A, 'A20', ['Total capital call (Call Summary)', null, null, null, "='Call Summary'!C12"]);
  b.row(A, 'A21', ['Difference', null, null, null, '=ROUND(E18-E20,2)']);
  b.fmt(A, 'A21:E21', { font: { b: true } });
  b.fmt(A, 'C6:C18', { numFmt: FMT.usd }); b.fmt(A, 'E6:E21', { numFmt: FMT.usd }); b.fmt(A, 'D6:D18', { numFmt: FMT.pct4 });
  b.widths(A, { A: 9, B: 46, C: 16, D: 14, E: 18 });
  b.freeze(A, 5, 0);
  write('Capital Calls/Capital Call 4 Notice.xlsx', b);
}

// ---------------------------------------------------------------- 3. Bank reconciliation
const GL_ROWS: [number, string, string, number, number][] = [
  // day, ref, description, debit, credit
  [2, 'W-0912', 'LP contribution — Call 3 true-up (Quarry Hill Partners LLC)', 125_000, 0],
  [5, 'CHK 4471', `Fund administration fee — September (${ADMIN})`, 0, 18_750],
  [9, 'W-0915', 'Investment — Alder Robotics, Inc. Series B', 0, 3_000_000],
  [12, 'W-0918', 'Distribution received — Kestrel Software recapitalization', 1_520_000, 0], // bank: 1,250,000 (transposed in the GL)
  [15, 'CHK 4472', 'Audit fees — FY2025 audit (Whitcombe & Hale LLP)', 0, 42_600],
  [18, 'W-0921', `Management fee — Q3 2026 (${GP})`, 0, 1_250_000],
  [24, 'CHK 4473', 'Valuation services — Q2 2026 (Corran Advisory)', 0, 9_800],               // not yet cleared
  [29, 'DEP-0930', 'LP contribution — late funding (Osprey Ridge Charitable Trust)', 200_000, 0], // deposit in transit
];
const BANK_ROWS: [number, string, number, number][] = [
  // day, description, withdrawal, deposit
  [2, 'Incoming wire — QUARRY HILL PARTNERS', 0, 125_000],
  [6, 'Check 4471', 18_750, 0],
  [9, 'Outgoing wire — ALDER ROBOTICS INC', 3_000_000, 0],
  [12, 'Incoming wire — KESTREL SOFTWARE HOLDINGS', 0, 1_250_000],
  [17, 'Check 4472', 42_600, 0],
  [18, 'Outgoing wire — HALVERSON GP III LLC', 1_250_000, 0],
  [30, 'Account maintenance fee', 85, 0],
  [30, 'Interest credit', 0, 3_412.57],
];
const OPENING = 4_812_350;

function bankRec() {
  const b = new Book();
  const R = 'Reconciliation', G = 'GL Cash', K = 'Bank Statement';
  b.sheet(R); b.sheet(G); b.sheet(K);

  title(b, G, 'General Ledger — Cash, Operating Account (1010)', 'September 2026 (USD)');
  b.row(G, 'A4', ['Date', 'Ref', 'Description', 'Debit', 'Credit', 'Balance']);
  header(b, G, 'A4:F4');
  b.row(G, 'A5', [serial(2026, 9, 1), null, 'Opening balance', null, null, OPENING]);
  GL_ROWS.forEach(([d, ref, desc, dr, cr], i) => {
    const r = 6 + i;
    b.row(G, `A${r}`, [serial(2026, 9, d), ref, desc, dr || null, cr || null, `=F${r - 1}+D${r}-E${r}`]);
  });
  b.row(G, 'A14', [serial(2026, 9, 30), null, 'Closing balance', '=SUM(D6:D13)', '=SUM(E6:E13)', '=F13']);
  totalRow(b, G, 'A14:F14');
  b.fmt(G, 'A5:A14', { numFmt: FMT.date }); b.fmt(G, 'D5:F14', { numFmt: FMT.usd });
  b.widths(G, { A: 11, B: 10, C: 60, D: 15, E: 15, F: 16 });
  b.freeze(G, 4, 0);

  title(b, K, 'Bank Statement — Operating Account (x4417)', 'Statement period 09/01/2026 – 09/30/2026 (USD)');
  b.row(K, 'A4', ['Date', 'Description', 'Withdrawals', 'Deposits', 'Balance']);
  header(b, K, 'A4:E4');
  b.row(K, 'A5', [serial(2026, 9, 1), 'Opening balance', null, null, OPENING]);
  BANK_ROWS.forEach(([d, desc, w, dep], i) => {
    const r = 6 + i;
    b.row(K, `A${r}`, [serial(2026, 9, d), desc, w || null, dep || null, `=E${r - 1}-C${r}+D${r}`]);
  });
  b.row(K, 'A14', [serial(2026, 9, 30), 'Closing balance', '=SUM(C6:C13)', '=SUM(D6:D13)', '=E13']);
  totalRow(b, K, 'A14:E14');
  b.fmt(K, 'A5:A14', { numFmt: FMT.date }); b.fmt(K, 'C5:E14', { numFmt: FMT.usd });
  b.widths(K, { A: 11, B: 44, C: 15, D: 15, E: 16 });
  b.freeze(K, 4, 0);

  title(b, R, 'Bank Reconciliation — Operating Account (x4417)', 'As of September 30, 2026 (USD)');
  b.row(R, 'A4', ['Balance per bank statement', null, "='Bank Statement'!E14"]);
  b.row(R, 'A5', ['Add: deposits in transit']);
  b.row(R, 'B6', ['DEP-0930 Osprey Ridge Charitable Trust (09/29)', 200_000]);
  b.row(R, 'A8', ['Less: outstanding checks']);
  b.row(R, 'A11', ['Adjusted bank balance', null, '=C4+SUM(C6:C7)-SUM(C9:C10)']);
  b.row(R, 'A13', ['Balance per general ledger', null, "='GL Cash'!F14"]);
  b.row(R, 'A14', ['Add: bank credits not recorded in the GL']);
  b.row(R, 'A17', ['Less: bank charges not recorded in the GL']);
  b.row(R, 'A20', ['Adjusted GL balance', null, '=C13+SUM(C15:C16)-SUM(C18:C19)']);
  b.row(R, 'A22', ['Difference', null, '=ROUND(C11-C20,2)']);
  b.row(R, 'A23', ['Status', null, '=IF(C22=0,"Reconciled","Not reconciled")']);
  for (const a of ['A4', 'A11', 'A13', 'A20', 'A22', 'A23']) b.fmt(R, `${a}:C${a.slice(1)}`, { font: { b: true } });
  b.fmt(R, 'A11:C11', { border: 'top' }); b.fmt(R, 'A20:C20', { border: 'top' });
  b.fmt(R, 'B6:C7', { fill: INPUT_FILL }); b.fmt(R, 'B9:C10', { fill: INPUT_FILL }); b.fmt(R, 'B15:C16', { fill: INPUT_FILL }); b.fmt(R, 'B18:C19', { fill: INPUT_FILL });
  b.fmt(R, 'C4:C22', { numFmt: FMT.usd }); b.fmt(R, 'C23', { align: { h: 'right' } });
  b.widths(R, { A: 40, B: 46, C: 18 });
  write('Accounting/Bank Reconciliation Sep 2026.xlsx', b);
}

// ---------------------------------------------------------------- 4. Q4 management fee
function managementFee() {
  const b = new Book();
  const S = 'Fee Calc';
  b.sheet(S);
  title(b, S, 'Management Fee Calculation — Q4 2026', '2.00% a year on fee-paying commitments during the investment period; Actual/365 day count.');
  const inputs: [string, string | number | null, string][] = [
    ['Annual fee rate', FEE.rate, FMT.pct],
    ['Quarter start', serial(...FEE.start), FMT.date],
    ['Quarter end', serial(...FEE.end), FMT.date],
    ['Day-count basis (days per year)', FEE.basis, FMT.int],
    ['Days in quarter', null, FMT.int],
    ['Fee-paying commitments', '=SUMIF(E13:E24,"Yes",D13:D24)', FMT.usd],
    ['Quarterly management fee', null, FMT.usd],
  ];
  inputs.forEach(([k, v, f], i) => { b.set(S, `A${4 + i}`, k); b.set(S, `C${4 + i}`, v); b.fmt(S, `C${4 + i}`, { numFmt: f, align: { h: 'right' } }); });
  b.fmt(S, 'C8', { fill: INPUT_FILL }); b.fmt(S, 'C10', { fill: INPUT_FILL }); b.fmt(S, 'A10:C10', { font: { b: true } });
  b.row(S, 'A12', ['LP ID', 'Investor', 'Investor Type', 'Commitment', 'Fee-paying?', 'Q4 2026 Fee']);
  header(b, S, 'A12:F12');
  LPS.forEach((lp, i) => {
    const r = 13 + i;
    b.row(S, `A${r}`, [lp.id, lp.name, lp.type, lp.commitmentM * M, lp.gp ? 'No' : 'Yes']);
  });
  b.fmt(S, 'F13:F24', { fill: INPUT_FILL });
  b.row(S, 'A25', ['Total', null, null, '=SUM(D13:D24)', null, '=SUM(F13:F24)']);
  totalRow(b, S, 'A25:F25');
  b.fmt(S, 'D13:D25', { numFmt: FMT.usd }); b.fmt(S, 'F13:F25', { numFmt: FMT.usd }); b.fmt(S, 'E13:E24', { align: { h: 'center' } });
  b.widths(S, { A: 34, B: 46, C: 18, D: 16, E: 12, F: 16 });
  write('Accounting/Q4 2026 Management Fee.xlsx', b);
}

// ---------------------------------------------------------------- 5. Capital account roll-forward
const LP005 = LPS.findIndex(l => l.id === 'LP-005');
const LP009 = LPS.findIndex(l => l.id === 'LP-009');

function capitalAccounts() {
  const b = new Book();
  const S = 'Roll-forward', F = 'Fund Summary';
  b.sheet(S); b.sheet(F);
  title(b, S, 'Partners\' Capital Account Roll-forward', 'Quarter ended September 30, 2026 (USD)');
  b.row(S, 'A5', ['LP ID', 'Investor', 'Beginning Balance 06/30/2026', 'Contributions', 'Distributions', 'Management Fees', 'Realized Gain/(Loss)', 'Unrealized Gain/(Loss)', 'Ending Balance 09/30/2026']);
  header(b, S, 'A5:I5', true);
  b.wb.setRowHeight(b.si(S), 4, 4, 30);
  LPS.forEach((lp, i) => {
    const r = 6 + i;
    const x = rollForward(lp);
    let end: string | number = `=C${r}+D${r}-E${r}-F${r}+G${r}+H${r}`;
    if (i === LP005) end = `=C${r}+D${r}-E${r}-F${r}+G${r}`;          // seeded error: unrealized omitted
    if (i === LP009) end = round2(x.b);                                // seeded error: stale hard-coded value
    b.row(S, `A${r}`, [lp.id, lp.name, round2(x.b), round2(x.c), round2(x.d), round2(x.f), round2(x.g), round2(x.h), end]);
  });
  b.row(S, 'A18', ['Total', null, ...['C', 'D', 'E', 'F', 'G', 'H', 'I'].map(c => `=SUM(${c}6:${c}17)`)]);
  totalRow(b, S, 'A18:I18');
  b.row(S, 'A20', ['Partners\' capital per Fund Summary', ...Array(7).fill(null), "='Fund Summary'!C11"]);
  b.row(S, 'A21', ['Difference', ...Array(7).fill(null), '=ROUND(I18-I20,2)']);
  b.fmt(S, 'A21:I21', { font: { b: true } });
  b.fmt(S, 'C6:I21', { numFmt: FMT.usd });
  b.widths(S, { A: 9, B: 46, C: 16, D: 15, E: 15, F: 15, G: 15, H: 15, I: 17 });
  b.freeze(S, 5, 2);

  const tot = LPS.map(rollForward).reduce((s, x) => ({ b: s.b + x.b, c: s.c + x.c, d: s.d + x.d, f: s.f + x.f, g: s.g + x.g, h: s.h + x.h }), { b: 0, c: 0, d: 0, f: 0, g: 0, h: 0 });
  title(b, F, 'Statement of Changes in Partners\' Capital', 'Quarter ended September 30, 2026 (USD), per the fund\'s general ledger');
  const lines: [string, number][] = [
    ['Partners\' capital, beginning of quarter', tot.b], ['Capital contributions', tot.c], ['Distributions', -tot.d],
    ['Management fees', -tot.f], ['Net realized gain on investments', tot.g], ['Net change in unrealized gain on investments', tot.h],
  ];
  lines.forEach(([k, v], i) => b.row(F, `A${4 + i}`, [k, null, round2(v)]));
  b.row(F, 'A11', ['Partners\' capital, end of quarter', null, '=SUM(C4:C9)']);
  totalRow(b, F, 'A11:C11');
  b.fmt(F, 'C4:C11', { numFmt: FMT.usd });
  b.widths(F, { A: 48, B: 2, C: 18 });
  write('Investor Reporting/Capital Account Statements Q3 2026.xlsx', b);
}

// ---------------------------------------------------------------- tasks.json
function tasks() {
  const share = (i: number) => round2(LPS[i].commitmentM / TOTAL_M * CALL4_TOTAL);
  const lpFees = LPS.map(l => round2(feeFor(l)));
  const fee = round2(LPS.reduce((s, l) => s + (l.gp ? 0 : l.commitmentM * M), 0) * FEE.rate * FEE_DAYS / FEE.basis);
  const rf = LPS.map(rollForward);
  const totalEnd = round2(rf.reduce((s, x) => s + x.end, 0));
  const bankClose = round2(OPENING + BANK_ROWS.reduce((s, [, , w, d]) => s + d - w, 0));
  const glFixed = round2(OPENING + GL_ROWS.reduce((s, [, ref, , dr, cr]) => s + (ref === 'W-0918' ? 1_250_000 : dr) - cr, 0));
  const adjusted = round2(glFixed + 3_412.57 - 85);
  if (Math.abs(round2(bankClose + 200_000 - 9_800) - adjusted) > 0.001) throw new Error('bank reconciliation data does not tie out');
  const brightwater = LPS.findIndex(l => l.id === 'LP-001');
  const quarry = LPS.findIndex(l => l.id === 'LP-010');
  const gp = LPS.findIndex(l => l.gp);
  return {
    scenario: 'fund-admin',
    fund: FUND,
    tasks: [
      {
        id: 'capital-call-4-allocation',
        title: 'Allocate Capital Call No. 4 to the partners',
        instructions: 'In Capital Calls/Capital Call 4 Notice.xlsx, fill in the Call Amount column (Allocation!E6:E17): each partner\'s share is its Commitment % multiplied by the Total capital call on the Call Summary sheet, rounded to the cent (for example =ROUND(D6*\'Call Summary\'!$C$12,2)). The Difference cell (Allocation!E21) must be 0. Save the workbook.',
        file: 'Capital Calls/Capital Call 4 Notice.xlsx',
        checks: [
          { sheet: 'Allocation', cell: `E${6 + brightwater}`, approx: share(brightwater), tolerance: 0.01 },
          { sheet: 'Allocation', cell: `E${6 + LP005}`, approx: share(LP005), tolerance: 0.01 },
          { sheet: 'Allocation', cell: `E${6 + gp}`, approx: share(gp), tolerance: 0.01 },
          { sheet: 'Allocation', cell: 'E18', approx: CALL4_TOTAL, tolerance: 0.01 },
          { sheet: 'Allocation', cell: 'E21', approx: 0, tolerance: 0.005 },
        ],
      },
      {
        id: 'bank-reconciliation-sep-2026',
        title: 'Reconcile the September 2026 operating account',
        instructions: 'In Accounting/Bank Reconciliation Sep 2026.xlsx, compare the GL Cash sheet with the Bank Statement sheet and clear the difference on the Reconciliation sheet: (1) correct the GL entry that was posted with the wrong amount (the bank statement is right); (2) list the check that has not cleared the bank under "Less: outstanding checks" (Reconciliation!B9:C9); (3) record the bank credit and the bank charge that are missing from the GL under the two "not recorded in the GL" sections (B15:C15 and B18:C18). Enter amounts as positive numbers. Reconciliation!C22 (Difference) must be 0. Save the workbook.',
        file: 'Accounting/Bank Reconciliation Sep 2026.xlsx',
        checks: [
          { sheet: 'GL Cash', cell: 'D9', equals: 1_250_000 },
          { sheet: 'GL Cash', cell: 'F14', approx: glFixed, tolerance: 0.005 },
          { sheet: 'Reconciliation', cell: 'C11', approx: adjusted, tolerance: 0.005 },
          { sheet: 'Reconciliation', cell: 'C20', approx: adjusted, tolerance: 0.005 },
          { sheet: 'Reconciliation', cell: 'C22', approx: 0, tolerance: 0.005 },
        ],
      },
      {
        id: 'management-fee-q4-2026',
        title: 'Calculate the Q4 2026 management fee',
        instructions: 'In Accounting/Q4 2026 Management Fee.xlsx (sheet Fee Calc): enter the number of days in the quarter in C8 (quarter start and end both count); compute the quarterly fee in C10 as fee-paying commitments × annual rate × days ÷ 365, rounded to the cent; and fill each partner\'s Q4 2026 Fee in F13:F24 the same way on its own commitment (0 for a partner that is not fee-paying). Save the workbook.',
        file: 'Accounting/Q4 2026 Management Fee.xlsx',
        checks: [
          { sheet: 'Fee Calc', cell: 'C8', equals: FEE_DAYS },
          { sheet: 'Fee Calc', cell: 'C10', approx: fee, tolerance: 0.01 },
          { sheet: 'Fee Calc', cell: `F${13 + brightwater}`, approx: lpFees[brightwater], tolerance: 0.01 },
          { sheet: 'Fee Calc', cell: `F${13 + quarry}`, approx: lpFees[quarry], tolerance: 0.01 },
          { sheet: 'Fee Calc', cell: 'F25', approx: round2(lpFees.reduce((s, x) => s + x, 0)), tolerance: 0.05 },
        ],
      },
      {
        id: 'capital-account-rollforward-q3-2026',
        title: 'Fix the Q3 2026 capital account roll-forward',
        instructions: 'In Investor Reporting/Capital Account Statements Q3 2026.xlsx, every partner\'s Ending Balance (Roll-forward!I6:I17) must equal Beginning + Contributions − Distributions − Management Fees + Realized + Unrealized. Two rows do not follow that rule; fix them so the Difference against the Fund Summary (Roll-forward!I21) is 0. Save the workbook.',
        file: 'Investor Reporting/Capital Account Statements Q3 2026.xlsx',
        checks: [
          { sheet: 'Roll-forward', cell: `I${6 + LP005}`, approx: round2(rf[LP005].end), tolerance: 0.01 },
          { sheet: 'Roll-forward', cell: `I${6 + LP009}`, approx: round2(rf[LP009].end), tolerance: 0.01 },
          { sheet: 'Roll-forward', cell: 'I18', approx: totalEnd, tolerance: 0.01 },
          { sheet: 'Roll-forward', cell: 'I21', approx: 0, tolerance: 0.005 },
        ],
      },
    ],
  };
}

rmSync(FILES, { recursive: true, force: true });
commitmentsRegister();
capitalCall4();
bankRec();
managementFee();
capitalAccounts();
const t = tasks();
mkdirSync(ROOT, { recursive: true });
writeFileSync(join(ROOT, 'tasks.json'), `${JSON.stringify(t, null, 2)}\n`);
console.log(`wrote ${written.length} workbooks to ${FILES}`);
for (const w of written) console.log(`  ${w}`);
console.log(`wrote ${t.tasks.length} tasks to ${join(ROOT, 'tasks.json')}`);
