// Produces solved copies of the fund-admin seed workbooks (for testing the task checks only):
// loads each seed file with the engine, makes the edits a person would type, saves to dev/solved/.
//
//   node scripts/solve.ts
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { Book } from './xl.ts';

const SEED = resolve(import.meta.dirname, '../../workbooks/seed/fund-admin/files');
const OUT = resolve(import.meta.dirname, '../dev/solved');

const solutions: Record<string, (b: Book) => void> = {
  'Capital Calls/Capital Call 4 Notice.xlsx': b => {
    for (let r = 6; r <= 17; r++) b.set('Allocation', `E${r}`, `=ROUND(D${r}*'Call Summary'!$C$12,2)`);
  },
  'Accounting/Bank Reconciliation Sep 2026.xlsx': b => {
    b.set('GL Cash', 'D9', 1_250_000);
    b.set('Reconciliation', 'B9', 'CHK 4473 Corran Advisory (09/24)');
    b.set('Reconciliation', 'C9', 9_800);
    b.set('Reconciliation', 'B15', 'Interest credit (09/30)');
    b.set('Reconciliation', 'C15', 3_412.57);
    b.set('Reconciliation', 'B18', 'Account maintenance fee (09/30)');
    b.set('Reconciliation', 'C18', 85);
  },
  'Accounting/Q4 2026 Management Fee.xlsx': b => {
    b.set('Fee Calc', 'C8', '=C6-C5+1');
    b.set('Fee Calc', 'C10', '=ROUND(C9*C4*C8/C7,2)');
    for (let r = 13; r <= 24; r++) b.set('Fee Calc', `F${r}`, `=IF(E${r}="Yes",ROUND(D${r}*$C$4*$C$8/$C$7,2),0)`);
  },
  'Investor Reporting/Capital Account Statements Q3 2026.xlsx': b => {
    for (const r of [10, 14]) b.set('Roll-forward', `I${r}`, `=C${r}+D${r}-E${r}-F${r}+G${r}+H${r}`);
  },
};

rmSync(OUT, { recursive: true, force: true });
for (const [rel, solve] of Object.entries(solutions)) {
  const b = Book.open(new Uint8Array(readFileSync(join(SEED, rel))));
  solve(b);
  const out = join(OUT, rel);
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, b.wb.save());
  console.log(`solved ${rel}`);
}
