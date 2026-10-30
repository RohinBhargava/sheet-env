// Evaluates tasks.json checks against workbook files, reading the cached cell values the files carry (what the
// backend's non-calculating xlsx reader sees). Also flags any formula whose cached value differs from what the
// engine computes (a stale cache would grade wrong).
//
//   node scripts/check-tasks.ts                     seed files: every check must fail
//   node scripts/check-tasks.ts --solved            dev/solved copies: every check must pass
//   node scripts/check-tasks.ts --dir <root>        any Documents root; prints results only
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { Book } from './xl.ts';
import { isErr, unkey, addrName, type Scalar } from '../src/engine/index.ts';

const SCEN = resolve(import.meta.dirname, '../../workbooks/seed/fund-admin');
const args = process.argv.slice(2);
const solved = args.includes('--solved');
const dirArg = args.indexOf('--dir');
const root = dirArg >= 0 ? resolve(args[dirArg + 1]) : solved ? resolve(import.meta.dirname, '../dev/solved') : join(SCEN, 'files');
const expect: 'fail' | 'pass' | null = dirArg >= 0 ? null : solved ? 'pass' : 'fail';

interface Check { sheet: string; cell: string; equals?: number | string; approx?: number; tolerance?: number }
interface Task { id: string; file: string; checks: Check[] }
const { tasks } = JSON.parse(readFileSync(join(SCEN, 'tasks.json'), 'utf8')) as { tasks: Task[] };

function passes(c: Check, v: Scalar): boolean {
  if (c.approx !== undefined) return typeof v === 'number' && Math.abs(v - c.approx) <= (c.tolerance ?? 1e-9);
  if (typeof c.equals === 'number') return typeof v === 'number' && v === c.equals;
  return typeof v === 'string' && v === c.equals;
}

let bad = 0;
for (const t of tasks) {
  const p = join(root, t.file);
  if (!existsSync(p)) { console.log(`${t.id}: missing ${p}`); bad++; continue; }
  const b = Book.open(new Uint8Array(readFileSync(p)));
  const results = t.checks.map(c => {
    const v = b.cached(c.sheet, c.cell);
    return { c, v, ok: passes(c, v) };
  });
  const nPass = results.filter(r => r.ok).length;
  const verdict = expect === 'fail' ? (nPass === 0 ? 'OK (all fail)' : 'BAD (some pass on seed)') : expect === 'pass' ? (nPass === results.length ? 'OK (all pass)' : 'BAD (some fail)') : `${nPass}/${results.length} pass`;
  if (verdict.startsWith('BAD')) bad++;
  console.log(`${t.id}: ${verdict}`);
  for (const r of results) console.log(`  ${r.ok ? 'pass' : 'fail'}  ${r.c.sheet}!${r.c.cell} = ${JSON.stringify(r.v)}  (want ${r.c.approx !== undefined ? `≈${r.c.approx} ±${r.c.tolerance}` : JSON.stringify(r.c.equals)})`);
}

// cached values must match the engine's own calculation in every file under root
function walk(d: string): string[] {
  return readdirSync(d).flatMap(n => { const p = join(d, n); return statSync(p).isDirectory() ? walk(p) : /\.xlsx$/i.test(n) ? [p] : []; });
}
let stale = 0, formulas = 0;
for (const f of walk(root)) {
  const b = Book.open(new Uint8Array(readFileSync(f)));
  b.wb.sheets.forEach((sh, si) => {
    for (const [k, cell] of sh.cells) {
      if (!cell.f) continue;
      formulas++;
      const { r, c } = unkey(k);
      const want = b.wb.value(si, r, c), got = cell.v;
      const same = typeof want === 'number' && typeof got === 'number' ? Math.abs(want - got) < 1e-6 : isErr(want) && isErr(got) ? want.error === got.error : want === got;
      if (!same) { stale++; console.log(`  stale cache ${relative(root, f)} ${sh.name}!${addrName(r, c)}: file ${JSON.stringify(got)} vs engine ${JSON.stringify(want)}`); }
    }
  });
}
console.log(`cached formula values: ${formulas - stale}/${formulas} match the engine`);
if (stale) bad++;
process.exit(bad ? 1 : 0);
