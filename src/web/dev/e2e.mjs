// End-to-end check of the web app against a running API (mock or real) and a running dev/preview server.
//
//   node dev/e2e.mjs [--url http://127.0.0.1:5491] [--shots <dir>]
//
// Loads Home, opens every workbook, enters a formula, saves with Ctrl+S, reloads and confirms the value
// persisted; does the Capital Call No. 4 task through the UI; creates a blank workbook through Save As.
import { chromium } from 'playwright';
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { Workbook, parseAddr } from '../src/engine/index.ts';

const args = process.argv.slice(2);
const opt = (n, d) => { const i = args.indexOf(n); return i >= 0 ? args[i + 1] : d; };
const BASE = opt('--url', 'http://127.0.0.1:5491');
const SHOTS = opt('--shots', join(import.meta.dirname, 'shots'));
mkdirSync(SHOTS, { recursive: true });

let failures = 0;
const ok = (cond, msg) => { console.log(`${cond ? 'PASS' : 'FAIL'}  ${msg}`); if (!cond) failures++; };

async function cachedValue(path, sheet, a1) {
  const res = await fetch(`${BASE}/api/files/raw?path=${encodeURIComponent(path)}`);
  const wb = Workbook.load(new Uint8Array(await res.arrayBuffer()));
  const si = sheet ? wb.sheetIndex(sheet) : 0;
  const a = parseAddr(a1);
  return wb.cell(si, a.r, a.c)?.v ?? null;
}

// start from the scenario's seed files (the run edits them)
const reset = await fetch(`${BASE}/api/env/reset`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ scenario: 'fund-admin' }) });
if (!reset.ok) throw new Error(`reset failed: ${reset.status}`);

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
const errors = [];
page.on('pageerror', e => errors.push(String(e)));
page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });

async function waitBook() {
  await page.waitForSelector('.xl-root[tabindex] .xl-grid', { timeout: 15000 });
  await page.waitForTimeout(300);
}
async function goto(ref) {
  const nb = page.getByLabel('Name Box', { exact: true });
  await nb.click();
  await nb.fill(ref);
  await nb.press('Enter');
  await page.waitForTimeout(150);
  // under fast automation the Name Box sometimes keeps the focus after Enter (inherited behaviour): hand it back
  if (await page.evaluate(() => document.activeElement?.getAttribute('aria-label') === 'Name Box')) {
    nameBoxFocusFixes++;
    await page.evaluate(() => document.querySelector('.xl-root[tabindex]')?.focus());
  }
  const shown = await nb.inputValue();
  if (shown !== ref.toUpperCase()) throw new Error(`Name Box shows ${shown}, wanted ${ref}`);
}
let nameBoxFocusFixes = 0;
async function saveAndWait() {
  const resp = page.waitForResponse(r => r.url().includes('/api/files/save') && r.request().method() === 'POST', { timeout: 10000 });
  await page.keyboard.press('Control+s');
  return (await resp).ok();
}

// ---------------------------------------------------------------- Home
await page.goto(BASE);
await page.waitForSelector('.home-tr[data-path]');
const paths = await page.$$eval('.home-tr[data-path]', els => els.map(e => e.getAttribute('data-path')));
ok(paths.length === 5, `Home lists ${paths.length} workbooks`);
ok(await page.title() === 'ForkLaunch Sheets', `Home title "${await page.title()}"`);
await page.screenshot({ path: join(SHOTS, 'home.png') });

// ---------------------------------------------------------------- every workbook: edit, Ctrl+S, reload
for (const [i, p] of paths.entries()) {
  await page.goto(BASE);
  await page.click(`.home-tr[data-path="${p.replace(/"/g, '\\"')}"]`);
  await waitBook();
  ok(new URL(page.url()).searchParams.get('file') === p, `URL follows ${p}`);
  const title = await page.title();
  ok(/ - ForkLaunch Sheets$/.test(title), `title "${title}"`);
  const sheet = await page.$eval('.xl-sheettab.active', e => e.textContent);
  const n = 7 + i;
  await goto('Z40');
  await page.keyboard.type(`=6*${n}+0.5`);
  await page.keyboard.press('Enter');
  ok(await saveAndWait(), `Ctrl+S saved ${p}`);
  await page.waitForTimeout(300);
  await page.reload();
  await waitBook();
  await goto('Z40');
  const fbar = await page.getByLabel('Formula Bar').inputValue();
  ok(fbar === `=6*${n}+0.5`, `after reload Z40 formula is "${fbar}"`);
  const v = await cachedValue(p, sheet, 'Z40');
  ok(v === 6 * n + 0.5, `saved file caches Z40 = ${v} on ${sheet}`);
}

// ---------------------------------------------------------------- the Capital Call No. 4 task through the UI
const call = 'Capital Calls/Capital Call 4 Notice.xlsx';
await page.goto(`${BASE}/?file=${encodeURIComponent(call)}`);
await waitBook();
await page.click('.xl-sheettab:has-text("Allocation")');
await page.waitForTimeout(200);
await goto('E6');
for (let r = 6; r <= 17; r++) {
  await page.keyboard.type(`=ROUND(D${r}*'Call Summary'!$C$12,2)`);
  await page.keyboard.press('Enter');
}
ok(await saveAndWait(), 'saved the allocation');
await page.waitForTimeout(300);
await goto('E21');
await page.screenshot({ path: join(SHOTS, 'workbook-capital-call.png') });
ok(await cachedValue(call, 'Allocation', 'E18') === 16500000, 'allocation total 16,500,000 in the saved file');
ok(await cachedValue(call, 'Allocation', 'E21') === 0, 'allocation difference 0 in the saved file');

// ---------------------------------------------------------------- bank reconciliation view (screenshot)
await page.goto(`${BASE}/?file=${encodeURIComponent('Accounting/Bank Reconciliation Sep 2026.xlsx')}`);
await waitBook();
await page.screenshot({ path: join(SHOTS, 'workbook-bank-rec.png') });

// ---------------------------------------------------------------- blank workbook -> Save As
await page.goto(BASE);
await page.waitForSelector('.home-tr[data-path]');
await page.click('button[aria-label="Blank workbook"]');
await waitBook();
await page.keyboard.type('Hello');
await page.keyboard.press('Tab');
await page.keyboard.type('=LEN(A1)');
await page.keyboard.press('Enter');
await page.keyboard.press('Control+s');
const dlg = page.getByRole('dialog', { name: 'Save As' });
await dlg.waitFor();
await page.screenshot({ path: join(SHOTS, 'save-as.png') });
await dlg.getByLabel('Folder').fill('Scratch');
await dlg.getByLabel('File name').fill('E2E Blank');
const saved = page.waitForResponse(r => r.url().includes('/api/files/save'));
await dlg.getByRole('button', { name: 'Save' }).click();
ok((await saved).ok(), 'Save As wrote Scratch/E2E Blank.xlsx');
await page.waitForTimeout(300);
ok(new URL(page.url()).searchParams.get('file') === 'Scratch/E2E Blank.xlsx', 'URL follows the Save As path');
ok(await cachedValue('Scratch/E2E Blank.xlsx', null, 'B1') === 5, 'blank workbook B1 =LEN(A1) cached 5');

// ---------------------------------------------------------------- dirty close prompts (in-page message box)
await page.keyboard.type('changed');
await page.keyboard.press('Enter');
await page.click('button[aria-label="Home"]');
const prompt = page.getByRole('dialog').filter({ hasText: 'Want to save your changes' });
await prompt.waitFor({ timeout: 5000 });
await page.screenshot({ path: join(SHOTS, 'save-prompt.png') });
await prompt.getByRole('button', { name: "Don't Save" }).click();
await page.waitForSelector('.home-tr[data-path]');
ok((await page.$$('.home-tr[data-path]')).length === 6, 'Home lists the new workbook');

// ---------------------------------------------------------------- embed + panel
await page.goto(`${BASE}/?embed=1&panel=1&file=${encodeURIComponent('Accounting/Q4 2026 Management Fee.xlsx')}`);
await waitBook();
await page.waitForSelector('.task');
ok((await page.$$('.task')).length === 4, 'tasks panel lists 4 tasks');
ok(!(await page.$('button[aria-label="Home"]')), 'embed hides Home');
await page.screenshot({ path: join(SHOTS, 'embed-panel.png') });

console.log(`(Name Box kept focus after Enter ${nameBoxFocusFixes} time(s))`);
ok(errors.length === 0, `no page errors${errors.length ? `: ${errors.join(' | ')}` : ''}`);
await browser.close();
console.log(failures ? `${failures} failure(s)` : 'all passed');
process.exit(failures ? 1 : 0);
