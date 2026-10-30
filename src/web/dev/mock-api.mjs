// Development stand-in for the workbooks service's file API, for running the web app without the Rust service.
// Files live in a temp directory seeded from src/workbooks/seed/<scenario>/files.
//
//   node dev/mock-api.mjs [--port 8791] [--dir <documents root>]
//
//   GET    /api/files                      {"files":[{path,size,modifiedMs,version}]}
//   GET    /api/files/raw?path=            xlsx bytes (404 if missing)
//   POST   /api/files/save                 {"path","kind","contentBase64"} -> {path,size,modifiedMs,version}
//   DELETE /api/files/raw?path=            204
//   GET    /api/files/versions?path=       {"versions":[{version,modifiedMs,size}]}
//   GET    /api/env/tasks                  {"tasks":[{id,title,instructions,file}]}
//   POST   /api/env/reset {"scenario"}     reseeds the documents root
import { createServer } from 'node:http';
import { cpSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, unlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, relative, resolve, sep } from 'node:path';

const args = process.argv.slice(2);
const opt = (name, dflt) => { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : dflt; };
const PORT = Number(opt('--port', process.env.PORT ?? 8791));
const SEEDS = resolve(import.meta.dirname, '../../workbooks/seed');
const DOCS = resolve(opt('--dir', mkdtempSync(join(tmpdir(), 'sheet-docs-'))));
const MAX_BODY = 2 * 1024 * 1024;

/** path -> [{version, modifiedMs, size}] */
const versions = new Map();

function seed(scenario = 'fund-admin') {
  const src = join(SEEDS, scenario, 'files');
  if (!existsSync(src)) throw new Error(`unknown scenario ${scenario}`);
  rmSync(DOCS, { recursive: true, force: true });
  mkdirSync(DOCS, { recursive: true });
  cpSync(src, DOCS, { recursive: true });
  versions.clear();
  for (const p of walk(DOCS)) { const st = statSync(join(DOCS, p)); versions.set(p, [{ version: 1, modifiedMs: Math.round(st.mtimeMs), size: st.size }]); }
}

function walk(dir, base = dir) {
  return readdirSync(dir).flatMap(n => {
    const p = join(dir, n);
    return statSync(p).isDirectory() ? walk(p, base) : [relative(base, p).split(sep).join('/')];
  });
}

/** Documents-relative path -> absolute path, refusing anything that escapes the root */
function resolvePath(p) {
  if (typeof p !== 'string' || !p || p.startsWith('/') || p.includes('\\') || p.split('/').some(s => s === '..' || s === '')) return null;
  const abs = resolve(DOCS, p);
  return abs.startsWith(DOCS + sep) ? abs : null;
}

function entry(p) {
  const st = statSync(join(DOCS, p));
  const v = versions.get(p) ?? [];
  return { path: p, size: st.size, modifiedMs: Math.round(st.mtimeMs), version: v.length ? v[v.length - 1].version : 1 };
}

function send(res, status, body, type = 'application/json') {
  const data = body === undefined ? '' : type === 'application/json' ? JSON.stringify(body) : body;
  res.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(data);
}

function readBody(req) {
  return new Promise((ok, fail) => {
    const chunks = []; let n = 0;
    req.on('data', c => { n += c.length; if (n > MAX_BODY) { fail(Object.assign(new Error('request too large'), { status: 413 })); req.destroy(); } else chunks.push(c); });
    req.on('end', () => ok(Buffer.concat(chunks)));
    req.on('error', fail);
  });
}

const server = createServer(async (req, res) => {
  try {
    const url = new URL(req.url, 'http://localhost');
    const route = `${req.method} ${url.pathname}`;
    if (route === 'GET /api/files') return send(res, 200, { files: walk(DOCS).sort().map(entry) });
    if (url.pathname === '/api/files/raw') {
      const abs = resolvePath(url.searchParams.get('path'));
      if (!abs) return send(res, 400, { error: 'invalid path' });
      if (req.method === 'GET') return existsSync(abs) ? send(res, 200, readFileSync(abs), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') : send(res, 404, { error: 'not found' });
      if (req.method === 'DELETE') { if (existsSync(abs)) unlinkSync(abs); versions.delete(url.searchParams.get('path')); return send(res, 204); }
      return send(res, 405, { error: 'method not allowed' });
    }
    if (route === 'POST /api/files/save') {
      const body = JSON.parse((await readBody(req)).toString('utf8'));
      const abs = resolvePath(body.path);
      if (!abs || !/\.xlsx$/i.test(body.path)) return send(res, 400, { error: 'invalid path' });
      if (body.kind !== undefined && !['save', 'save_as', 'upload'].includes(body.kind)) return send(res, 400, { error: 'invalid kind' });
      const bytes = Buffer.from(String(body.contentBase64 ?? ''), 'base64');
      if (bytes.length < 4 || bytes.readUInt32LE(0) !== 0x04034b50) return send(res, 400, { error: 'not an xlsx (zip) file' });
      mkdirSync(dirname(abs), { recursive: true });
      writeFileSync(abs, bytes);
      const list = versions.get(body.path) ?? [];
      const e = { ...entry(body.path), version: (list.at(-1)?.version ?? 0) + 1 };
      list.push({ version: e.version, modifiedMs: e.modifiedMs, size: e.size });
      versions.set(body.path, list);
      console.log(`save ${body.kind ?? 'save'} ${body.path} (${bytes.length} bytes, v${e.version})`);
      return send(res, 200, e);
    }
    if (route === 'GET /api/files/versions') {
      const p = url.searchParams.get('path');
      if (!resolvePath(p)) return send(res, 400, { error: 'invalid path' });
      return send(res, 200, { versions: versions.get(p) ?? [] });
    }
    if (route === 'GET /api/env/tasks') {
      const t = JSON.parse(readFileSync(join(SEEDS, 'fund-admin', 'tasks.json'), 'utf8'));
      return send(res, 200, { tasks: t.tasks.map(({ id, title, instructions, file }) => ({ id, title, instructions, file })) });
    }
    if (route === 'POST /api/env/reset') {
      const raw = (await readBody(req)).toString('utf8');
      seed((raw ? JSON.parse(raw).scenario : undefined) ?? 'fund-admin');
      return send(res, 200, { ok: true });
    }
    send(res, 404, { error: 'not found' });
  } catch (e) {
    send(res, e.status ?? 500, { error: e.message });
  }
});

seed();
server.listen(PORT, '127.0.0.1', () => console.log(`mock workbooks API on http://127.0.0.1:${PORT}, documents in ${DOCS}`));
