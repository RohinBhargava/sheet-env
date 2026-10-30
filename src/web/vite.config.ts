import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';

// Built into the workbooks service's web-dist (served from '/'). In dev and preview, /api proxies to the
// workbooks service (or the mock in dev/mock-api.mjs): http://127.0.0.1:8791 unless SHEET_API_URL says otherwise.
const api = process.env.SHEET_API_URL ?? 'http://127.0.0.1:8791';
export default defineConfig({
  plugins: [react()],
  base: '/',
  server: { port: 5491, strictPort: true, proxy: { '/api': api } },
  preview: { port: 5491, strictPort: true, proxy: { '/api': api } },
  build: { outDir: fileURLToPath(new URL('../workbooks/web-dist', import.meta.url)), emptyOutDir: true, sourcemap: false },
});
