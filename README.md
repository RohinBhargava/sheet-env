# sheet-env

ForkLaunch Sheets: a hosted, Excel-compatible spreadsheet that people and agents can work in. One Rust service (`src/workbooks`) serves the web client, stores workbooks (versioned, encrypted at rest), and exposes environment controls and MCP tools. The web client (`src/web`) runs the spreadsheet engine in the browser: formulas recalculate client side and every save writes cached values, so the server reads results with a plain xlsx reader.

## Layout

| Path | What |
|---|---|
| `src/workbooks/` | Rust service (ForkLaunch framework 0.2.5): API, MCP, embedded UI (`web-dist/`), seed scenarios (`seed/`) |
| `src/web/` | Vite + React client and the xlsx engine; `pnpm build` writes `src/workbooks/web-dist` |
| `src/workbooks/seed/<scenario>/` | `files/**.xlsx` loaded on reset, `tasks.json` with graded checks |

## API

| Method | Path | MCP | Notes |
|---|---|---|---|
| GET | `/api/files` | `listFiles` | current version of every workbook |
| GET | `/api/files/raw?path=` | – | .xlsx bytes |
| POST | `/api/files/save` | `saveFile` | `{path, kind?, contentBase64}` → next version |
| DELETE | `/api/files/raw?path=` | – | tombstone; history kept |
| GET | `/api/files/versions?path=` | `listVersions` | newest first |
| GET | `/api/files/cells?path=&sheet=&range=` | `readCells` | cached values, numbers, formulas (≤ 5000 cells) |
| GET | `/api/env/tasks` | `listTasks` | current scenario's tasks |
| POST | `/api/env/tasks/:id/verify` | – | grade against the last saved workbook |
| POST | `/api/env/reset` | – | `{scenario?, key?}`; `key` must match `ENV_CONTROL_KEY` when set |
| GET | `/api/env/state` | – | scenario, reset generation, state hash (timestamps excluded) |

MCP is served on the same port at `/mcp`. Reset, grading, state and raw bytes are kept out of MCP so an agent cannot reset or grade itself.

Saves are JSON with base64 content because framework 0.2.5 rejects binary request bodies that also carry query or path fields. With the 2 MiB request limit, a workbook can be about 1.5 MB.

## Storage

`file_versions` holds one row per saved version (path, size, SHA-256, timestamps; deletes are tombstones). `file_chunks` holds the bytes sealed with the framework's protected fields in chunks of at most 1 MiB, each its own encryption resource. Every plaintext read writes an `access_audit` row before the bytes are released. All tables use row-level security on `TENANT_ID`.

## Local run

```sh
cd src/web && pnpm install && pnpm build          # UI into src/workbooks/web-dist
cd src && cargo run -p workbooks -- --migrate      # needs MIGRATION_DATABASE_URL
cd src && cargo run -p workbooks                   # see src/workbooks/.env.local for settings
```

Deployed settings: the platform supplies the database URLs, `TENANT_ID`, `ENCRYPTION_KEY_HEX`, `ENCRYPTION_KEY_ID` and `STORAGE_ENCRYPTION_NAMESPACE`. The template sets `RUST_LOG`, `MCP_ENABLED` and `ENCRYPTION_LEGACY_KEY_HEX=disabled`. Each instance sets `PUBLIC_ORIGIN` to its own https origin (the framework rejects browser writes from any other origin) and optionally `ENV_CONTROL_KEY`.
