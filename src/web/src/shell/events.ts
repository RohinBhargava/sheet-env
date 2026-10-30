// Session (who is signed in, shown as the account initials) and the app event log.
// Events are kept in memory (window.__sheetEvents) and printed to the console in development; the backend
// grades from the saved files, not from these events.
import { createStore } from './store';

export interface Persona { name: string; initials: string }

export const session = createStore<{ persona: Persona }>({ persona: { name: 'Guest', initials: 'G' } });

export interface AppEvent { t: number; app: string; event: string; detail: Record<string, unknown> }
const log: AppEvent[] = [];
(globalThis as { __sheetEvents?: AppEvent[] }).__sheetEvents = log;

export function logAppEvent(app: string, event: string, detail: Record<string, unknown> = {}) {
  const e = { t: Date.now(), app, event, detail };
  log.push(e);
  if (log.length > 500) log.shift();
  if (import.meta.env.DEV) console.debug('[event]', event, detail);
}
