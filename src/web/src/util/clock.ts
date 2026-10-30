// The app clock. TODAY() and NOW() in formulas and the "saved at" times come from here. It follows the host
// clock unless a caller pins it (setScenarioClock), e.g. for reproducible task runs.
export interface ScenarioClockInfo { startMs?: number; endMs?: number; timeZone?: string }

const PLACEHOLDER_START_MS = Date.now();
let scenarioStartMs = PLACEHOLDER_START_MS;   // the task's scenario start (from the session once known)
let scenarioEndMs: number | null = null;
let anchorMs = PLACEHOLDER_START_MS;          // scenario time at `realAtAnchor`
let realAtAnchor = performance.now();

/** TIME_ZONE is a live binding: it follows the session's clock.timeZone once the backend has answered. */
export let TIME_ZONE = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';

export function setScenarioClock(nowMs: number, info?: ScenarioClockInfo) {
  anchorMs = nowMs;
  realAtAnchor = performance.now();
  if (info) {
    if (typeof info.startMs === 'number') scenarioStartMs = info.startMs;
    if (typeof info.endMs === 'number') scenarioEndMs = info.endMs;
    if (typeof info.timeZone === 'string' && info.timeZone) TIME_ZONE = info.timeZone;
  }
}

/** scenario start (epoch ms) as the backend declared it; the placeholder before the session answers */
export function scenarioStart(): number { return scenarioStartMs; }
/** scenario end (start + agent timeout), or null until the session answers */
export function scenarioEnd(): number | null { return scenarioEndMs; }

/** current scenario time in epoch milliseconds */
export function now(): number {
  return anchorMs + (performance.now() - realAtAnchor);
}

export function nowDate(): Date {
  return new Date(now());
}
