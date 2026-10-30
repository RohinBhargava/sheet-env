// Minimal external store (no dependency): one immutable state object, subscribe + useStore selector.
import { useSyncExternalStore, useRef } from 'react';

export function createStore<T extends object>(initial: T) {
  let state = initial;
  const subs = new Set<() => void>();
  return {
    get: () => state,
    set(patch: Partial<T> | ((s: T) => Partial<T>)) {
      const p = typeof patch === 'function' ? patch(state) : patch;
      state = { ...state, ...p };
      subs.forEach(f => f());
    },
    subscribe(f: () => void) { subs.add(f); return () => { subs.delete(f); }; },
  };
}

export type Store<T extends object> = ReturnType<typeof createStore<T>>;

/** subscribe to a slice; re-renders only when the selected value changes (shallow for arrays/objects) */
export function useStore<T extends object, S>(store: Store<T>, sel: (s: T) => S): S {
  const last = useRef<{ v: S } | null>(null);
  return useSyncExternalStore(store.subscribe, () => {
    const v = sel(store.get());
    if (last.current && shallowEq(last.current.v, v)) return last.current.v;
    last.current = { v };
    return v;
  });
}

function shallowEq(a: unknown, b: unknown): boolean {
  if (Object.is(a, b)) return true;
  if (typeof a !== 'object' || typeof b !== 'object' || !a || !b) return false;
  if (Array.isArray(a) !== Array.isArray(b)) return false;
  const ka = Object.keys(a as object), kb = Object.keys(b as object);
  if (ka.length !== kb.length) return false;
  for (const k of ka) if (!Object.is((a as Record<string, unknown>)[k], (b as Record<string, unknown>)[k])) return false;
  return true;
}
