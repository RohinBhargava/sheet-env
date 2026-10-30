// Mouse-wheel notches from a WheelEvent.
// Chromium coalesces the wheel events queued while the main thread is busy into ONE event carrying the summed delta
// (a Word relayout after typing stalls it for about a second: xdotool's 30 notches at 30 ms arrived as one 120 px
// event and one 3720 px event). Anything that scrolls a fixed step per notch must therefore count notches from the
// delta, never from the number of events.

/** pixel delta of one wheel notch on this platform: the smallest notch-sized delta seen so far (X11 Chromium 120,
 *  CDP input 100); a coalesced event carries a multiple of it */
let unit = 120;

/** Smallest pixel delta that counts as a wheel notch; below it (precision touchpads) callers keep native scrolling. */
export const NOTCH_MIN_PX = 40;

/** Signed notch count of a wheel event along one axis delta (0 when the delta is touchpad-sized). */
export function wheelNotches(delta: number, deltaMode: number): number {
  const d = Math.abs(delta);
  if (deltaMode === 1) return d >= 1 ? Math.sign(delta) * Math.max(1, Math.round(d / 3)) : 0;   // lines: 3 per notch
  if (deltaMode === 2) return d > 0 ? Math.sign(delta) * Math.max(1, Math.round(d)) : 0;        // pages
  if (d < NOTCH_MIN_PX) return 0;
  unit = Math.min(unit, d);
  return Math.sign(delta) * Math.max(1, Math.round(d / unit));
}
