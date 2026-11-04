/**
 * Deterministic randomness for the swap test (SCHEMA.md grading rule 4).
 * The same seed always gives the same "random" data, in the browser, in Node and in the checker.
 */

/** Mulberry32: a tiny, fast, seeded generator. Returns numbers in [0, 1). */
export function createRng(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4_294_967_296;
  };
}

/** A swap-test factor in [0.5, 1.5). */
export function randomFactor(rng: () => number): number {
  return 0.5 + rng();
}

/** A whole number from `min` to `max` (inclusive). */
function randomInt(rng: () => number, min: number, max: number): number {
  return min + Math.floor(rng() * (max - min + 1));
}

/** Decimal places of a number as written: 2500 → 0, 2360.5 → 1, 0.784 → 3, 1e-7 → 7 (max 15). */
export function countDecimals(value: number): number {
  if (!Number.isFinite(value) || Number.isInteger(value)) return 0;
  const match = /^(?:\d+)(?:\.(\d+))?(?:e([+-]\d+))?$/i.exec(Math.abs(value).toString());
  if (!match) return 0;
  const fraction = match[1]?.length ?? 0;
  const exponent = match[2] ? Number(match[2]) : 0;
  return Math.min(Math.max(fraction - exponent, 0), 15);
}

/** Rounds half away from zero to `decimals` places, avoiding binary floating-point surprises. */
function roundTo(value: number, decimals: number): number {
  const sign = value < 0 ? -1 : 1;
  const scale = 10 ** decimals;
  // toPrecision(15) turns 100.49999999999999 (from 1.005 × 100) back into 100.5 before rounding.
  const shifted = Number((Math.abs(value) * scale).toPrecision(15));
  return (sign * Math.round(shifted)) / scale;
}

/**
 * A new value for a `vary` number (SCHEMA.md rule 4): `value` × a random factor in [0.5, 1.5],
 * with the same number of decimal places. If that gives back `value` or 0, it moves 1–3 units of
 * the last decimal place away from zero instead; 0 becomes 1–9 units. Never the original, never 0.
 */
export function perturbNumber(value: number, rng: () => number): number {
  if (!Number.isFinite(value)) return value;
  const decimals = countDecimals(value);
  const unit = 10 ** -decimals;
  if (value === 0) return roundTo(randomInt(rng, 1, 9) * unit, decimals);
  const scaled = roundTo(value * randomFactor(rng), decimals);
  if (scaled !== value && scaled !== 0) return scaled;
  return roundTo(value + Math.sign(value) * randomInt(rng, 1, 3) * unit, decimals);
}

/** A new value for a `vary` date (Excel serial): 1–60 days earlier or later. */
export function perturbDateSerial(serial: number, rng: () => number): number {
  const days = randomInt(rng, 1, 60);
  return rng() < 0.5 ? serial - days : serial + days;
}
