import { describe, expect, it } from "vitest";
import { countDecimals, createRng, perturbDateSerial, perturbNumber, randomFactor } from "./random";

describe("createRng", () => {
  it("is deterministic for a seed and stays in [0, 1)", () => {
    const a = createRng(42);
    const b = createRng(42);
    const values = Array.from({ length: 1000 }, () => a());
    expect(values).toEqual(Array.from({ length: 1000 }, () => b()));
    expect(values.every((v) => v >= 0 && v < 1)).toBe(true);
    expect(createRng(43)()).not.toBe(createRng(42)());
  });

  it("gives swap factors in [0.5, 1.5)", () => {
    const rng = createRng(7);
    for (let i = 0; i < 1000; i++) {
      const f = randomFactor(rng);
      expect(f).toBeGreaterThanOrEqual(0.5);
      expect(f).toBeLessThan(1.5);
    }
  });
});

describe("countDecimals", () => {
  it.each([
    [2500, 0],
    [2360.5, 1],
    [849.99, 2],
    [0.784, 3],
    [1e-7, 7],
    [-12.25, 2],
  ])("%d has %d decimals", (value, decimals) => {
    expect(countDecimals(value)).toBe(decimals);
  });
});

/** A fake generator that returns the given numbers in turn. */
const sequence = (...values: number[]) => {
  let i = 0;
  return () => values[i++ % values.length]!;
};

describe("perturbNumber (SCHEMA.md rule 4)", () => {
  it("multiplies by a factor in [0.5, 1.5] and keeps the decimals", () => {
    expect(perturbNumber(2500, sequence(0.2312))).toBe(1828); // × 0.7312
    expect(perturbNumber(849.99, sequence(0.6))).toBe(934.99); // × 1.1
    expect(perturbNumber(2360.5, sequence(0))).toBe(1180.3); // × 0.5
  });

  it("is never the original and never 0, keeps sign and decimals", () => {
    const rng = createRng(1);
    for (const value of [1, 2, 0.01, 0.02, -1, 3, 0.18, 1450, 89.9, -0.5]) {
      for (let i = 0; i < 200; i++) {
        const next = perturbNumber(value, rng);
        expect(next).not.toBe(value);
        expect(next).not.toBe(0);
        expect(Math.sign(next)).toBe(Math.sign(value));
        expect(countDecimals(next)).toBeLessThanOrEqual(countDecimals(value));
      }
    }
  });

  it("falls back to 1–3 units of the last decimal place when the factor gives the original", () => {
    // 1 × 1.2 rounds back to 1, so it moves by 1–3 units instead.
    expect(perturbNumber(1, sequence(0.7, 0.5))).toBe(3);
    expect(perturbNumber(0.01, sequence(0.5, 0))).toBe(0.02);
  });

  it("turns 0 into 1–9 units", () => {
    const rng = createRng(3);
    for (let i = 0; i < 100; i++) {
      const next = perturbNumber(0, rng);
      expect(Number.isInteger(next) && next >= 1 && next <= 9).toBe(true);
    }
  });
});

describe("perturbDateSerial", () => {
  it("moves a date 1–60 days, never 0", () => {
    const rng = createRng(11);
    for (let i = 0; i < 300; i++) {
      const shift = perturbDateSerial(46054, rng) - 46054;
      expect(Math.abs(shift)).toBeGreaterThanOrEqual(1);
      expect(Math.abs(shift)).toBeLessThanOrEqual(60);
      expect(Number.isInteger(shift)).toBe(true);
    }
  });
});
