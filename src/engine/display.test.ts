import { describe, expect, it } from "vitest";
import { alignmentOf, formatNumber, formatValue, needsApostrophe, valueToInput } from "./display";

describe("formatNumber (Excel General format)", () => {
  it.each([
    [1960, "1960", "1960"],
    [1960.5, "1960.5", "1960,5"],
    [-3.25, "-3.25", "-3,25"],
    [1 / 3, "0.333333333", "0,333333333"],
    [0.1 + 0.2, "0.3", "0,3"],
    [123456789.123, "123456789.1", "123456789,1"],
    [123456789012, "1.23457E+11", "1,23457E+11"],
    [1e-10, "1E-10", "1E-10"],
    [0, "0", "0"],
  ])("%d → %s / %s", (value, en, ka) => {
    expect(formatNumber(value, "en")).toBe(en);
    expect(formatNumber(value, "ka")).toBe(ka);
  });

  it("never uses thousands separators", () => {
    expect(formatNumber(2500000, "en")).toBe("2500000");
  });
});

describe("formatValue", () => {
  it("formats every kind of value", () => {
    expect(formatValue(null, "ka")).toBe("");
    expect(formatValue("გიორგი", "ka")).toBe("გიორგი");
    expect(formatValue(true, "ka")).toBe("TRUE");
    expect(formatValue({ kind: "error", type: "NA", code: "#N/A", detail: "" }, "ka")).toBe("#N/A");
  });

  it("formats percents, dates and times", () => {
    expect(formatValue(0.18, "en", { kind: "percent" })).toBe("18%");
    expect(formatValue(0.125, "ka", { kind: "percent" })).toBe("12,5%");
    expect(formatValue(46037, "ka", { kind: "date" })).toBe("15.01.2026");
    expect(formatValue(46037, "en", { kind: "date" })).toBe("15.01.2026");
    expect(formatValue(0.5208333333, "en", { kind: "time" })).toBe("12:30");
    expect(formatValue(46037.5, "ka", { kind: "datetime" })).toBe("15.01.2026 12:00");
  });

  it("formats currencies", () => {
    expect(formatValue(2500.5, "ka", { kind: "currency", currency: "₾" })).toBe("2500,5 ₾");
    expect(formatValue(5, "en", { kind: "currency", currency: "$" })).toBe("$5");
  });
});

describe("valueToInput (formula bar text)", () => {
  it("writes numbers so they can be typed back", () => {
    expect(valueToInput(2500.5, "ka")).toBe("2500,5");
    expect(valueToInput(2500.5, "en")).toBe("2500.5");
    expect(valueToInput(0.125, "ka", { kind: "percent" })).toBe("12,5%");
    expect(valueToInput(46037, "en", { kind: "date" })).toBe("15.01.2026");
    expect(valueToInput(5, "ka", { kind: "currency", currency: "$" })).toBe("$5");
  });

  it("adds an apostrophe to text that would otherwise change type", () => {
    expect(valueToInput("01001012345", "en")).toBe("'01001012345");
    expect(valueToInput("TRUE", "en")).toBe("'TRUE");
    expect(valueToInput("=B2", "en")).toBe("'=B2");
    expect(valueToInput("გიორგი", "en")).toBe("გიორგი");
  });
});

describe("needsApostrophe", () => {
  it("detects text that looks like a number, date, boolean or error", () => {
    expect(needsApostrophe("2500,5")).toBe(true);
    expect(needsApostrophe("15.01.2026")).toBe(true);
    expect(needsApostrophe("false")).toBe(true);
    expect(needsApostrophe("#N/A")).toBe(true);
    expect(needsApostrophe("INV-101")).toBe(false);
    expect(needsApostrophe("")).toBe(false);
  });
});

describe("alignmentOf", () => {
  it("follows Excel's defaults", () => {
    expect(alignmentOf(5)).toBe("right");
    expect(alignmentOf("a")).toBe("left");
    expect(alignmentOf(true)).toBe("center");
  });
});
