import { describe, expect, it } from "vitest";
import { excelRound } from "./excel-functions";
import { createEngine, toCellValue } from "./hyperformula";
import type { CellValue } from "./types";

function calc(formula: string, options: { todayIso?: string; mode?: "en" | "ka" } = {}): CellValue {
  const hf = createEngine(
    [
      [10, 30, 20, "'A-101", "15.01.2026"],
      [30, 1.25, 5.75, "'A-102", "31.03.2027"],
    ],
    options.mode ?? "en",
    { todayIso: options.todayIso },
  );
  try {
    const value = hf.calculateFormula(formula, 0);
    if (Array.isArray(value)) throw new Error("unexpected array");
    return toCellValue(value);
  } finally {
    hf.destroy();
  }
}

describe("custom Excel functions", () => {
  it("RANK.EQ and RANK rank like Excel (ties share a rank)", () => {
    expect(calc("=RANK.EQ(A1,A1:C1)")).toBe(3);
    expect(calc("=RANK.EQ(B1,A1:C1)")).toBe(1);
    expect(calc("=RANK(A1,A1:C1,1)")).toBe(1);
    expect(calc("=RANK.EQ(A2,A1:C2)")).toBe(1);
    expect(calc("=RANK.EQ(99,A1:C1)")).toMatchObject({ kind: "error", type: "NA" });
  });

  it("CONCAT joins values and ranges", () => {
    expect(calc('=CONCAT(D1:D2,"-",A1)')).toBe("A-101A-102-10");
    expect(calc("=CONCAT(TRUE,1.5)")).toBe("TRUE1.5");
  });

  it("ROUND rounds like Excel", () => {
    expect(calc("=ROUND(1.25*0.18,2)")).toBe(0.23);
    expect(calc("=ROUND(B2*0.18,2)")).toBe(0.23);
    expect(calc("=ROUND(C2*0.18,2)")).toBe(1.04);
    expect(calc("=ROUND(2.5,0)")).toBe(3);
    expect(calc("=ROUND(-2.5,0)")).toBe(-3);
    expect(calc("=ROUND(1234.5678,-2)")).toBe(1200);
    expect(calc("=ROUND(B2*0,18;2)", { mode: "ka" })).toBe(0.23);
    expect(excelRound(1.005, 2)).toBe(1.01);
  });

  it("TODAY returns the task's frozen date", () => {
    expect(calc("=TODAY()", { todayIso: "2026-11-02" })).toBe(46328);
    expect(calc("=TODAY()-E1", { todayIso: "2026-11-02" })).toBe(291);
    expect(calc('=TEXT(TODAY(),"dd.mm.yyyy")', { todayIso: "2026-11-02" })).toBe("02.11.2026");
  });

  it("DATEDIF accepts lower-case units", () => {
    expect(calc('=DATEDIF(E1,E2,"d")')).toBe(440);
    expect(calc('=DATEDIF(E1,E2,"m")')).toBe(14);
    expect(calc('=DATEDIF(E1,E2,"Y")')).toBe(1);
  });

  it("bare TRUE and FALSE work, e.g. VLOOKUP(…, FALSE)", () => {
    expect(calc('=VLOOKUP("A-102",D1:E2,1,FALSE)')).toBe("A-102");
    expect(calc("=IF(TRUE,1,0)")).toBe(1);
    expect(calc("=IF(FALSE;1;0)", { mode: "ka" })).toBe(0);
  });
});
