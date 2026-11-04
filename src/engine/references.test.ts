import { describe, expect, it } from "vitest";
import { cycleAbsolute } from "./references";

/** Presses F4 `times` times with the caret at `caret`, returns every formula seen. */
function pressF4(formula: string, caret: number, times: number): string[] {
  const seen: string[] = [];
  let state = { formula, caret };
  for (let i = 0; i < times; i++) {
    state = cycleAbsolute(state.formula, state.caret);
    seen.push(state.formula);
  }
  return seen;
}

describe("cycleAbsolute (F4)", () => {
  it("cycles A1 → $A$1 → A$1 → $A1 → A1", () => {
    expect(pressF4("=A1", 3, 4)).toEqual(["=$A$1", "=A$1", "=$A1", "=A1"]);
  });

  it("puts the caret right after the changed reference", () => {
    expect(cycleAbsolute("=B2*(1+E1)", 9)).toEqual({ formula: "=B2*(1+$E$1)", caret: 11 });
  });

  it("acts on the reference containing the caret", () => {
    expect(cycleAbsolute("=B2*(1+E1)", 8).formula).toBe("=B2*(1+$E$1)");
    expect(cycleAbsolute("=B12*C3", 2).formula).toBe("=$B$12*C3");
  });

  it("prefers the reference before the caret, then the one after it", () => {
    expect(cycleAbsolute("=B2+C3", 3).formula).toBe("=$B$2+C3");
    expect(cycleAbsolute("=B2+C3", 4).formula).toBe("=B2+$C$3");
    expect(cycleAbsolute("=B2+C3", 1).formula).toBe("=$B$2+C3");
  });

  it("changes both ends of a range together", () => {
    expect(pressF4("=SUM(B2:B6)", 8, 4)).toEqual([
      "=SUM($B$2:$B$6)",
      "=SUM(B$2:B$6)",
      "=SUM($B2:$B6)",
      "=SUM(B2:B6)",
    ]);
  });

  it("starts from the current state of mixed references", () => {
    expect(cycleAbsolute("=$A$1", 5).formula).toBe("=A$1");
    expect(cycleAbsolute("=A$1", 4).formula).toBe("=$A1");
    expect(cycleAbsolute("=$A1", 4).formula).toBe("=A1");
  });

  it("toggles whole-column and whole-row ranges", () => {
    expect(cycleAbsolute("=SUM(A:C)", 6).formula).toBe("=SUM($A:$C)");
    expect(cycleAbsolute("=SUM($A:$C)", 6).formula).toBe("=SUM(A:C)");
    expect(cycleAbsolute("=SUM(1:3)", 6).formula).toBe("=SUM($1:$3)");
  });

  it("does nothing inside text, on function names or without a reference", () => {
    expect(cycleAbsolute('="A1"&B2', 3)).toEqual({ formula: '="A1"&B2', caret: 3 });
    expect(cycleAbsolute("=LOG10(100)", 4)).toEqual({ formula: "=LOG10(100)", caret: 4 });
    expect(cycleAbsolute("=1+2", 2)).toEqual({ formula: "=1+2", caret: 2 });
    expect(cycleAbsolute("B2", 1)).toEqual({ formula: "B2", caret: 1 });
  });

  it("works on Georgian-format formulas", () => {
    expect(cycleAbsolute("=ROUND(B2*0,784;2)", 9, "ka").formula).toBe("=ROUND($B$2*0,784;2)");
  });

  it("clamps an out-of-range caret", () => {
    expect(cycleAbsolute("=A1", 99)).toEqual({ formula: "=$A$1", caret: 5 });
  });
});
