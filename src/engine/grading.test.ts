import { describe, expect, it } from "vitest";
import {
  describeMustUseKa,
  gradeTask,
  satisfiesMustUse,
  solutionFormulas,
  swapTrialInputs,
  valuesMatch,
  type GradeFail,
  type GradeResult,
} from "./grading";
import type { GradableTask } from "./schema";
import { SheetSession } from "./workbook";

/** Build plan §6.1 example: Gross → Net in one cell. */
const payroll: GradableTask = {
  sheet: {
    cells: [
      ["თანამშრომელი", "დარიცხული (₾)", "ხელზე ასაღები (₾)"],
      ["გიორგი ბერიძე", 2500, null],
    ],
    editable: ["C2"],
  },
  check: { cells: "C2", solution: "=B2*0.98*0.8", vary: ["B2"], mustUse: [] },
};

/** VAT extraction where the answer must be rounded. */
const vatRound: GradableTask = {
  sheet: { cells: [["თანხა დღგ-ით", "დღგ"], [1180, null]], editable: ["B2"] },
  check: { cells: "B2", solution: "=ROUND(A2/1.18*0.18,2)", vary: ["A2"] },
};

/** Add VAT with one rate cell for the whole column: needs an absolute reference. */
const vatColumn: GradableTask = {
  sheet: {
    cells: [
      ["პროდუქტი", "ფასი (₾)", "ფასი დღგ-ით (₾)", "განაკვეთი", 0.18],
      ["მაცივარი", 1450, null],
      ["ტელევიზორი", 1199, null],
      ["მტვერსასრუტი", 320, null],
      ["მიკროტალღური", 279, null],
      ["უთო", 89.9, null],
    ],
    editable: ["C2:C6"],
  },
  check: { cells: "C2:C6", solution: "=B2*(1+$E$1)", vary: ["B2:B6", "E1"] },
};

/** Price lookup (review case 5): unsorted keys, prices varied, cases put other keys in A2. */
const lookup: GradableTask = {
  sheet: {
    cells: [
      ["კოდი", "ფასი", null, "კოდი", "ფასი"],
      ["A-103", null, null, "A-102", 12],
      [null, null, null, "A-104", 3],
      [null, null, null, "A-101", 10.5],
      [null, null, null, "A-103", 7.25],
    ],
    editable: ["B2"],
  },
  check: {
    cells: "B2",
    solution: "=VLOOKUP(A2,$D$2:$E$5,2,FALSE)",
    vary: ["E2:E5"],
    mustUse: ["VLOOKUP", "XLOOKUP", ["INDEX", "MATCH"]],
    cases: [{ A2: "A-101" }, { A2: "A-104" }],
  },
};

/** Sales bonus with a threshold: cases at the boundary and just below it. */
const bonus: GradableTask = {
  sheet: { cells: [["გაყიდვები", "ბონუსი"], [2500, null]], editable: ["B2"] },
  check: {
    cells: "B2",
    solution: "=IF(A2>=2000,A2*0.05,0)",
    vary: ["A2"],
    cases: [{ A2: 2000 }, { A2: 1999.99 }],
  },
};

function fail(result: GradeResult): GradeFail {
  if (result.pass) throw new Error(`expected a failing grade, got: ${result.messageKa}`);
  return result;
}

describe("gradeTask: the build-plan cases (payroll)", () => {
  it("passes =B2*0.784", () => {
    expect(gradeTask(payroll, { C2: "=B2*0.784" }, "en").pass).toBe(true);
  });

  it("passes other correct formulas: =0.784*B2, =B2*0.98*0.8, =B2-B2*0.216", () => {
    expect(gradeTask(payroll, { C2: "=0.784*B2" }, "en").pass).toBe(true);
    expect(gradeTask(payroll, { C2: "=B2*0.98*0.8" }, "en").pass).toBe(true);
    expect(gradeTask(payroll, { C2: "=B2-B2*0.216" }, "en").pass).toBe(true);
  });

  it("rejects the typed number 1960: not a formula", () => {
    const result = fail(gradeTask(payroll, { C2: "1960" }, "en"));
    expect(result).toMatchObject({ reason: "not_formula", cell: "C2", actual: "1960" });
    expect(result.messageKa).toBe("C2: ჩაწერეთ ფორმულა და არა მზა რიცხვი. ფორმულა იწყება „=“ ნიშნით.");
  });

  it("rejects =2500*0.784 as hardcoded (it only fails the swap test)", () => {
    const result = fail(gradeTask(payroll, { C2: "=2500*0.784" }, "en"));
    expect(result).toMatchObject({ reason: "hardcoded", trial: 1, actual: "1960" });
    expect(result.trialInputs).toEqual(swapTrialInputs(payroll)[0]);
    expect(result.expected).not.toBe("1960");
    // Without the swap test it would have passed.
    expect(gradeTask(payroll, { C2: "=2500*0.784" }, "en", { trials: 0 }).pass).toBe(true);
  });

  it("passes =ROUND(B2*0,784;2) in Georgian mode when the lesson rounds (or allows rounding)", () => {
    const rounded = { ...payroll, check: { ...payroll.check, solution: "=ROUND(B2*0.98*0.8,2)" } };
    expect(gradeTask(rounded, { C2: "=ROUND(B2*0,784;2)" }, "ka").pass).toBe(true);
    // SCHEMA.md: tolerance 0.005 means "rounded or not are both acceptable".
    const either = { ...payroll, check: { ...payroll.check, tolerance: 0.005 } };
    expect(gradeTask(either, { C2: "=ROUND(B2*0,784;2)" }, "ka").pass).toBe(true);
    expect(gradeTask(either, { C2: "=B2*0,784" }, "ka").pass).toBe(true);
    // With the default tolerance an unrounded solution does not accept a rounded answer.
    expect(fail(gradeTask(payroll, { C2: "=ROUND(B2*0,784;2)" }, "ka")).reason).toBe("hardcoded");
  });

  it("does not accept Georgian format in English mode, or English format in Georgian mode", () => {
    expect(fail(gradeTask(payroll, { C2: "=ROUND(B2*0,784;2)" }, "en"))).toMatchObject({
      reason: "excel_error",
      errorType: "ERROR",
    });
    expect(fail(gradeTask(payroll, { C2: "=ROUND(B2*0.784,2)" }, "ka")).reason).toBe("excel_error");
  });

  it("reports an empty cell", () => {
    expect(fail(gradeTask(payroll, {}, "en"))).toMatchObject({ reason: "empty", cell: "C2" });
    expect(fail(gradeTask(payroll, { C2: "  " }, "en")).reason).toBe("empty");
  });

  it("reports a wrong value with expected and actual, formatted for the mode", () => {
    const result = fail(gradeTask(payroll, { C2: "=B2*0,78" }, "ka"));
    expect(result).toMatchObject({ reason: "wrong_value", expected: "1960", actual: "1950" });
    expect(fail(gradeTask(payroll, { C2: "=B2*0,7841" }, "ka")).actual).toBe("1960,25");
  });

  it("reports an Excel error with its type and a Georgian explanation", () => {
    const result = fail(gradeTask(payroll, { C2: "=B2/A3" }, "en"));
    expect(result).toMatchObject({ reason: "excel_error", errorType: "DIV_BY_ZERO", actual: "#DIV/0!" });
    expect(result.messageKa).toContain("ნულზე გაყოფა");
  });

  it("uses a relative tolerance (default 1e-9)", () => {
    expect(gradeTask(payroll, { C2: "=B2*0.784+0.0000001" }, "en").pass).toBe(true);
    expect(gradeTask(payroll, { C2: "=B2*0.784+0.001" }, "en").pass).toBe(false);
    const loose = { ...payroll, check: { ...payroll.check, tolerance: 0.005 } };
    expect(gradeTask(loose, { C2: "=ROUND(B2*0.784,-1)" }, "en").pass).toBe(true);
  });

  it("compareDecimals: rounded or not are both accepted, wrong values still fail", () => {
    const tetri = { ...payroll, check: { ...payroll.check, compareDecimals: 2 } };
    expect(gradeTask(tetri, { C2: "=ROUND(B2*0,784;2)" }, "ka").pass).toBe(true);
    expect(gradeTask(tetri, { C2: "=B2*0.784" }, "en").pass).toBe(true);
    expect(gradeTask(tetri, { C2: "=ROUND(B2*0.784,1)" }, "en").pass).toBe(false);
    expect(gradeTask(tetri, { C2: "=B2*0.78" }, "en").pass).toBe(false);
  });

  it("fails a circular reference", () => {
    expect(fail(gradeTask(payroll, { C2: "=C2*0.784" }, "en"))).toMatchObject({ reason: "excel_error", errorType: "CYCLE" });
  });

  it("ignores spaces before =, lower-case keys and inputs for locked cells", () => {
    expect(gradeTask(payroll, { c2: "  =B2*0.784" }, "en").pass).toBe(true);
    expect(gradeTask(payroll, { B2: "1", C2: "=B2*0.784" }, "en").pass).toBe(true);
  });
});

describe("gradeTask: ROUND lessons", () => {
  it("fails a missing ROUND: =A2/1.18*0.18", () => {
    expect(gradeTask(vatRound, { B2: "=A2/1.18*0.18" }, "en").pass).toBe(false);
  });

  it("passes any correctly rounded formula", () => {
    expect(gradeTask(vatRound, { B2: "=ROUND(A2*0.18/1.18,2)" }, "en").pass).toBe(true);
    expect(gradeTask(vatRound, { B2: "=ROUND(A2*18/118;2)" }, "ka").pass).toBe(true);
  });

  it("fails a rounded constant multiplier", () => {
    expect(gradeTask(vatRound, { B2: "=ROUND(A2*0.1525,2)" }, "en").pass).toBe(false);
  });
});

describe("gradeTask: ranges (fill down)", () => {
  const correct = {
    C2: "=B2*(1+$E$1)",
    C3: "=B3*(1+$E$1)",
    C4: "=B4*(1+$E$1)",
    C5: "=B5*(1+$E$1)",
    C6: "=B6*(1+$E$1)",
  };

  it("passes a correctly filled column", () => {
    const result = gradeTask(vatColumn, correct, "en");
    expect(result.pass).toBe(true);
    expect(result.cells.map((c) => c.cell)).toEqual(["C2", "C3", "C4", "C5", "C6"]);
  });

  it("reports the first failing cell, and every cell's verdict", () => {
    const { C5: _skipped, ...withoutC5 } = correct;
    const result = fail(gradeTask(vatColumn, { ...withoutC5, C4: "=B4*1.2" }, "en"));
    expect(result).toMatchObject({ cell: "C4", reason: "wrong_value" });
    expect(result.cells.map((c) => [c.cell, c.pass, c.reason])).toEqual([
      ["C2", true, undefined],
      ["C3", true, undefined],
      ["C4", false, "wrong_value"],
      ["C5", false, "empty"],
      ["C6", true, undefined],
    ]);
  });

  it("fails a relative reference that should be absolute ($E$1 filled down as E1, E2…)", () => {
    const session = new SheetSession(vatColumn.sheet);
    session.setCellInput("C2", "=B2*(1+E1)");
    session.fillDown("C2:C6");
    expect(session.getRawInput("C3")).toBe("=B3*(1+E2)");
    expect(fail(gradeTask(vatColumn, session.getLearnerInputs(), "en"))).toMatchObject({ reason: "wrong_value", cell: "C3" });
    session.destroy();
  });

  it("passes the same task done in the simulator with $E$1, in Georgian mode", () => {
    const session = new SheetSession(vatColumn.sheet, { mode: "ka" });
    session.setCellInput("C2", "=B2*(1+$E$1)");
    session.fillDown("C2:C6");
    expect(gradeTask(vatColumn, session.getLearnerInputs(), session.mode).pass).toBe(true);
    session.destroy();
  });

  it("fills the solution across the range with relative references adjusted", () => {
    expect(solutionFormulas(vatColumn)).toEqual(correct);
  });

  it("does not accept values spilled from one array formula as formulas in every cell", () => {
    const result = fail(gradeTask(vatColumn, { C2: "=B2:B6*(1+$E$1)" }, "en"));
    expect(result).toMatchObject({ reason: "not_formula", cell: "C3" });
    expect(result.messageKa).toContain("ავტომატურად გადმოიტანა");
  });
});

describe("gradeTask: lookups, mustUse and cases", () => {
  it("passes =VLOOKUP(A2,$D$2:$E$5,2,FALSE) (bare FALSE works)", () => {
    expect(gradeTask(lookup, { B2: "=VLOOKUP(A2,$D$2:$E$5,2,FALSE)" }, "en").pass).toBe(true);
    expect(gradeTask(lookup, { B2: "=VLOOKUP(A2;$D$2:$E$5;2;FALSE)" }, "ka").pass).toBe(true);
  });

  it("passes XLOOKUP and INDEX together with MATCH", () => {
    expect(gradeTask(lookup, { B2: "=XLOOKUP(A2,D2:D5,E2:E5)" }, "en").pass).toBe(true);
    expect(gradeTask(lookup, { B2: "=INDEX($E$2:$E$5;MATCH(A2;$D$2:$D$5;0))" }, "ka").pass).toBe(true);
  });

  it("fails INDEX without MATCH when the alternative needs both", () => {
    const result = fail(gradeTask(lookup, { B2: "=INDEX($E$2:$E$5,4)" }, "en"));
    expect(result).toMatchObject({ reason: "missing_function", mustUse: lookup.check.mustUse });
    expect(result.messageKa).toContain("VLOOKUP, XLOOKUP ან INDEX და MATCH");
  });

  it('fails =VLOOKUP("A-103",…) when the cases change the key', () => {
    const result = fail(gradeTask(lookup, { B2: '=VLOOKUP("A-103",$D$2:$E$5,2,FALSE)' }, "en"));
    expect(result).toMatchObject({ reason: "hardcoded", caseIndex: 1, trialInputs: { A2: "A-101" } });
    expect(result.messageKa).toContain("ყველა შემთხვევაში");
  });

  it("fails a direct reference (=E5) and a lookup without FALSE on unsorted keys", () => {
    const anyLookup = { ...lookup, check: { ...lookup.check, mustUse: [] } };
    expect(gradeTask(anyLookup, { B2: "=E5" }, "en").pass).toBe(false);
    expect(gradeTask(lookup, { B2: "=VLOOKUP(A2,$D$2:$E$5,2)" }, "en").pass).toBe(false);
  });

  it("matches whole function names outside text: IF is not IFERROR or COUNTIF", () => {
    expect(satisfiesMustUse("=IFERROR(A1,0)+COUNTIF(A:A,1)", ["IF"], "en")).toBe(false);
    expect(satisfiesMustUse('="IF("&A1', ["IF"], "en")).toBe(false);
    expect(satisfiesMustUse("=if(A1>0,1,0)", ["IF"], "en")).toBe(true);
    expect(satisfiesMustUse("=A1", [], "en")).toBe(true);
    expect(describeMustUseKa(["ROUND"])).toBe("ROUND");
  });

  it("fails > instead of >= at the boundary case", () => {
    expect(gradeTask(bonus, { B2: "=IF(A2>=2000,A2*0.05,0)" }, "en").pass).toBe(true);
    expect(gradeTask(bonus, { B2: "=IF(A2<2000,0,A2*5%)" }, "en").pass).toBe(true);
    const result = fail(gradeTask(bonus, { B2: "=IF(A2>2000,A2*0.05,0)" }, "en"));
    expect(result).toMatchObject({ reason: "hardcoded", caseIndex: 1, expected: "100", actual: "0" });
  });

  it("applies each case on top of the lesson data, not on top of the previous one", () => {
    const task: GradableTask = {
      sheet: { cells: [["a", "b", "sum"], [1, 2, null]], editable: ["C2"] },
      check: { cells: "C2", solution: "=A2+B2", vary: [], cases: [{ A2: 10 }, { B2: 20 }] },
    };
    // Case 2 sees A2 = 1 (restored), not 10: this formula would fail if cases piled up.
    expect(gradeTask(task, { C2: "=IF(AND(A2=10,B2=20),999,A2+B2)" }, "en").pass).toBe(true);
    expect(fail(gradeTask(task, { C2: "=IF(A2=10,0,A2+B2)" }, "en"))).toMatchObject({ caseIndex: 1 });
  });
});

describe("gradeTask: text, dates and TODAY", () => {
  const names: GradableTask = {
    sheet: {
      cells: [
        ["სახელი", "გვარი", "სრული სახელი", "ხელფასი"],
        ["გიორგი", "ბერიძე", null, 1500],
      ],
      editable: ["C2"],
    },
    check: { cells: "C2", solution: '=A2&" "&B2&" ("&D2&")"', vary: ["D2"] },
  };

  it("compares text exactly after trimming", () => {
    expect(gradeTask(names, { C2: '=A2&" "&B2&" ("&D2&") "' }, "en").pass).toBe(true);
    expect(gradeTask(names, { C2: '=CONCAT(A2;" ";B2;" (";D2;")")' }, "ka").pass).toBe(true);
    expect(fail(gradeTask(names, { C2: '=A2&B2&" ("&D2&")"' }, "en")).reason).toBe("wrong_value");
  });

  it("catches hardcoded text through the varied number", () => {
    expect(fail(gradeTask(names, { C2: '="გიორგი ბერიძე (1500)"' }, "en")).reason).toBe("hardcoded");
  });

  it("varies dates and freezes TODAY to sheet.today", () => {
    const overdue: GradableTask = {
      sheet: {
        cells: [["ვადა", "დღეები"], [{ date: "2026-10-01" }, null]],
        editable: ["B2"],
        today: "2026-11-02",
      },
      check: { cells: "B2", solution: "=TODAY()-A2", vary: ["A2"] },
    };
    expect(gradeTask(overdue, { B2: "=TODAY()-A2" }, "en").pass).toBe(true);
    expect(gradeTask(overdue, { B2: '=DATEDIF(A2,TODAY(),"d")' }, "en").pass).toBe(true);
    expect(fail(gradeTask(overdue, { B2: "=32" }, "en")).reason).toBe("hardcoded");
    const trial = swapTrialInputs(overdue)[0]!.A2 as { date: string };
    expect(trial.date).not.toBe("2026-10-01");
  });
});

describe("swap trials", () => {
  it("are deterministic: same seed, same trials and verdicts", () => {
    expect(swapTrialInputs(payroll, { seed: 5 })).toEqual(swapTrialInputs(payroll, { seed: 5 }));
    expect(swapTrialInputs(payroll, { seed: 5 })).not.toEqual(swapTrialInputs(payroll, { seed: 6 }));
    expect(swapTrialInputs(vatColumn)).toHaveLength(3);
    expect(Object.keys(swapTrialInputs(vatColumn)[0]!)).toEqual(["B2", "B3", "B4", "B5", "B6", "E1"]);
    const a = gradeTask(payroll, { C2: "=2500*0.784" }, "en", { seed: 9 });
    expect(gradeTask(payroll, { C2: "=2500*0.784" }, "en", { seed: 9 })).toEqual(a);
  });

  it("never reuse the original value or 0, and keep the decimals", () => {
    for (const trial of swapTrialInputs(vatColumn, { trials: 30 })) {
      expect(trial.B2).not.toBe(1450);
      expect(Number.isInteger(trial.B2)).toBe(true);
      expect(String(trial.B6).split(".")[1]?.length ?? 0).toBeLessThanOrEqual(1);
      expect(trial.E1).not.toBe(0.18);
      expect(trial.E1).not.toBe(0);
    }
  });
});

describe("valuesMatch", () => {
  it("compares by type, numbers with a relative tolerance", () => {
    expect(valuesMatch(1960, 1960 * (1 + 1e-12))).toBe(true);
    expect(valuesMatch(1960, 1960.01)).toBe(false);
    expect(valuesMatch(1960, 1960.01, 0.005)).toBe(true);
    expect(valuesMatch(1435.504, 1435.5, undefined, 2)).toBe(true);
    expect(valuesMatch(0.225, 0.23, undefined, 2)).toBe(true);
    expect(valuesMatch(1435.504, 1435.6, undefined, 2)).toBe(false);
    expect(valuesMatch(1960, "1960")).toBe(false);
    expect(valuesMatch(" a ", "a")).toBe(true);
    expect(valuesMatch("A", "a")).toBe(false);
    expect(valuesMatch(true, true)).toBe(true);
    const na = { kind: "error", type: "NA", code: "#N/A", detail: "" } as const;
    const div = { kind: "error", type: "DIV_BY_ZERO", code: "#DIV/0!", detail: "" } as const;
    expect(valuesMatch(na, na)).toBe(true);
    expect(valuesMatch(na, div)).toBe(false);
    expect(valuesMatch(na, 0)).toBe(false);
  });
});
