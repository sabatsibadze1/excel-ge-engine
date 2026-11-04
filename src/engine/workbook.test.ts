import { describe, expect, it, vi } from "vitest";
import type { Sheet } from "./schema";
import { SheetSession } from "./workbook";

/** Payroll-style sheet: names, gross, net (to fill), a VAT-rate cell in E1. */
const sheet: Sheet = {
  cells: [
    ["თანამშრომელი", "Gross (₾)", "Net (₾)", "განაკვეთი", 0.18],
    ["გიორგი ბერიძე", 2500, null],
    ["ნინო ლომიძე", 1800, null],
    ["ლევან გელაშვილი", 3200, null],
    ["სულ", "=SUM(B2:B4)", null],
  ],
  editable: ["C2:C4", "D2:D4"],
};

describe("SheetSession: reading", () => {
  it("loads the lesson data and computes pre-filled formulas", () => {
    const s = new SheetSession(sheet);
    expect(s.getValue("A2")).toBe("გიორგი ბერიძე");
    expect(s.getValue("B5")).toBe(7500);
    expect(s.getRawInput("B5")).toBe("=SUM(B2:B4)");
    expect(s.rowCount).toBe(5);
    expect(s.columnCount).toBe(5);
  });

  it("keeps lesson strings as text, even when they look like numbers, dates or booleans", () => {
    const s = new SheetSession({
      cells: [["01001012345", "TRUE", "15.01.2026", "10%", "=A1&\"\"", null]],
      editable: ["G1"],
    });
    expect(s.getValue("A1")).toBe("01001012345");
    expect(s.getValue("B1")).toBe("TRUE");
    expect(s.getValue("C1")).toBe("15.01.2026");
    expect(s.getValue("D1")).toBe("10%");
    expect(s.getValue("E1")).toBe("01001012345");
    expect(s.getValue("F1")).toBeNull();
    expect(s.getRawInput("A1")).toBe("'01001012345");
    expect(s.columnCount).toBe(7);
  });

  it("loads { date } cells as real dates shown as dd.mm.yyyy, and freezes TODAY", () => {
    const s = new SheetSession(
      { cells: [[{ date: "2026-02-01" }, "=TODAY()", "=B1-A1", 0.18]], editable: ["E1"], today: "2026-11-02" },
      { mode: "en" },
    );
    expect(s.getValue("A1")).toBe(46054);
    expect(s.getDisplayValue("A1")).toBe("01.02.2026");
    expect(s.getCell("A1").numberFormat).toBe("date");
    expect(s.getDisplayValue("B1")).toBe("02.11.2026");
    expect(s.getValue("C1")).toBe(274);
    s.setCellInput("E1", "=RANK.EQ(D1,D1:D1)");
    expect(s.getValue("E1")).toBe(1);
  });

  it("reports editable cells", () => {
    const s = new SheetSession(sheet);
    expect(s.isEditable("C3")).toBe(true);
    expect(s.isEditable("B2")).toBe(false);
    expect(s.editableCells()).toEqual(["C2", "C3", "C4", "D2", "D3", "D4"]);
  });

  it("gives the grid a full view of a cell", () => {
    const s = new SheetSession(sheet, { mode: "ka" });
    s.setCellInput("C2", "=B2/0");
    const view = s.getCell("C2");
    expect(view).toMatchObject({ address: "C2", row: 1, col: 2, display: "#DIV/0!", isFormula: true, isEditable: true, align: "center" });
    expect(view.error?.messageKa).toContain("ნულზე გაყოფა");
    expect(s.getCell("B2")).toMatchObject({ display: "2500", align: "right", isEditable: false, error: null });
  });
});

describe("SheetSession: typing in each format", () => {
  it("reads 2500,5 as a number in Georgian mode and shows a decimal comma", () => {
    const s = new SheetSession(sheet, { mode: "ka" });
    s.setCellInput("D2", "2500,5");
    expect(s.getValue("D2")).toBe(2500.5);
    expect(s.getDisplayValue("D2")).toBe("2500,5");
    expect(s.getRawInput("D2")).toBe("2500,5");
  });

  it("keeps 2500,5 as text in English mode", () => {
    const s = new SheetSession(sheet, { mode: "en" });
    s.setCellInput("D2", "2500,5");
    expect(s.getValue("D2")).toBe("2500,5");
  });

  it("computes Georgian-format formulas", () => {
    const s = new SheetSession(sheet, { mode: "ka" });
    s.setCellInput("C2", "=ROUND(B2*0,784;2)");
    expect(s.getValue("C2")).toBe(1960);
    s.setCellInput("C3", "=B3*0,784");
    expect(s.getDisplayValue("C3")).toBe("1411,2");
  });

  it("accepts bare TRUE and FALSE like Excel", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("D2", "=IF(B2>2000,TRUE,FALSE)");
    expect(s.getValue("D2")).toBe(true);
  });

  it("treats spaces before = as a formula and empty input as clearing", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C2", "  =B2*2");
    expect(s.getValue("C2")).toBe(5000);
    s.setCellInput("C2", "   ");
    expect(s.getValue("C2")).toBeNull();
  });

  it("refuses to edit locked cells", () => {
    const s = new SheetSession(sheet);
    const result = s.setCellInput("B2", "1");
    expect(result).toMatchObject({ ok: false, reason: "locked", cells: ["B2"] });
    expect(s.getValue("B2")).toBe(2500);
  });

  it("shows percents and dates the way they were typed", () => {
    const s = new SheetSession(sheet, { mode: "ka" });
    s.setCellInput("D2", "12,5%");
    s.setCellInput("D3", "15.01.2026");
    expect(s.getValue("D2")).toBe(0.125);
    expect(s.getDisplayValue("D2")).toBe("12,5%");
    expect(s.getRawInput("D3")).toBe("15.01.2026");
    expect(s.getCell("D3").numberFormat).toBe("date");
  });
});

describe("SheetSession: undo, copy/paste, fill", () => {
  it("undoes and redoes edits", () => {
    const s = new SheetSession(sheet);
    expect(s.canUndo()).toBe(false);
    s.setCellInput("C2", "=B2*0.784");
    s.setCellInput("C2", "=B2*2");
    expect(s.undo()).toBe(true);
    expect(s.getRawInput("C2")).toBe("=B2*0.784");
    expect(s.undo()).toBe(true);
    expect(s.getValue("C2")).toBeNull();
    expect(s.undo()).toBe(false);
    expect(s.redo()).toBe(true);
    expect(s.getRawInput("C2")).toBe("=B2*0.784");
  });

  it("pastes with relative references adjusted and absolute ones kept", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C2", "=B2*(1+$E$1)");
    s.copy("C2");
    expect(s.paste("C3:C4")).toEqual({ ok: true, changed: ["C3", "C4"] });
    expect(s.getRawInput("C3")).toBe("=B3*(1+$E$1)");
    expect(s.getRawInput("C4")).toBe("=B4*(1+$E$1)");
    expect(s.getValue("C4")).toBeCloseTo(3776);
    // The whole paste is one undo step.
    s.undo();
    expect(s.getValue("C3")).toBeNull();
    expect(s.getValue("C4")).toBeNull();
  });

  it("pastes into the next column too", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C2", "=B2*2");
    s.copy("C2");
    s.paste("D2");
    expect(s.getRawInput("D2")).toBe("=C2*2");
  });

  it("refuses a paste that would touch a locked cell", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C2", "=B2*2");
    s.copy("C2");
    const result = s.paste("C4:C5");
    expect(result).toMatchObject({ ok: false, reason: "locked", cells: ["C5"] });
    expect(s.getValue("C4")).toBeNull();
  });

  it("can paste the same copy twice, and says when nothing was copied", () => {
    const s = new SheetSession(sheet);
    expect(s.paste("C2")).toMatchObject({ ok: false, reason: "empty_clipboard" });
    s.setCellInput("C2", "=B2+1");
    s.copy("C2");
    s.paste("C3");
    s.paste("C4");
    expect(s.getRawInput("C4")).toBe("=B4+1");
  });

  it("copies locked cells too and returns the shown values as text", () => {
    const s = new SheetSession(sheet, { mode: "ka" });
    s.setCellInput("D2", "2,5");
    expect(s.copy("B2:D2")).toBe("2500\t\t2,5");
  });

  it("fills down like Ctrl+D", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C2", "=B2*0.784");
    expect(s.fillDown("C2:C4")).toEqual({ ok: true, changed: ["C3", "C4"] });
    expect(s.getRawInput("C3")).toBe("=B3*0.784");
    expect(s.getRawInput("C4")).toBe("=B4*0.784");
    s.undo();
    expect(s.getValue("C3")).toBeNull();
    // On a single cell, Ctrl+D copies the cell above.
    s.fillDown("C3");
    expect(s.getRawInput("C3")).toBe("=B3*0.784");
  });

  it("refuses to fill into locked cells", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C4", "=B4*2");
    expect(s.fillDown("C4:C5")).toMatchObject({ ok: false, reason: "locked", cells: ["C5"] });
  });

  it("fills right like Ctrl+R", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C2", "=B2*2");
    s.fillRight("C2:D2");
    expect(s.getRawInput("D2")).toBe("=C2*2");
  });

  it("pastes plain text from outside the grid", () => {
    const s = new SheetSession(sheet, { mode: "ka" });
    expect(s.pasteText("C2", "=B2*2\t1,5\n=B3*2\t2,5\n")).toMatchObject({ ok: true });
    expect(s.getValue("C3")).toBe(3600);
    expect(s.getValue("D3")).toBe(2.5);
    s.pasteText("D2:D4", "7");
    expect([s.getValue("D2"), s.getValue("D3"), s.getValue("D4")]).toEqual([7, 7, 7]);
  });

  it("clears a range of editable cells", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C2", "1");
    s.setCellInput("C3", "2");
    expect(s.clear("C2:C3")).toMatchObject({ ok: true });
    expect(s.getValue("C2")).toBeNull();
    expect(s.clear("B2:C2")).toMatchObject({ ok: false, reason: "locked" });
  });
});

describe("SheetSession: switching the Excel format", () => {
  it("rewrites learner formulas without changing their meaning", () => {
    const s = new SheetSession(sheet, { mode: "en" });
    s.setCellInput("C2", '=IF(B2>2000,ROUND(B2*0.784,2),"x;y")');
    s.setCellInput("D2", "2500.5");
    const before = s.getValue("C2");
    s.setMode("ka");
    expect(s.mode).toBe("ka");
    expect(s.getRawInput("C2")).toBe('=IF(B2>2000;ROUND(B2*0,784;2);"x;y")');
    expect(s.getValue("C2")).toBe(before);
    expect(s.getRawInput("D2")).toBe("2500,5");
    expect(s.getValue("D2")).toBe(2500.5);
    expect(s.getRawInput("B5")).toBe("=SUM(B2:B4)");
    s.setMode("en");
    expect(s.getRawInput("C2")).toBe('=IF(B2>2000,ROUND(B2*0.784,2),"x;y")');
  });

  it("keeps a formula typed in the other format, so switching fixes it", () => {
    const s = new SheetSession(sheet, { mode: "en" });
    s.setCellInput("C2", "=ROUND(B2*0,784;2)");
    expect(s.getCell("C2").error?.type).toBe("ERROR");
    s.setMode("ka");
    expect(s.getRawInput("C2")).toBe("=ROUND(B2*0,784;2)");
    expect(s.getValue("C2")).toBe(1960);
  });

  it("clears undo history and the copy clipboard", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C2", "=B2");
    s.copy("C2");
    s.setMode("ka");
    expect(s.canUndo()).toBe(false);
    expect(s.paste("C3")).toMatchObject({ ok: false, reason: "empty_clipboard" });
  });
});

describe("SheetSession: change subscription and learner inputs", () => {
  it("notifies subscribers after every change", () => {
    const s = new SheetSession(sheet);
    const listener = vi.fn();
    const unsubscribe = s.subscribe(listener);
    s.setCellInput("C2", "=B2");
    s.fillDown("C2:C3");
    s.undo();
    s.setMode("ka");
    expect(listener).toHaveBeenCalledTimes(4);
    expect(s.version).toBe(4);
    unsubscribe();
    s.setCellInput("C2", "1");
    expect(listener).toHaveBeenCalledTimes(4);
    // A refused edit changes nothing and notifies nobody.
    s.subscribe(listener);
    s.setCellInput("A1", "x");
    expect(listener).toHaveBeenCalledTimes(4);
  });

  it("returns the learner's inputs in their format, for grading", () => {
    const s = new SheetSession(sheet, { mode: "ka" });
    s.setCellInput("C2", "=B2*0,784");
    s.setCellInput("D3", "2500,5");
    expect(s.getLearnerInputs()).toEqual({ C2: "=B2*0,784", D3: "2500,5" });
  });

  it("does not count values spilled from another cell as inputs", () => {
    const s = new SheetSession(sheet);
    s.setCellInput("C2", "=B2:B4*0.784");
    expect(s.getValue("C3")).toBeCloseTo(1411.2);
    expect(s.getCell("C3").isSpilled).toBe(true);
    expect(s.getRawInput("C3")).toBe("");
    expect(s.getLearnerInputs()).toEqual({ C2: "=B2:B4*0.784" });
  });
});
