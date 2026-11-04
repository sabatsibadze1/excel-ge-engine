import { describe, expect, it } from "vitest";
import {
  columnToLetters,
  expandRange,
  expandRanges,
  formatCellAddress,
  formatRange,
  lettersToColumn,
  parseCellAddress,
  parseRange,
  tryParseRange,
} from "./addresses";

describe("addresses", () => {
  it("converts column letters both ways", () => {
    expect([0, 25, 26, 701, 702, 16383].map(columnToLetters)).toEqual(["A", "Z", "AA", "ZZ", "AAA", "XFD"]);
    expect(["A", "z", "AA", "XFD"].map(lettersToColumn)).toEqual([0, 25, 26, 16383]);
  });

  it("parses and formats cells", () => {
    expect(parseCellAddress("C2")).toEqual({ row: 1, col: 2 });
    expect(parseCellAddress(" c2 ")).toEqual({ row: 1, col: 2 });
    expect(formatCellAddress(1, 2)).toBe("C2");
    expect(() => parseCellAddress("$C$2")).toThrow();
    expect(() => parseCellAddress("C0")).toThrow();
  });

  it("normalizes ranges given in any corner order", () => {
    expect(parseRange("C6:B2")).toEqual({ start: { row: 1, col: 1 }, end: { row: 5, col: 2 } });
    expect(formatRange(parseRange("C6:B2"))).toBe("B2:C6");
    expect(formatRange(parseRange("D4"))).toBe("D4");
    expect(tryParseRange("C2-C6")).toBeNull();
  });

  it("expands ranges in reading order", () => {
    expect(expandRange("B2:C3")).toEqual(["B2", "C2", "B3", "C3"]);
    expect(expandRanges(["B2:B3", "B3", "E1"])).toEqual(["B2", "B3", "E1"]);
  });
});
