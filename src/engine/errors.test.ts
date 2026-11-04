import { describe, expect, it } from "vitest";
import { ERROR_CODES, explainError, unsupportedFunctionTip } from "./errors";
import { createEngine, toCellValue } from "./hyperformula";
import { isCellError, type ExcelErrorType } from "./types";

const GEORGIAN = /[ა-ჺ]/;

describe("explainError", () => {
  it.each(Object.keys(ERROR_CODES) as ExcelErrorType[])("has a Georgian explanation for %s", (type) => {
    const explanation = explainError(type);
    expect(explanation.code).toBe(ERROR_CODES[type]);
    expect(explanation.messageKa).toMatch(GEORGIAN);
    expect(explanation.messageKa.length).toBeLessThan(200);
  });

  it("maps what HyperFormula actually produces", () => {
    const hf = createEngine(
      [
        ["'A", 10],
        ["'B", 20],
        [
          '=VLOOKUP("Z",A1:B2,2,FALSE)',
          '="abc"*2',
          "=VLOKUP(1)",
          "=B1/0",
          "=SQRT(-1)",
          "=A0",
          "=ROUND(B2*0,784;2)",
          "=H3+1",
        ],
      ],
      "en",
    );
    const types = Array.from({ length: 8 }, (_, col) => {
      const value = toCellValue(hf.getCellValue({ sheet: 0, row: 2, col }));
      return isCellError(value) ? value.type : value;
    });
    expect(types).toEqual(["NA", "VALUE", "NAME", "DIV_BY_ZERO", "NUM", "REF", "ERROR", "CYCLE"]);
    hf.destroy();
  });

  it("names the misspelled function", () => {
    const { messageKa } = explainError("NAME", "Function name VLOKUP not recognized.");
    expect(messageKa).toContain("„VLOKUP“");
  });

  it("explains unquoted text", () => {
    const { messageKa } = explainError("NAME", "Named expression კი not recognized.");
    expect(messageKa).toContain('"კი"');
  });

  it("explains a wrong number of arguments instead of a failed lookup", () => {
    expect(explainError("NA", "Wrong number of arguments.").messageKa).toContain("არგუმენტების");
    expect(explainError("NA", "Value not found.").messageKa).toContain("ვერ მოიძებნა");
  });

  it("explains a range where one cell is expected", () => {
    expect(explainError("VALUE", "Cell range not allowed.").messageKa).toContain("დიაპაზონი");
  });

  it("says that Excel functions our engine lacks are not available yet", () => {
    const { messageKa } = explainError("NAME", "Function name RANK.AVG not recognized.", "ka");
    expect(messageKa).toContain("RANK.AVG Excel-ის ფუნქციაა");
    expect(messageKa).toContain("=RANK.EQ(B2;$B$2:$B$10)");
  });

  it("uses the reviewed Georgian wording", () => {
    expect(explainError("NA").messageKa).toContain("ზედმეტი ჰარი (Space)");
    expect(explainError("DIV_BY_ZERO").messageKa).toBe("ნულზე გაყოფა. შეამოწმეთ გამყოფი უჯრა: ცარიელი ან 0 ხომ არ არის?");
    expect(explainError("ERROR").messageKa).toContain("„;“ და „,“");
  });
});

describe("unsupportedFunctionTip", () => {
  it("shows the example formula in the learner's format", () => {
    expect(unsupportedFunctionTip("rank.avg", "en")).toContain("=RANK.EQ(B2,$B$2:$B$10)");
    expect(unsupportedFunctionTip("XMATCH")).toContain("MATCH");
    expect(unsupportedFunctionTip("RANK")).toBeUndefined();
    expect(unsupportedFunctionTip("SUM")).toBeUndefined();
  });
});
