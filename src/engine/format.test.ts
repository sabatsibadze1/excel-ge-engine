import { describe, expect, it } from "vitest";
import {
  FORMAT_SWITCH_MESSAGES_KA,
  convertFormula,
  convertFormulasInText,
  convertInput,
  detectLikelyKaFormat,
  enToKa,
  kaToEn,
  looksLikeNumber,
  suggestFormatSwitch,
} from "./format";
import { createEngine, toCellValue } from "./hyperformula";
import type { CellValue } from "./types";

/** The technical review's 14 formulas (review-tech/exp/07-convert-test.mjs). */
const REVIEW_FORMULAS = [
  "=ROUND(A1*0.784, 2)",
  "=ROUND(A1*.784,2)",
  '=IF(A1>=1000,"ბონუსი, დიახ; კი","არა")',
  "=RANK.EQ(A1,$A$1:$A$3)",
  '=SUMIFS(A1:A3,C1:C3,"თბილისი",B1:B3,">=1")',
  '=XLOOKUP("ბათუმი",C1:C3,A1:A3,"ვერ მოიძებნა")',
  "=SUM({1.5,2;3,4})",
  '=A1&", "&B1',
  "=NETWORKDAYS(E1,E2)",
  '=IFERROR(A1/0,"შეცდომა ""x""")',
  "=PERCENTILE.INC(A1:A3,0.25)",
  "=LOG10(100)+ATAN2(1,1)",
  '=TEXT(A1,"#,##0.00")',
  '=COUNTIF(B1:B3,">1.5")',
];

const DATA = [
  [2500, 1.5, "'თბილისი", 3, 46054],
  [1200, 2.25, "'ბათუმი", 4, 46100],
  [800, 0.5, "'თბილისი", 5, 46200],
];

function evaluate(formula: string, mode: "en" | "ka"): CellValue | string {
  const hf = createEngine(DATA, mode);
  try {
    if (!hf.validateFormula(formula)) return "INVALID";
    const value = hf.calculateFormula(formula, 0);
    return Array.isArray(value) ? JSON.stringify(value) : toCellValue(value);
  } finally {
    hf.destroy();
  }
}

describe("enToKa / kaToEn (the verified quote-aware converter)", () => {
  it.each(REVIEW_FORMULAS)("round-trips %s", (formula) => {
    expect(kaToEn(enToKa(formula))).toBe(formula);
  });

  it("gives the same values in both formats for the review formulas", () => {
    for (const formula of REVIEW_FORMULAS) {
      const en = evaluate(formula, "en");
      const ka = evaluate(enToKa(formula), "ka");
      if (formula.includes('""x""')) {
        // Known limitation: HyperFormula cannot parse "" inside text (the checker rejects it).
        expect(en).toBe("INVALID");
      } else if (formula.includes('">1.5"')) {
        // Known limitation: text inside quotes is never converted, so a decimal in a criteria
        // string stays "1.5" and fails in Georgian format (the checker's Georgian self-test catches it).
        expect(isNaN(Number(ka))).toBe(true);
      } else {
        expect(ka, formula).toEqual(en);
      }
    }
  });

  it.each([
    ["=ROUND(B2*0.784,2)", "=ROUND(B2*0,784;2)"],
    ["=B2*0.98*0.8", "=B2*0,98*0,8"],
    ['=IF(A1>0,"a,b;c","")', '=IF(A1>0;"a,b;c";"")'],
    ["=VLOOKUP(A2,$E$2:$F$6,2,FALSE)", "=VLOOKUP(A2;$E$2:$F$6;2;FALSE)"],
    ["={1,2;3,4}", "={1;2|3;4}"],
    ["=STDEV.S(A1:A3)+LOG10(100)", "=STDEV.S(A1:A3)+LOG10(100)"],
    ["=SUM( B2 , B3 )", "=SUM( B2 ; B3 )"],
    ["='My, sheet'!A1+1.5", "='My, sheet'!A1+1,5"],
  ])("converts %s ↔ %s", (en, ka) => {
    expect(enToKa(en)).toBe(ka);
    expect(kaToEn(ka)).toBe(en);
    expect(convertFormula(en, "en", "ka")).toBe(ka);
    expect(convertFormula(ka, "ka", "en")).toBe(en);
  });

  it("never touches text in quotes, even with ; and , inside", () => {
    const formula = '=CONCATENATE("x;y", ",", ";", "1.5", A1)';
    expect(enToKa(formula)).toBe('=CONCATENATE("x;y"; ","; ";"; "1.5"; A1)');
    expect(kaToEn(enToKa(formula))).toBe(formula);
  });

  it("leaves non-formulas and same-mode conversions alone", () => {
    expect(enToKa("2500.5")).toBe("2500.5");
    expect(convertFormula("=SUM(A1,B1)", "en", "en")).toBe("=SUM(A1,B1)");
  });
});

describe("convertInput", () => {
  it("converts formulas and plain numbers, leaves text alone", () => {
    expect(convertInput("=SUM(A1,B1)", "en", "ka")).toBe("=SUM(A1;B1)");
    expect(convertInput("2500.5", "en", "ka")).toBe("2500,5");
    expect(convertInput("2500,5", "ka", "en")).toBe("2500.5");
    expect(convertInput("12.5%", "en", "ka")).toBe("12,5%");
    expect(convertInput("Tbilisi, Batumi", "en", "ka")).toBe("Tbilisi, Batumi");
  });
});

describe("convertFormulasInText (theory and hints)", () => {
  it("converts backticked formulas for Georgian mode only", () => {
    const text = "ჩაწერეთ `=ROUND(B2/1.18*0.18, 2)` ან გამოიყენეთ `ROUND(რიცხვი, 2)`; ტექსტი 1.5, `B2`.";
    expect(convertFormulasInText(text, "ka")).toBe(
      "ჩაწერეთ `=ROUND(B2/1,18*0,18; 2)` ან გამოიყენეთ `ROUND(რიცხვი; 2)`; ტექსტი 1.5, `B2`.",
    );
    expect(convertFormulasInText(text, "en")).toBe(text);
  });
});

describe("looksLikeNumber", () => {
  it("follows the mode's decimal separator", () => {
    expect(looksLikeNumber("2500,5", "ka")).toBe(true);
    expect(looksLikeNumber("2500,5", "en")).toBe(false);
    expect(looksLikeNumber("2500.5", "en")).toBe(true);
    expect(looksLikeNumber(" -3 ", "en")).toBe(true);
    expect(looksLikeNumber("10%", "ka")).toBe(true);
    expect(looksLikeNumber("abc", "en")).toBe(false);
  });
});

describe("detectLikelyKaFormat / suggestFormatSwitch", () => {
  it("suggests Georgian format for Georgian-style input typed in English mode", () => {
    expect(detectLikelyKaFormat("=ROUND(B2*0,784;2)", "en")).toBe(true);
    expect(detectLikelyKaFormat("=SUM(B2;B3)", "en")).toBe(true);
    expect(detectLikelyKaFormat("=B2*0,784", "en")).toBe(true);
    expect(detectLikelyKaFormat("2500,5", "en")).toBe(true);
    expect(FORMAT_SWITCH_MESSAGES_KA.ka).toBe("როგორც ჩანს, წერტილმძიმეს (;) იყენებთ. გადავრთოთ ქართულ ფორმატზე?");
  });

  it("stays quiet for valid English input, text and thousands groups", () => {
    expect(detectLikelyKaFormat("=ROUND(B2*0.784,2)", "en")).toBe(false);
    expect(detectLikelyKaFormat("=B2*2", "en")).toBe(false);
    expect(detectLikelyKaFormat('="a;b"', "en")).toBe(false);
    expect(detectLikelyKaFormat("1,234", "en")).toBe(false);
    expect(detectLikelyKaFormat("გიორგი", "en")).toBe(false);
    expect(detectLikelyKaFormat("", "en")).toBe(false);
  });

  it("only fires in English mode", () => {
    expect(detectLikelyKaFormat("=ROUND(B2*0,784;2)", "ka")).toBe(false);
  });

  it("suggests English format for English-style input typed in Georgian mode", () => {
    expect(suggestFormatSwitch("=ROUND(B2*0.784,2)", "ka")).toBe("en");
    expect(suggestFormatSwitch("2500.5", "ka")).toBe("en");
    expect(suggestFormatSwitch("=ROUND(B2*0,784;2)", "ka")).toBeNull();
  });
});
