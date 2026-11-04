import { describe, expect, it } from "vitest";
import { functionNames, tokenizeFormula, type TokenType } from "./tokenizer";
import type { FormatMode } from "./types";

function types(formula: string, mode: FormatMode = "en"): string[] {
  return tokenizeFormula(formula, mode)
    .filter((t) => t.type !== "whitespace")
    .map((t) => `${t.type}:${t.text}`);
}

describe("tokenizeFormula", () => {
  const samples: [string, FormatMode][] = [
    ['=IF(A1>0,"a,b;c",ROUND(B2*0.784,2))', "en"],
    ['=IF(A1>0;"a,b;c";ROUND(B2*0,784;2))', "ka"],
    ["=SUM($A$1:B$2, A:A, 1:3)", "en"],
    ['="say ""hi"", ok"&A1', "en"],
    ["={1,2;3,4}", "en"],
    ["={1;2|3;4}", "ka"],
    ["='My sheet'!A1+Sheet2!B2", "en"],
    ["=ROUND(B2*0,784;2)", "en"],
    ['=unterminated("abc', "en"],
    ["=@#?!~ 1e3 1.5E+3 .5", "en"],
    ["", "en"],
    ["plain text, not a formula", "en"],
  ];

  it.each(samples)("is lossless for %s (%s)", (formula, mode) => {
    expect(tokenizeFormula(formula, mode).map((t) => t.text).join("")).toBe(formula);
  });

  it("records start and end positions", () => {
    const tokens = tokenizeFormula("=B2*0.784");
    expect(tokens.map((t) => [t.start, t.end])).toEqual([
      [0, 1],
      [1, 3],
      [3, 4],
      [4, 9],
    ]);
  });

  it("tells references, functions and names apart", () => {
    expect(types("=LOG10(A1)+SUM($A$1:B2)+TRUE+vat")).toEqual([
      "equals:=",
      "function:LOG10",
      "lparen:(",
      "cell:A1",
      "rparen:)",
      "operator:+",
      "function:SUM",
      "lparen:(",
      "cell:$A$1",
      "colon::",
      "cell:B2",
      "rparen:)",
      "operator:+",
      "name:TRUE",
      "operator:+",
      "name:vat",
    ]);
  });

  it("reads whole-column and whole-row ranges", () => {
    expect(types("=SUM(A:$C)+SUM($1:3)")).toEqual([
      "equals:=",
      "function:SUM",
      "lparen:(",
      "column:A",
      "colon::",
      "column:$C",
      "rparen:)",
      "operator:+",
      "function:SUM",
      "lparen:(",
      "row:$1",
      "colon::",
      "row:3",
      "rparen:)",
    ]);
  });

  it("uses the mode's separators and decimal point", () => {
    expect(types("=ROUND(B2*0.784,2)", "en")).toContain("number:0.784");
    expect(types("=ROUND(B2*0.784,2)", "en")).toContain("argSeparator:,");
    expect(types("=ROUND(B2*0,784;2)", "ka")).toContain("number:0,784");
    expect(types("=ROUND(B2*0,784;2)", "ka")).toContain("argSeparator:;");
    // A `;` typed in English format is not a separator there.
    expect(types("=ROUND(B2*0,784;2)", "en")).toContain("unknown:;");
  });

  it("knows array constant separators", () => {
    const en = tokenizeFormula("={1,2;3,4}", "en").map((t) => t.type);
    const ka = tokenizeFormula("={1;2|3;4}", "ka").map((t) => t.type);
    const expected: TokenType[] = [
      "equals",
      "lbrace",
      "number",
      "arrayColumnSeparator",
      "number",
      "arrayRowSeparator",
      "number",
      "arrayColumnSeparator",
      "number",
      "rbrace",
    ];
    expect(en).toEqual(expected);
    expect(ka).toEqual(expected);
  });

  it("keeps separators inside text as part of the string token", () => {
    expect(types('=A1&"x;y,z"')).toEqual(["equals:=", "cell:A1", "operator:&", 'string:"x;y,z"']);
  });

  it("reads error literals, sheet prefixes and two-character operators", () => {
    expect(types("=IFERROR(#N/A,0)<>#DIV/0!")).toContain("error:#N/A");
    expect(types("=IFERROR(#N/A,0)<>#DIV/0!")).toContain("operator:<>");
    expect(types("='My sheet'!A1")).toEqual(["equals:=", "sheet:'My sheet'!", "cell:A1"]);
  });

  it("treats 1E3 as a number, not 1 followed by cell E3", () => {
    expect(types("=1E3*2")).toEqual(["equals:=", "number:1E3", "operator:*", "number:2"]);
  });
});

describe("functionNames", () => {
  it("lists called functions in upper case, once each", () => {
    expect(functionNames("=iferror(vlookup(A2,$E$2:$F$6,2,FALSE),0)+VLOOKUP(1,A:B,2,0)")).toEqual([
      "IFERROR",
      "VLOOKUP",
    ]);
  });

  it("ignores function names inside text and without parentheses", () => {
    expect(functionNames('="VLOOKUP("&A1')).toEqual([]);
    expect(functionNames("=XLOOKUP+1")).toEqual([]);
  });

  it("strips Excel's _xlfn. prefix", () => {
    expect(functionNames("=_xlfn.XLOOKUP(A1,B:B,C:C)")).toEqual(["XLOOKUP"]);
  });

  it("works in Georgian format", () => {
    expect(functionNames("=INDEX($F$2:$F$6;MATCH(A2;$E$2:$E$6;0))", "ka")).toEqual(["INDEX", "MATCH"]);
  });
});
