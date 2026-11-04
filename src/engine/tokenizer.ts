/**
 * A small, lossless tokenizer for Excel formulas in English or Georgian format.
 *
 * "Lossless" means that joining the `text` of all tokens gives back the original string,
 * character for character, even for invalid formulas (unknown characters become `unknown` tokens).
 * That makes it safe for rewriting parts of a formula (format conversion, F4 cycling)
 * without touching anything else, especially the contents of "text strings".
 *
 * It is not a parser: HyperFormula does the real parsing and evaluation.
 */
import { FORMAT_SEPARATORS, type FormatMode } from "./types";

export type TokenType =
  /** The leading `=` of a formula. */
  | "equals"
  | "whitespace"
  /** A text literal including its quotes: `"a;b"`. Doubled quotes `""` are escapes. */
  | "string"
  /** A number literal in the formula's format: `0.784` (en) or `0,784` (ka), `1e3`. */
  | "number"
  /** An error literal: `#N/A`, `#DIV/0!`… */
  | "error"
  /** A cell reference: `B2`, `$B$2`, `B$2`, `$B2`. */
  | "cell"
  /** One side of a whole-column range: the `A` and `$C` in `A:$C`. */
  | "column"
  /** One side of a whole-row range: the `1` and `$3` in `1:$3`. */
  | "row"
  /** The `:` range operator. */
  | "colon"
  /** A function name directly followed by `(`: `SUM`, `LOG10`, `STDEV.S`. */
  | "function"
  /** Any other identifier: `TRUE`, `FALSE`, named ranges, misspelled words. */
  | "name"
  /** A sheet prefix including the `!`: `Sheet1!` or `'My sheet'!`. */
  | "sheet"
  /** Between function arguments: `,` (en) or `;` (ka). */
  | "argSeparator"
  /** Between columns inside `{…}`: `,` (en) or `;` (ka). */
  | "arrayColumnSeparator"
  /** Between rows inside `{…}`: `;` (en) or `|` (ka). */
  | "arrayRowSeparator"
  /** `+ - * / ^ & = < > <= >= <> %` */
  | "operator"
  | "lparen"
  | "rparen"
  | "lbrace"
  | "rbrace"
  /** Anything else, kept verbatim (e.g. a `;` typed in English format). */
  | "unknown";

export interface Token {
  readonly type: TokenType;
  readonly text: string;
  /** Index of the first character in the formula. */
  readonly start: number;
  /** Index just after the last character. */
  readonly end: number;
}

const WHITESPACE_RE = /^\s+/;
const ERROR_LITERAL_RE =
  /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|CYCLE!|ERROR!|LIC!)/i;
const CELL_RE = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}/;
const COLUMN_RANGE_RE = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})/;
const ROW_RANGE_RE = /^(\$?[0-9]{1,7}):(\$?[0-9]{1,7})/;
const IDENTIFIER_RE = /^[A-Za-z_\\\p{L}][A-Za-z0-9_.\\?\p{L}\p{N}]*/u;
const IDENTIFIER_CHAR_RE = /[A-Za-z0-9_.\\?\p{L}\p{N}$]/u;
const NUMBER_RE: Record<FormatMode, RegExp> = {
  en: /^(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/,
  ka: /^(?:[0-9]+(?:,[0-9]*)?|,[0-9]+)(?:[eE][+-]?[0-9]+)?/,
};
const TWO_CHAR_OPERATORS = new Set(["<>", "<=", ">="]);
const ONE_CHAR_OPERATORS = new Set(["+", "-", "*", "/", "^", "&", "=", "<", ">", "%"]);

/** True if `ch` would continue a reference or name (so the token before it is not a reference). */
function continuesIdentifier(ch: string | undefined): boolean {
  return ch !== undefined && (IDENTIFIER_CHAR_RE.test(ch) || ch === "(" || ch === "!");
}

/** Reads a quoted run (`"…"` or `'…'`) where a doubled quote is an escape. Returns its end index. */
function readQuoted(input: string, start: number, quote: string): number {
  let i = start + 1;
  while (i < input.length) {
    if (input[i] === quote) {
      if (input[i + 1] === quote) {
        i += 2;
        continue;
      }
      return i + 1;
    }
    i++;
  }
  return input.length; // unterminated: runs to the end
}

/**
 * Splits a formula into tokens. `mode` decides which characters are separators and decimals.
 * Works on any string; a leading `=` becomes an `equals` token.
 */
export function tokenizeFormula(formula: string, mode: FormatMode = "en"): Token[] {
  const seps = FORMAT_SEPARATORS[mode];
  const tokens: Token[] = [];
  let braceDepth = 0;
  let i = 0;

  const push = (type: TokenType, end: number) => {
    tokens.push({ type, text: formula.slice(i, end), start: i, end });
    i = end;
  };

  if (formula.startsWith("=")) push("equals", 1);

  while (i < formula.length) {
    const rest = formula.slice(i);
    const ch = formula[i]!;

    const ws = WHITESPACE_RE.exec(rest);
    if (ws) {
      push("whitespace", i + ws[0].length);
      continue;
    }

    if (ch === '"') {
      push("string", readQuoted(formula, i, '"'));
      continue;
    }

    if (ch === "'") {
      const end = readQuoted(formula, i, "'");
      if (formula[end] === "!") push("sheet", end + 1);
      else push("unknown", end);
      continue;
    }

    if (ch === "#") {
      const err = ERROR_LITERAL_RE.exec(rest);
      push(err ? "error" : "unknown", i + (err ? err[0].length : 1));
      continue;
    }

    if (ch === "{") {
      braceDepth++;
      push("lbrace", i + 1);
      continue;
    }
    if (ch === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      push("rbrace", i + 1);
      continue;
    }
    if (ch === "(") {
      push("lparen", i + 1);
      continue;
    }
    if (ch === ")") {
      push("rparen", i + 1);
      continue;
    }

    // Whole-row range `1:3` / `$1:$3` (before numbers, which would eat the `1`).
    const rowRange = ROW_RANGE_RE.exec(rest);
    if (rowRange && !continuesIdentifier(formula[i + rowRange[0].length])) {
      const [, left, right] = rowRange;
      push("row", i + left!.length);
      push("colon", i + 1);
      push("row", i + right!.length);
      continue;
    }

    const number = NUMBER_RE[mode].exec(rest);
    if (number) {
      push("number", i + number[0].length);
      continue;
    }

    if (ch === seps.arg) {
      push(braceDepth > 0 ? "arrayColumnSeparator" : "argSeparator", i + 1);
      continue;
    }
    if (ch === seps.arrayRow && braceDepth > 0) {
      push("arrayRowSeparator", i + 1);
      continue;
    }

    // Whole-column range `A:C` / `$A:$C`.
    const columnRange = COLUMN_RANGE_RE.exec(rest);
    if (columnRange && !continuesIdentifier(formula[i + columnRange[0].length])) {
      const [, left, right] = columnRange;
      push("column", i + left!.length);
      push("colon", i + 1);
      push("column", i + right!.length);
      continue;
    }

    const cell = CELL_RE.exec(rest);
    if (cell && !continuesIdentifier(formula[i + cell[0].length])) {
      push("cell", i + cell[0].length);
      continue;
    }

    const identifier = IDENTIFIER_RE.exec(rest);
    if (identifier) {
      const end = i + identifier[0].length;
      if (formula[end] === "(") push("function", end);
      else if (formula[end] === "!") push("sheet", end + 1);
      else push("name", end);
      continue;
    }

    const two = rest.slice(0, 2);
    if (TWO_CHAR_OPERATORS.has(two)) {
      push("operator", i + 2);
      continue;
    }
    if (ONE_CHAR_OPERATORS.has(ch)) {
      push("operator", i + 1);
      continue;
    }
    if (ch === ":") {
      push("colon", i + 1);
      continue;
    }

    push("unknown", i + 1);
  }

  return tokens;
}

/** "_xlfn.XLOOKUP" / "xlookup" → "XLOOKUP". */
export function normalizeFunctionName(name: string): string {
  return name.toUpperCase().replace(/^(?:_XLFN\.|_XLWS\.)+/, "");
}

/**
 * Function names called in a formula (uppercase, unique, in order of appearance).
 * Text inside "string literals" is ignored, so `="VLOOKUP"` calls no function.
 */
export function functionNames(formula: string, mode: FormatMode = "en"): string[] {
  const names: string[] = [];
  for (const token of tokenizeFormula(formula, mode)) {
    if (token.type !== "function") continue;
    const name = normalizeFunctionName(token.text);
    if (!names.includes(name)) names.push(name);
  }
  return names;
}
