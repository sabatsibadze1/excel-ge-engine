/**
 * English ↔ Georgian Excel format (build plan §6.2): the converter. No HyperFormula, so it can run
 * outside the engine worker (the app renders lesson texts with it).
 *
 *   en: =ROUND(B2*0.784, 2)      `,` between arguments, `.` for decimals, arrays {1,2;3,4}
 *   ka: =ROUND(B2*0,784; 2)      `;` between arguments, `,` for decimals, arrays {1;2|3;4}
 *
 * `enToKa` / `kaToEn` are the technical review's verified, quote-aware converter
 * (review-tech/exp/engine.mjs): text in "quotes" and 'quoted sheet names' is copied unchanged.
 * Never a naive `;` → `,` replace, and never HyperFormula's `updateConfig()` on a live engine
 * (it turns formulas containing "," or "." into #ERROR!).
 */
import type { FormatMode } from "./types";

export function otherMode(mode: FormatMode): FormatMode {
  return mode === "en" ? "ka" : "en";
}

/** Characters that can be part of a name or reference (so a digit after them is not a number). */
const IDENT_CHAR_RE = /[A-Za-z0-9_$.\u0080-￿]/;
const isIdentChar = (ch: string | undefined) => ch !== undefined && IDENT_CHAR_RE.test(ch);
const NUMBER_AT_RE = /^(\d*\.\d+|\d+\.?)(?![A-Za-z_(])/;

type Step = { text: string; len: number };

/** Walks a formula, copying quoted text verbatim and letting `onChar` rewrite everything else. */
function walk(formula: string, onChar: (f: string, i: number, inArray: boolean) => Step): string {
  let out = "";
  let i = 0;
  let arrayDepth = 0;
  while (i < formula.length) {
    const ch = formula[i]!;
    if (ch === '"' || ch === "'") {
      let end = formula.indexOf(ch, i + 1);
      end = end < 0 ? formula.length : end;
      out += formula.slice(i, end + 1);
      i = end + 1;
      continue;
    }
    if (ch === "{") arrayDepth++;
    else if (ch === "}") arrayDepth--;
    const step = onChar(formula, i, arrayDepth > 0);
    out += step.text;
    i += step.len;
  }
  return out;
}

/** English-format formula → Georgian format. Non-formulas (no leading "=") are returned unchanged. */
export function enToKa(formula: string): string {
  if (!formula.startsWith("=")) return formula;
  return walk(formula, (f, i, inArray) => {
    const number = !isIdentChar(f[i - 1]) ? NUMBER_AT_RE.exec(f.slice(i)) : null;
    if (number) return { text: number[0].replace(".", ","), len: number[0].length };
    const ch = f[i]!;
    return { text: ch === "," ? ";" : ch === ";" && inArray ? "|" : ch, len: 1 };
  });
}

/** Georgian-format formula → English format. Non-formulas are returned unchanged. */
export function kaToEn(formula: string): string {
  if (!formula.startsWith("=")) return formula;
  return walk(formula, (f, i, inArray) => {
    const ch = f[i]!;
    return { text: ch === "," ? "." : ch === ";" ? "," : ch === "|" && inArray ? ";" : ch, len: 1 };
  });
}

/**
 * Converts a formula from one format to the other without changing its meaning.
 * Only formulas that are valid in `from` are guaranteed to keep their meaning.
 */
export function convertFormula(formula: string, from: FormatMode, to: FormatMode): string {
  if (from === to) return formula;
  return from === "en" ? enToKa(formula) : kaToEn(formula);
}

const NUMBER_INPUT_RE: Record<FormatMode, RegExp> = {
  en: /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?%?$/i,
  ka: /^[+-]?(?:\d+(?:,\d*)?|,\d+)(?:e[+-]?\d+)?%?$/i,
};

/** True if `raw` is a plain number in that format: "2500.5" (en), "2500,5" (ka), "-3", "10%". */
export function looksLikeNumber(raw: string, mode: FormatMode): boolean {
  return NUMBER_INPUT_RE[mode].test(raw.trim());
}

/**
 * Converts a raw cell input between formats: formulas with {@link convertFormula}, plain numbers by
 * swapping the decimal separator ("2500.5" ↔ "2500,5"). Other text is returned unchanged.
 */
export function convertInput(raw: string, from: FormatMode, to: FormatMode): string {
  if (from === to) return raw;
  const trimmed = raw.trimStart();
  if (trimmed.startsWith("=")) return convertFormula(trimmed, from, to);
  if (looksLikeNumber(raw, from)) return from === "en" ? raw.replace(".", ",") : raw.replace(",", ".");
  return raw;
}

const FUNCTION_CALL_RE = /^[A-Za-z_][A-Za-z0-9_.]*\(/;

/**
 * Lesson texts (theory, hints) write formulas in backticks in English format: `=ROUND(B2*0.784, 2)`.
 * This returns the text with those formulas in the learner's format. Backticked snippets without
 * "=" that look like a function call (`ROUND(რიცხვი, 2)`) are converted too; anything else is kept.
 */
export function convertFormulasInText(text: string, mode: FormatMode): string {
  if (mode === "en") return text;
  return text.replace(/`([^`]+)`/g, (whole, code: string) => {
    if (code.startsWith("=")) return `\`${enToKa(code)}\``;
    if (FUNCTION_CALL_RE.test(code)) return `\`${enToKa(`=${code}`).slice(1)}\``;
    return whole;
  });
}
