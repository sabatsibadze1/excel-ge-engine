/**
 * English ↔ Georgian Excel format: suggestions that need HyperFormula's parser.
 * The pure converter lives in convert.ts and is re-exported here.
 */
import { isValidFormula } from "./hyperformula";
import { looksLikeNumber, otherMode } from "./convert";
import type { FormatMode } from "./types";

export * from "./convert";

const THOUSANDS_GROUPED_RE = /^[+-]?\d{1,3}(?:,\d{3})+$/;

/**
 * Suggests switching the Excel format when an input only makes sense in the other one:
 * - a formula HyperFormula cannot parse in `mode` but can in the other mode
 *   (`=ROUND(B2*0,784;2)` typed in English mode → `'ka'`), or
 * - a number written with the other mode's decimal separator (`2500,5` in English mode → `'ka'`).
 * Returns the suggested mode, or null when nothing is wrong (or it is ambiguous).
 */
export function suggestFormatSwitch(raw: string, mode: FormatMode): FormatMode | null {
  const text = raw.trim();
  const other = otherMode(mode);
  if (text === "") return null;
  if (text.startsWith("=")) {
    return !isValidFormula(text, mode) && isValidFormula(text, other) ? other : null;
  }
  if (looksLikeNumber(text, mode) || !looksLikeNumber(text, other)) return null;
  // "1,234" in English mode reads as one thousand two hundred thirty-four, not a Georgian decimal.
  if (mode === "en" && THOUSANDS_GROUPED_RE.test(text)) return null;
  return other;
}

/**
 * For the "Switch to Georgian format?" hint: true when the learner is in English mode and typed
 * something that only works in Georgian format, e.g. `=ROUND(B2*0,784;2)` or `2500,5`.
 */
export function detectLikelyKaFormat(raw: string, mode: FormatMode): boolean {
  return mode === "en" && suggestFormatSwitch(raw, mode) === "ka";
}

/** The hint text for {@link suggestFormatSwitch}: "Switch to Georgian/English format?". */
export const FORMAT_SWITCH_MESSAGES_KA: Readonly<Record<FormatMode, string>> = {
  ka: "როგორც ჩანს, წერტილმძიმეს (;) იყენებთ. გადავრთოთ ქართულ ფორმატზე?",
  en: "როგორც ჩანს, ინგლისური ფორმატით წერთ (არგუმენტებს შორის მძიმე, ათწილადში წერტილი). გადავრთოთ ინგლისურ ფორმატზე?",
};
