/**
 * Shared public types of the spreadsheet engine.
 * Nothing here depends on HyperFormula, so UI code can import these freely.
 */

/**
 * The learner's Excel format ("Excel-format switch", build plan §6.2):
 * - `'en'`: English format, `=ROUND(B2*0.784, 2)` (`,` between arguments, `.` for decimals)
 * - `'ka'`: Georgian Windows format, `=ROUND(B2*0,784; 2)` (`;` between arguments, `,` for decimals)
 */
export type FormatMode = "en" | "ka";

export const FORMAT_MODES: readonly FormatMode[] = ["en", "ka"];

/** Separator characters per format mode (also used to configure HyperFormula). */
export interface FormatSeparators {
  /** Between function arguments: `,` (en) or `;` (ka). */
  readonly arg: string;
  /** Decimal point in numbers: `.` (en) or `,` (ka). */
  readonly decimal: "." | ",";
  /** Between columns of an array constant `{1,2}`: `,` (en) or `;` (ka). */
  readonly arrayColumn: "," | ";";
  /** Between rows of an array constant `{1;2}`: `;` (en) or `|` (ka, HyperFormula's convention). */
  readonly arrayRow: ";" | "|";
}

export const FORMAT_SEPARATORS: Readonly<Record<FormatMode, FormatSeparators>> = {
  en: { arg: ",", decimal: ".", arrayColumn: ",", arrayRow: ";" },
  ka: { arg: ";", decimal: ",", arrayColumn: ";", arrayRow: "|" },
};

/**
 * Excel error kinds, named like HyperFormula's `ErrorType`.
 * `CYCLE` (circular reference) and `ERROR` (formula could not be parsed) are HyperFormula-specific:
 * Excel shows a warning dialog instead.
 */
export type ExcelErrorType =
  | "DIV_BY_ZERO"
  | "NA"
  | "NAME"
  | "NUM"
  | "REF"
  | "VALUE"
  | "CYCLE"
  | "ERROR"
  | "SPILL"
  | "LIC";

/** An error value in a cell, e.g. `#DIV/0!`. */
export interface CellError {
  readonly kind: "error";
  readonly type: ExcelErrorType;
  /** What the cell shows, e.g. `"#DIV/0!"`. */
  readonly code: string;
  /** HyperFormula's technical message in English (may be empty). For logs, never for learners. */
  readonly detail: string;
}

/** The computed value of a cell. `null` means empty. */
export type CellValue = number | string | boolean | null | CellError;

/**
 * How a number should be shown. HyperFormula infers it, like Excel's automatic formats:
 * typing `10%` gives a percent, `=DATE(2026,1,15)` gives a date.
 */
export type NumberFormatKind = "number" | "percent" | "date" | "time" | "datetime" | "currency";

export function isCellError(value: unknown): value is CellError {
  return (
    typeof value === "object" &&
    value !== null &&
    (value as { kind?: unknown }).kind === "error"
  );
}
