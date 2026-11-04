/**
 * How values look in the grid (`formatValue`) and in the formula bar (`valueToInput`),
 * per format mode. In Georgian mode numbers use a decimal comma: 1960,5.
 */
import { FORMAT_SEPARATORS, type CellValue, type FormatMode, type NumberFormatKind } from "./types";

/** Excel day 0 in the 1900 date system as used by HyperFormula (correct from 1 March 1900). */
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

export interface DisplayFormat {
  /** The number format HyperFormula inferred (default "number"). */
  readonly kind?: NumberFormatKind | null;
  /** Currency symbol for `kind: "currency"`, e.g. "₾". */
  readonly currency?: string;
}

function withDecimal(text: string, mode: FormatMode): string {
  return mode === "ka" ? text.replace(".", FORMAT_SEPARATORS.ka.decimal) : text;
}

function trimTrailingZeros(text: string): string {
  return text.includes(".") ? text.replace(/\.?0+$/, "") : text;
}

/** 123456789012 → "1.23457E+11", like Excel's General format. */
function formatScientific(value: number): string {
  const [mantissa = "", exponent = "0"] = value.toExponential(5).split("e");
  const exp = Number(exponent);
  const sign = exp < 0 ? "-" : "+";
  return `${trimTrailingZeros(mantissa)}E${sign}${String(Math.abs(exp)).padStart(2, "0")}`;
}

/**
 * A number as Excel's "General" format shows it: at most 10 significant digits, no thousands
 * separators, E-notation for very large or tiny numbers. 1960 → "1960", 1/3 → "0.333333333",
 * 1960.5 → "1960,5" in Georgian mode.
 */
export function formatNumber(value: number, mode: FormatMode): string {
  if (!Number.isFinite(value)) return "#NUM!";
  if (value === 0) return "0";
  const abs = Math.abs(value);
  if (abs >= 1e11 || abs < 1e-9) return withDecimal(formatScientific(value), mode);
  const integerDigits = Math.max(1, Math.floor(Math.log10(abs)) + 1);
  const decimals = Math.max(0, 10 - integerDigits);
  return withDecimal(trimTrailingZeros(value.toFixed(decimals)), mode);
}

/** A number written so that typing it back gives the same value: full precision, mode's decimal. */
function plainNumber(value: number, mode: FormatMode): string {
  return withDecimal(String(value), mode);
}

interface DateTimeParts {
  year: number;
  month: number;
  day: number;
  hours: number;
  minutes: number;
  seconds: number;
}

function serialToParts(serial: number): DateTimeParts {
  const totalSeconds = Math.round(serial * 86_400);
  const days = Math.floor(totalSeconds / 86_400);
  const secondsOfDay = totalSeconds - days * 86_400;
  const date = new Date(EXCEL_EPOCH_MS + days * MS_PER_DAY);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hours: Math.floor(secondsOfDay / 3600),
    minutes: Math.floor((secondsOfDay % 3600) / 60),
    seconds: secondsOfDay % 60,
  };
}

const pad2 = (n: number) => String(n).padStart(2, "0");

/** Excel date serial → "15.01.2026" (dd.mm.yyyy, the Georgian convention, in both modes). */
export function formatDate(serial: number): string {
  const { year, month, day } = serialToParts(serial);
  return `${pad2(day)}.${pad2(month)}.${year}`;
}

/** Excel time fraction → "12:30" or "12:30:15". */
export function formatTime(serial: number): string {
  const { hours, minutes, seconds } = serialToParts(serial);
  const base = `${pad2(hours)}:${pad2(minutes)}`;
  return seconds === 0 ? base : `${base}:${pad2(seconds)}`;
}

function formatFormattedNumber(
  value: number,
  mode: FormatMode,
  format: DisplayFormat,
  number: (n: number) => string,
): string {
  switch (format.kind) {
    case "percent":
      return `${number(Number((value * 100).toPrecision(15)))}%`;
    case "date":
      return formatDate(value);
    case "time":
      return formatTime(value);
    case "datetime":
      return `${formatDate(value)} ${formatTime(value)}`;
    default:
      return number(value);
  }
}

/** What a cell shows in the grid, e.g. "1960,5" (ka), "18%", "15.01.2026", "#DIV/0!", "TRUE". */
export function formatValue(value: CellValue, mode: FormatMode, format: DisplayFormat = {}): string {
  if (value === null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") {
    if (format.kind === "currency" && format.currency) {
      const shown = formatNumber(value, mode);
      return format.currency === "$" ? `$${shown}` : `${shown} ${format.currency}`;
    }
    return formatFormattedNumber(value, mode, format, (n) => formatNumber(n, mode));
  }
  return value.code;
}

const ERROR_LITERAL_RE =
  /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|CYCLE!|ERROR!|LIC!)$/i;
const NUMBER_LIKE_RE = /^[\d\s.,/:+\-%$e]+$/i;

/**
 * True if typing `text` would NOT give back the same text (it would become a number, date, boolean,
 * error or formula). Such text needs a leading apostrophe in the formula bar, like in Excel.
 */
export function needsApostrophe(text: string): boolean {
  const trimmed = text.trim();
  if (trimmed === "") return false;
  if (text.startsWith("=") || text.startsWith("'")) return true;
  if (/^(?:TRUE|FALSE)$/i.test(trimmed) || ERROR_LITERAL_RE.test(trimmed)) return true;
  return NUMBER_LIKE_RE.test(trimmed) && /\d/.test(trimmed);
}

/**
 * The text to put in the formula bar for a (non-formula) value, so that pressing Enter on it
 * gives the same value back: "2500,5" (ka), "10%", "15.01.2026", "'01001012345".
 */
export function valueToInput(value: CellValue, mode: FormatMode, format: DisplayFormat = {}): string {
  if (value === null) return "";
  if (typeof value === "string") return needsApostrophe(value) ? `'${value}` : value;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") {
    if (format.kind === "currency" && format.currency) {
      const plain = plainNumber(value, mode);
      return format.currency === "$" ? `$${plain}` : `${plain}${format.currency}`;
    }
    return formatFormattedNumber(value, mode, format, (n) => plainNumber(n, mode));
  }
  return value.code;
}

/** Excel's default alignment: numbers right, text left, booleans and errors centred. */
export function alignmentOf(value: CellValue): "left" | "right" | "center" {
  if (typeof value === "number") return "right";
  if (typeof value === "string" || value === null) return "left";
  return "center";
}
