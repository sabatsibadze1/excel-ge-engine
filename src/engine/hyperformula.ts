/**
 * Internal: the only place that builds HyperFormula instances (`createEngine`). Every engine gets
 * the same configuration, the Excel-compatibility functions and the TRUE/FALSE names.
 *
 * HyperFormula is GPLv3 (see app/LICENSE), used with `licenseKey: 'gpl-v3'`.
 */
import {
  DetailedCellError,
  HyperFormula,
  type CellValue as HfCellValue,
  type ConfigParams,
  type RawCellContent,
  type SerializedNamedExpression,
  type SimpleCellAddress,
  type SimpleCellRange,
} from "hyperformula";
import { parseCellAddress, type CellCoords, type RangeCoords } from "./addresses";
import { registerExcelFunctions, type EngineContext } from "./excel-functions";
import {
  FORMAT_SEPARATORS,
  type CellValue,
  type ExcelErrorType,
  type FormatMode,
  type NumberFormatKind,
} from "./types";

export type { RawCellContent };

/** Sheet id of a single-sheet engine. */
export const SHEET_ID = 0;

/** Settings shared by both formats (technical review, verified). */
const BASE_CONFIG: Partial<ConfigParams> = {
  licenseKey: "gpl-v3",
  // Like Excel: =A1 on an empty cell gives 0, not an empty value.
  evaluateNullToZero: true,
  // Like Excel 365: =SUM(B2:B6*2) works and =B2:B6*2 spills into the cells below.
  useArrayArithmetic: true,
  ignoreWhiteSpace: "any",
  accentSensitive: true,
  localeLang: "ka-GE",
};

/** HyperFormula settings for a format mode. Never switch a live engine with `updateConfig()`. */
export function engineConfig(mode: FormatMode): Partial<ConfigParams> {
  const seps = FORMAT_SEPARATORS[mode];
  return {
    ...BASE_CONFIG,
    functionArgSeparator: seps.arg,
    decimalSeparator: seps.decimal,
    // Must stay empty: the converter relies on "," never being a thousands separator.
    thousandSeparator: "",
    arrayColumnSeparator: seps.arrayColumn,
    arrayRowSeparator: seps.arrayRow,
  };
}

/**
 * HyperFormula 3.5.0 only knows `TRUE()` / `FALSE()`. Excel also accepts bare `TRUE` and `FALSE`
 * (=VLOOKUP(A2,D:E,2,FALSE) would give #NAME?), so every engine defines them as names.
 */
const EXCEL_NAMES: SerializedNamedExpression[] = [
  { name: "TRUE", expression: "=TRUE()" },
  { name: "FALSE", expression: "=FALSE()" },
];

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);

/** "2026-11-02" → Excel date serial (46328). */
export function isoDateToSerial(iso: string): number {
  return (Date.parse(`${iso}T00:00:00Z`) - EXCEL_EPOCH_MS) / 86_400_000;
}

/** Excel date serial → "2026-11-02". */
export function serialToIsoDate(serial: number): string {
  return new Date(EXCEL_EPOCH_MS + Math.round(serial) * 86_400_000).toISOString().slice(0, 10);
}

export interface EngineOptions {
  /** The date TODAY() returns (a task's `sheet.today`). Default: the real date. */
  readonly todayIso?: string | null;
  /** Undo steps kept (only matters for the live grid). */
  readonly undoLimit?: number;
}

/**
 * Builds an engine. Pass one grid for a one-sheet engine (sheet id 0), or named grids
 * (`{ sol, lrn }`) for a multi-sheet engine. Cells must already be HyperFormula input
 * (see `lessonValueToRaw`).
 */
export function createEngine(
  data: RawCellContent[][] | Record<string, RawCellContent[][]>,
  mode: FormatMode,
  options: EngineOptions = {},
): HyperFormula {
  registerExcelFunctions();
  const context: EngineContext | undefined = options.todayIso
    ? { todaySerial: isoDateToSerial(options.todayIso) }
    : undefined;
  const config: Partial<ConfigParams> = {
    ...engineConfig(mode),
    context,
    ...(options.undoLimit !== undefined ? { undoLimit: options.undoLimit } : {}),
  };
  return Array.isArray(data)
    ? HyperFormula.buildFromArray(data, config, EXCEL_NAMES)
    : HyperFormula.buildFromSheets(data, config, EXCEL_NAMES);
}

export function hfAddress({ row, col }: CellCoords, sheet = SHEET_ID): SimpleCellAddress {
  return { sheet, row, col };
}

export function hfAddressOf(address: string, sheet = SHEET_ID): SimpleCellAddress {
  return hfAddress(parseCellAddress(address), sheet);
}

export function hfRange(range: RangeCoords, sheet = SHEET_ID): SimpleCellRange {
  return { start: hfAddress(range.start, sheet), end: hfAddress(range.end, sheet) };
}

/** Converts a HyperFormula value into the engine's public `CellValue`. */
export function toCellValue(value: HfCellValue): CellValue {
  if (value instanceof DetailedCellError) {
    return {
      kind: "error",
      type: value.type as ExcelErrorType,
      code: value.value,
      detail: value.message ?? "",
    };
  }
  return value ?? null;
}

/** The number format HyperFormula inferred for a cell, or null if the value is not a number. */
export function numberFormatAt(hf: HyperFormula, address: SimpleCellAddress): NumberFormatKind | null {
  switch (hf.getCellValueDetailedType(address)) {
    case "NUMBER_RAW":
      return "number";
    case "NUMBER_PERCENT":
      return "percent";
    case "NUMBER_DATE":
      return "date";
    case "NUMBER_TIME":
      return "time";
    case "NUMBER_DATETIME":
      return "datetime";
    case "NUMBER_CURRENCY":
      return "currency";
    default:
      return null;
  }
}

/** The currency symbol of a currency-formatted cell (`"$"`), if any. */
export function currencySymbolAt(hf: HyperFormula, address: SimpleCellAddress): string | undefined {
  const format = hf.getCellValueFormat(address);
  return typeof format === "string" && format.length > 0 ? format : undefined;
}

/** True for a cell filled by another cell's spilling array formula (not the formula cell itself). */
export function isSpilledCell(hf: HyperFormula, address: SimpleCellAddress): boolean {
  return hf.isCellPartOfArray(address) && !hf.doesCellHaveFormula(address);
}

const validators: Partial<Record<FormatMode, HyperFormula>> = {};

/** True if HyperFormula can parse `formula` (which must start with `=`) in the given format. */
export function isValidFormula(formula: string, mode: FormatMode): boolean {
  if (!formula.startsWith("=")) return false;
  validators[mode] ??= createEngine([], mode);
  return validators[mode].validateFormula(formula);
}

let supportedFunctions: ReadonlySet<string> | undefined;

/** Function names our engine supports (HyperFormula's plus the custom ones), upper case. */
export function supportedFunctionNames(): ReadonlySet<string> {
  registerExcelFunctions();
  supportedFunctions ??= new Set(HyperFormula.getRegisteredFunctionNames("enGB"));
  return supportedFunctions;
}
