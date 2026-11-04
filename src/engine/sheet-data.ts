/**
 * Internal: turns a lesson's starting data into HyperFormula input for a format mode.
 */
import { expandRanges, formatCellAddress, parseRange } from "./addresses";
import { enToKa } from "./format";
import type { RawCellContent } from "./hyperformula";
import { isDateCell, type Sheet, type SheetCell } from "./schema";
import type { FormatMode } from "./types";

/**
 * One lesson cell → HyperFormula input (SCHEMA.md "Cell values and addresses"; the technical
 * review's `toRaw`):
 * - numbers stay numbers and null stays empty;
 * - `{ "date": "2026-02-01" }` becomes "01.02.2026", which HyperFormula reads as a real date;
 * - strings starting with "=" are English-format formulas, converted for Georgian mode;
 * - any other string is text: a leading apostrophe stops HyperFormula from turning
 *   "01001012345" into a number, "TRUE" into a boolean or "01.02.2026" into a date.
 */
export function lessonValueToRaw(value: SheetCell | undefined, mode: FormatMode): RawCellContent {
  if (value === undefined || value === null || typeof value === "number") return value ?? null;
  if (isDateCell(value)) {
    const [year, month, day] = value.date.split("-");
    return `${day}.${month}.${year}`;
  }
  if (value.startsWith("=")) return mode === "ka" ? enToKa(value) : value;
  return `'${value}`;
}

/**
 * A learner's typed text → HyperFormula input. Empty text clears the cell. Spaces before "=" are
 * ignored, so " =B2" is still a formula (Excel would store it as text).
 */
export function learnerInputToRaw(raw: string | undefined): RawCellContent {
  if (raw === undefined || raw.trim() === "") return null;
  const trimmed = raw.trimStart();
  return trimmed.startsWith("=") ? trimmed.trimEnd() : raw;
}

export interface SheetSize {
  readonly rows: number;
  readonly cols: number;
}

/** Rows × columns covering the data, the editable cells and any extra ranges. */
export function sheetSize(sheet: Sheet, extraRanges: readonly string[] = []): SheetSize {
  let rows = sheet.cells.length;
  let cols = Math.max(0, ...sheet.cells.map((row) => row.length));
  for (const range of [...sheet.editable, ...extraRanges]) {
    const { end } = parseRange(range);
    rows = Math.max(rows, end.row + 1);
    cols = Math.max(cols, end.col + 1);
  }
  return { rows, cols };
}

/**
 * The full grid for HyperFormula. `override(cell)` may return content for a cell instead of the
 * lesson data (return `undefined` to keep the lesson data).
 */
export function buildSheetData(
  sheet: Sheet,
  mode: FormatMode,
  size: SheetSize,
  override?: (cell: string) => RawCellContent | undefined,
): RawCellContent[][] {
  const data: RawCellContent[][] = [];
  for (let row = 0; row < size.rows; row++) {
    const line: RawCellContent[] = [];
    for (let col = 0; col < size.cols; col++) {
      const replaced = override?.(formatCellAddress(row, col));
      line.push(replaced !== undefined ? replaced : lessonValueToRaw(sheet.cells[row]?.[col], mode));
    }
    data.push(line);
  }
  return data;
}

/** The set of editable cells of a sheet ("C2:C6" expanded). */
export function editableCellSet(sheet: Sheet): Set<string> {
  return new Set(expandRanges(sheet.editable));
}

/** The lesson value at a cell address, or undefined outside the data. */
export function lessonValueAt(sheet: Sheet, row: number, col: number): SheetCell | undefined {
  return sheet.cells[row]?.[col];
}
