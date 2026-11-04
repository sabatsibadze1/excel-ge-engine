/**
 * `SheetSession`: the live spreadsheet behind one task's grid. It wraps one HyperFormula instance
 * configured for the learner's format mode and enforces locked cells.
 *
 *   const session = new SheetSession(lesson.sheet, { mode: "ka" });
 *   session.setCellInput("C2", "=B2*0,784");
 *   session.getDisplayValue("C2"); // "1960"
 *   const result = gradeTask(lesson, session.getLearnerInputs(), session.mode);
 *
 * Every edit (input, paste, fill, clear) is one undo step. Subscribe to re-render:
 *   useSyncExternalStore(session.subscribe, () => session.version)
 */
import type { HyperFormula } from "hyperformula";
import {
  expandRange,
  expandRanges,
  formatCellAddress,
  parseCellAddress,
  parseRange,
  rangeContains,
  rangeSize,
  type CellCoords,
  type RangeCoords,
} from "./addresses";
import { alignmentOf, formatValue, valueToInput, type DisplayFormat } from "./display";
import { explainError, type ErrorExplanation } from "./errors";
import { convertFormula } from "./format";
import {
  createEngine,
  currencySymbolAt,
  hfAddress,
  hfRange,
  isSpilledCell,
  isValidFormula,
  numberFormatAt,
  toCellValue,
  type RawCellContent,
} from "./hyperformula";
import type { Sheet } from "./schema";
import { buildSheetData, learnerInputToRaw, sheetSize, type SheetSize } from "./sheet-data";
import { isCellError, type CellValue, type FormatMode, type NumberFormatKind } from "./types";

export interface SheetSessionOptions {
  /** The learner's Excel format. Default "en". */
  readonly mode?: FormatMode;
}

/** Result of an edit. Edits touching a locked cell are refused as a whole (like a protected sheet). */
export type EditResult =
  | { readonly ok: true; readonly changed: string[] }
  | {
      readonly ok: false;
      readonly reason: "locked" | "empty_clipboard";
      /** The locked cells that blocked the edit. */
      readonly cells: string[];
      readonly messageKa: string;
    };

/** Everything the grid needs to draw one cell. */
export interface CellView {
  readonly address: string;
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
  /** What the cell shows, formatted for the mode: "1960,5", "18%", "#N/A". */
  readonly display: string;
  /** What the formula bar shows: the formula in the learner's format, or the typed value. */
  readonly input: string;
  readonly isFormula: boolean;
  readonly isEditable: boolean;
  /** Filled by another cell's spilling array formula (like Excel 365 dynamic arrays). */
  readonly isSpilled: boolean;
  /** Number format inferred by HyperFormula, or null when the value is not a number. */
  readonly numberFormat: NumberFormatKind | null;
  /** Georgian explanation when the cell shows an error. */
  readonly error: ErrorExplanation | null;
  /** Excel's default alignment for the value. */
  readonly align: "left" | "right" | "center";
}

interface Clipboard {
  readonly range: RangeCoords;
  readonly data: RawCellContent[][];
}

/** Undo steps kept in the live grid. */
const UNDO_LIMIT = 100;

const LOCKED_ONE_KA = "ეს უჯრა ჩაკეტილია, მისი შეცვლა არ შეიძლება.";
const LOCKED_MANY_KA = "მონიშნულ დიაპაზონში ჩაკეტილი უჯრებიცაა, ამიტომ ცვლილება ვერ შესრულდა.";
const EMPTY_CLIPBOARD_KA = "ჯერ დააკოპირეთ უჯრები (Ctrl+C).";

/** Puts a block of content at `origin` inside an otherwise empty grid. */
function placeAt(origin: CellCoords, block: RawCellContent[][]): RawCellContent[][] {
  const grid: RawCellContent[][] = Array.from({ length: origin.row }, () => []);
  for (const line of block) grid.push([...Array<RawCellContent>(origin.col).fill(null), ...line]);
  return grid;
}

function blockAt(start: CellCoords, rows: number, cols: number): RangeCoords {
  return { start, end: { row: start.row + rows - 1, col: start.col + cols - 1 } };
}

/**
 * A formula typed in mode `from`, rewritten for mode `to`. A formula that HyperFormula cannot
 * parse in `from` but can in `to` was typed in the other format: it is kept as typed, so the
 * "Switch to Georgian format?" hint fixes it.
 */
function convertLearnerFormula(formula: string, from: FormatMode, to: FormatMode): string {
  if (!isValidFormula(formula, from) && isValidFormula(formula, to)) return formula;
  return convertFormula(formula, from, to);
}

export class SheetSession {
  /** The lesson sheet this session started from (never modified). */
  readonly sheet: Sheet;
  private hf: HyperFormula;
  private currentMode: FormatMode;
  private readonly size: SheetSize;
  private readonly editableRanges: readonly RangeCoords[];
  private readonly editableList: readonly string[];
  private clipboard: Clipboard | null = null;
  private readonly listeners = new Set<() => void>();
  private changeCount = 0;

  constructor(sheet: Sheet, options: SheetSessionOptions = {}) {
    this.sheet = sheet;
    this.currentMode = options.mode ?? "en";
    this.size = sheetSize(sheet);
    this.editableRanges = sheet.editable.map((range) => parseRange(range));
    this.editableList = expandRanges(sheet.editable);
    this.hf = this.newEngine(buildSheetData(sheet, this.currentMode, this.size), this.currentMode);
  }

  // ---------------------------------------------------------------- state

  /** The learner's current Excel format. */
  get mode(): FormatMode {
    return this.currentMode;
  }

  /** Increases after every change. Handy for `useSyncExternalStore`. */
  get version(): number {
    return this.changeCount;
  }

  /** Number of rows in the grid (covers the data and every editable cell). */
  get rowCount(): number {
    return this.size.rows;
  }

  /** Number of columns in the grid. */
  get columnCount(): number {
    return this.size.cols;
  }

  /**
   * Calls `listener` after every change (edit, undo/redo, paste, fill, mode switch).
   * Returns the unsubscribe function. Arrow function, so it can be passed around unbound.
   */
  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  };

  /** Frees the HyperFormula instance. The session must not be used afterwards. */
  destroy(): void {
    this.listeners.clear();
    this.hf.destroy();
  }

  // ---------------------------------------------------------------- reading

  isEditable(address: string): boolean {
    return this.isEditableAt(parseCellAddress(address));
  }

  /** All editable cells, in reading order. */
  editableCells(): string[] {
    return [...this.editableList];
  }

  /** The computed value: number, text, boolean, null (empty) or a `CellError`. */
  getValue(address: string): CellValue {
    return toCellValue(this.hf.getCellValue(hfAddress(parseCellAddress(address))));
  }

  /** The value as the cell shows it, formatted for the mode ("1960,5" in Georgian mode). */
  getDisplayValue(address: string): string {
    const coords = parseCellAddress(address);
    return formatValue(this.valueAt(coords), this.currentMode, this.formatAt(coords));
  }

  /**
   * What the formula bar shows: the formula in the learner's format (`=ROUND(B2*0,784;2)`),
   * or the value as it would be typed ("2500,5", "10%", "'01001012345"). Empty for empty cells.
   */
  getRawInput(address: string): string {
    return this.rawInputAt(parseCellAddress(address));
  }

  isFormula(address: string): boolean {
    return this.hf.doesCellHaveFormula(hfAddress(parseCellAddress(address)));
  }

  /** Georgian explanation if the cell shows an error, else null. */
  getError(address: string): ErrorExplanation | null {
    const value = this.getValue(address);
    return isCellError(value) ? explainError(value.type, value.detail, this.currentMode) : null;
  }

  /** Everything the grid needs for one cell. */
  getCell(address: string): CellView {
    const coords = parseCellAddress(address);
    const at = hfAddress(coords);
    const value = this.valueAt(coords);
    const format = this.formatAt(coords);
    return {
      address: formatCellAddress(coords.row, coords.col),
      row: coords.row,
      col: coords.col,
      value,
      display: formatValue(value, this.currentMode, format),
      input: this.rawInputAt(coords),
      isFormula: this.hf.doesCellHaveFormula(at),
      isEditable: this.isEditableAt(coords),
      isSpilled: isSpilledCell(this.hf, at),
      numberFormat: format.kind ?? null,
      error: isCellError(value) ? explainError(value.type, value.detail, this.currentMode) : null,
      align: alignmentOf(value),
    };
  }

  /**
   * The learner's inputs for grading: raw text of every non-empty editable cell, in the current
   * mode. Pass to `gradeTask(task, inputs, session.mode)`.
   */
  getLearnerInputs(): Record<string, string> {
    const inputs: Record<string, string> = {};
    for (const cell of this.editableList) {
      const raw = this.rawInputAt(parseCellAddress(cell));
      if (raw !== "") inputs[cell] = raw;
    }
    return inputs;
  }

  // ---------------------------------------------------------------- editing

  /**
   * Enters what the learner typed, in their format: `=B2*0,784` and `2500,5` in Georgian mode,
   * `=B2*0.784` and `2500.5` in English mode. Empty text clears the cell.
   */
  setCellInput(address: string, raw: string): EditResult {
    const coords = parseCellAddress(address);
    const cell = formatCellAddress(coords.row, coords.col);
    if (!this.isEditableAt(coords)) return this.refuseLocked([cell]);
    this.clipboard = null;
    this.hf.setCellContents(hfAddress(coords), learnerInputToRaw(raw));
    this.emitChange();
    return { ok: true, changed: [cell] };
  }

  /** Delete key: empties every cell of the range (all of them must be editable). */
  clear(range: string): EditResult {
    const target = parseRange(range);
    const refused = this.checkEditable(target);
    if (refused) return refused;
    const { rows, cols } = rangeSize(target);
    this.clipboard = null;
    this.hf.setCellContents(
      hfAddress(target.start),
      Array.from({ length: rows }, () => Array<RawCellContent>(cols).fill(null)),
    );
    this.emitChange();
    return { ok: true, changed: expandRange(target) };
  }

  /**
   * Ctrl+C: remembers the range for `paste`. Locked cells may be copied.
   * Returns the shown values as tab-separated text, for the system clipboard.
   */
  copy(range: string): string {
    const source = parseRange(range);
    this.clipboard = { range: source, data: this.hf.getRangeSerialized(hfRange(source)) };
    const lines: string[] = [];
    for (let row = source.start.row; row <= source.end.row; row++) {
      const values: string[] = [];
      for (let col = source.start.col; col <= source.end.col; col++) {
        values.push(formatValue(this.valueAt({ row, col }), this.currentMode, this.formatAt({ row, col })));
      }
      lines.push(values.join("\t"));
    }
    return lines.join("\n");
  }

  /**
   * Ctrl+V: pastes the copied cells with Excel-style relative-reference adjustment
   * (`=B2*$E$1` copied one row down becomes `=B3*$E$1`). `target` is a cell (the top-left corner)
   * or a range; a range that is a multiple of the copied block is filled by repeating it.
   */
  paste(target: string): EditResult {
    const clip = this.clipboard;
    if (!clip) return { ok: false, reason: "empty_clipboard", cells: [], messageKa: EMPTY_CLIPBOARD_KA };
    const size = rangeSize(clip.range);
    const selected = parseRange(target);
    const selectedSize = rangeSize(selected);
    const tiles = selectedSize.rows % size.rows === 0 && selectedSize.cols % size.cols === 0;
    const dest = tiles ? selected : blockAt(selected.start, size.rows, size.cols);
    const refused = this.checkEditable(dest);
    if (refused) return refused;

    // Adjust references with HyperFormula itself, in a scratch sheet holding the copied block,
    // so the result does not depend on edits made since the copy.
    const scratch = createEngine(placeAt(clip.range.start, clip.data), this.currentMode);
    let data: RawCellContent[][];
    try {
      data = scratch.getFillRangeData(hfRange(clip.range), hfRange(dest), true);
    } finally {
      scratch.destroy();
    }
    this.hf.setCellContents(hfAddress(dest.start), data);
    this.emitChange();
    return { ok: true, changed: expandRange(dest) };
  }

  /**
   * Pastes plain text from outside the grid (tab-separated columns, one line per row), as if
   * typed. A single value pasted onto a range goes into every cell of the range.
   */
  pasteText(target: string, text: string): EditResult {
    const lines = text.replace(/\r\n?/g, "\n").replace(/\n$/, "").split("\n");
    const block = lines.map((line) => line.split("\t"));
    const selected = parseRange(target);
    const single = block.length === 1 && block[0]!.length === 1;
    const width = Math.max(...block.map((line) => line.length));
    const dest = single ? selected : blockAt(selected.start, block.length, width);
    const refused = this.checkEditable(dest);
    if (refused) return refused;
    const { rows, cols } = rangeSize(dest);
    const data = Array.from({ length: rows }, (_, r) =>
      Array.from({ length: cols }, (_, c) => learnerInputToRaw(single ? block[0]![0] : block[r]?.[c])),
    );
    this.clipboard = null;
    this.hf.setCellContents(hfAddress(dest.start), data);
    this.emitChange();
    return { ok: true, changed: expandRange(dest) };
  }

  /**
   * Ctrl+D: copies the top row of the range into the rows below it, adjusting relative
   * references (`=B2*0.784` → `=B3*0.784`…). On a single row, copies the row above into it.
   */
  fillDown(range: string): EditResult {
    const r = parseRange(range);
    if (r.start.row === r.end.row) {
      if (r.start.row === 0) return { ok: true, changed: [] };
      const above = { row: r.start.row - 1, col: r.start.col };
      return this.fill({ start: above, end: { row: above.row, col: r.end.col } }, r);
    }
    const source = { start: r.start, end: { row: r.start.row, col: r.end.col } };
    return this.fill(source, { start: { row: r.start.row + 1, col: r.start.col }, end: r.end });
  }

  /** Ctrl+R: like `fillDown`, but copies the left column to the right. */
  fillRight(range: string): EditResult {
    const r = parseRange(range);
    if (r.start.col === r.end.col) {
      if (r.start.col === 0) return { ok: true, changed: [] };
      const left = { row: r.start.row, col: r.start.col - 1 };
      return this.fill({ start: left, end: { row: r.end.row, col: left.col } }, r);
    }
    const source = { start: r.start, end: { row: r.end.row, col: r.start.col } };
    return this.fill(source, { start: { row: r.start.row, col: r.start.col + 1 }, end: r.end });
  }

  canUndo(): boolean {
    return this.hf.isThereSomethingToUndo();
  }

  canRedo(): boolean {
    return this.hf.isThereSomethingToRedo();
  }

  /** Ctrl+Z. Returns false when there is nothing to undo. */
  undo(): boolean {
    if (!this.hf.isThereSomethingToUndo()) return false;
    this.clipboard = null;
    this.hf.undo();
    this.emitChange();
    return true;
  }

  /** Ctrl+Y. Returns false when there is nothing to redo. */
  redo(): boolean {
    if (!this.hf.isThereSomethingToRedo()) return false;
    this.clipboard = null;
    this.hf.redo();
    this.emitChange();
    return true;
  }

  /**
   * Switches the Excel format by rebuilding the engine (never `updateConfig()`, which breaks
   * formulas). The learner's formulas are rewritten for the new format without changing their
   * meaning (`=ROUND(B2*0.784,2)` ↔ `=ROUND(B2*0,784;2)`); numbers keep their value.
   * A formula that only works in the new format (typed in the wrong one) is kept as typed.
   * Clears undo/redo history and the copy clipboard.
   */
  setMode(mode: FormatMode): void {
    if (mode === this.currentMode) return;
    const from = this.currentMode;
    const learnerContent = new Map<string, RawCellContent>();
    for (const cell of this.editableList) {
      const coords = parseCellAddress(cell);
      const at = hfAddress(coords);
      if (this.hf.isCellEmpty(at) || isSpilledCell(this.hf, at)) continue;
      const formula = this.hf.getCellFormula(at);
      learnerContent.set(
        cell,
        formula !== undefined
          ? convertLearnerFormula(formula, from, mode)
          : valueToInput(this.valueAt(coords), mode, this.formatAt(coords)),
      );
    }
    const isEditable = new Set(this.editableList);
    const next = this.newEngine(
      buildSheetData(this.sheet, mode, this.size, (cell) =>
        isEditable.has(cell) ? (learnerContent.get(cell) ?? null) : undefined,
      ),
      mode,
    );
    this.hf.destroy();
    this.hf = next;
    this.currentMode = mode;
    this.clipboard = null;
    this.emitChange();
  }

  // ---------------------------------------------------------------- internals

  private newEngine(data: RawCellContent[][], mode: FormatMode): HyperFormula {
    return createEngine(data, mode, { todayIso: this.sheet.today, undoLimit: UNDO_LIMIT });
  }

  private isEditableAt(coords: CellCoords): boolean {
    return this.editableRanges.some((range) => rangeContains(range, coords));
  }

  private valueAt(coords: CellCoords): CellValue {
    return toCellValue(this.hf.getCellValue(hfAddress(coords)));
  }

  private formatAt(coords: CellCoords): DisplayFormat {
    const at = hfAddress(coords);
    const kind = numberFormatAt(this.hf, at);
    return { kind, currency: kind === "currency" ? currencySymbolAt(this.hf, at) : undefined };
  }

  private rawInputAt(coords: CellCoords): string {
    const at = hfAddress(coords);
    const formula = this.hf.getCellFormula(at);
    if (formula !== undefined) return formula;
    if (this.hf.isCellEmpty(at) || isSpilledCell(this.hf, at)) return "";
    return valueToInput(this.valueAt(coords), this.currentMode, this.formatAt(coords));
  }

  private checkEditable(range: RangeCoords): EditResult | null {
    const locked = expandRange(range).filter((cell) => !this.isEditableAt(parseCellAddress(cell)));
    return locked.length > 0 ? this.refuseLocked(locked) : null;
  }

  private refuseLocked(cells: string[]): EditResult {
    return {
      ok: false,
      reason: "locked",
      cells,
      messageKa: cells.length === 1 ? LOCKED_ONE_KA : LOCKED_MANY_KA,
    };
  }

  private fill(source: RangeCoords, target: RangeCoords): EditResult {
    const refused = this.checkEditable(target);
    if (refused) return refused;
    const data = this.hf.getFillRangeData(hfRange(source), hfRange(target));
    this.clipboard = null;
    this.hf.setCellContents(hfAddress(target.start), data);
    this.emitChange();
    return { ok: true, changed: expandRange(target) };
  }

  private emitChange(): void {
    this.changeCount++;
    for (const listener of this.listeners) listener();
  }
}
