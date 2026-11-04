/**
 * A1-style cell addresses and ranges ("C2", "C2:C6").
 * Rows and columns are 0-based in coordinates: "A1" is { row: 0, col: 0 }.
 */

export interface CellCoords {
  readonly row: number;
  readonly col: number;
}

/** A rectangular range. Always normalized: `start` is the top-left cell, `end` the bottom-right. */
export interface RangeCoords {
  readonly start: CellCoords;
  readonly end: CellCoords;
}

/** Excel's limits: 16 384 columns (XFD) and 1 048 576 rows. */
export const MAX_COLUMNS = 16_384;
export const MAX_ROWS = 1_048_576;

const CELL_ADDRESS_RE = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/;

/** 0 → "A", 25 → "Z", 26 → "AA". */
export function columnToLetters(col: number): string {
  if (!Number.isInteger(col) || col < 0 || col >= MAX_COLUMNS) {
    throw new RangeError(`Column index out of range: ${col}`);
  }
  let n = col + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/** "A" → 0, "Z" → 25, "AA" → 26. Case-insensitive. */
export function lettersToColumn(letters: string): number {
  const upper = letters.toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(upper)) {
    throw new RangeError(`Not a column name: "${letters}"`);
  }
  let n = 0;
  for (const ch of upper) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

/** Parses "C2" (case-insensitive, surrounding spaces allowed). Returns null if invalid. */
export function tryParseCellAddress(address: string): CellCoords | null {
  const match = CELL_ADDRESS_RE.exec(address.trim().toUpperCase());
  if (!match) return null;
  const col = lettersToColumn(match[1]!);
  const row = Number(match[2]) - 1;
  if (col >= MAX_COLUMNS || row >= MAX_ROWS) return null;
  return { row, col };
}

/** Parses "C2" into { row: 1, col: 2 }. Throws on an invalid address. */
export function parseCellAddress(address: string): CellCoords {
  const coords = tryParseCellAddress(address);
  if (!coords) throw new RangeError(`Not a cell address: "${address}"`);
  return coords;
}

/** { row: 1, col: 2 } → "C2". */
export function formatCellAddress(row: number, col: number): string {
  if (!Number.isInteger(row) || row < 0 || row >= MAX_ROWS) {
    throw new RangeError(`Row index out of range: ${row}`);
  }
  return `${columnToLetters(col)}${row + 1}`;
}

export function isCellAddress(value: string): boolean {
  return tryParseCellAddress(value) !== null;
}

/** Parses "C2:C6" or a single cell "C2" (corners in any order). Returns null if invalid. */
export function tryParseRange(range: string): RangeCoords | null {
  const parts = range.split(":");
  if (parts.length > 2) return null;
  const a = tryParseCellAddress(parts[0]!);
  const b = parts.length === 2 ? tryParseCellAddress(parts[1]!) : a;
  if (!a || !b) return null;
  return {
    start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
}

/** Parses "C2:C6" or "C2". Throws on an invalid range. */
export function parseRange(range: string): RangeCoords {
  const coords = tryParseRange(range);
  if (!coords) throw new RangeError(`Not a cell or range address: "${range}"`);
  return coords;
}

export function isRangeAddress(value: string): boolean {
  return tryParseRange(value) !== null;
}

/** Normalized text form: "C2" for a single cell, "C2:C6" otherwise. */
export function formatRange(range: RangeCoords): string {
  const start = formatCellAddress(range.start.row, range.start.col);
  if (range.start.row === range.end.row && range.start.col === range.end.col) return start;
  return `${start}:${formatCellAddress(range.end.row, range.end.col)}`;
}

export function rangeSize(range: RangeCoords): { rows: number; cols: number } {
  return { rows: range.end.row - range.start.row + 1, cols: range.end.col - range.start.col + 1 };
}

/** All cells of a range in reading order (left to right, then top to bottom). */
export function expandRange(range: string | RangeCoords): string[] {
  const r = typeof range === "string" ? parseRange(range) : range;
  const cells: string[] = [];
  for (let row = r.start.row; row <= r.end.row; row++) {
    for (let col = r.start.col; col <= r.end.col; col++) {
      cells.push(formatCellAddress(row, col));
    }
  }
  return cells;
}

/** Expands a list of cells/ranges into unique cells, keeping the first-seen order. */
export function expandRanges(ranges: readonly string[]): string[] {
  const seen = new Set<string>();
  const cells: string[] = [];
  for (const range of ranges) {
    for (const cell of expandRange(range)) {
      if (!seen.has(cell)) {
        seen.add(cell);
        cells.push(cell);
      }
    }
  }
  return cells;
}

export function rangeContains(range: RangeCoords, cell: CellCoords): boolean {
  return (
    cell.row >= range.start.row &&
    cell.row <= range.end.row &&
    cell.col >= range.start.col &&
    cell.col <= range.end.col
  );
}

/** Normalizes a user-supplied address: " c2 " → "C2". Throws if invalid. */
export function normalizeCellAddress(address: string): string {
  const { row, col } = parseCellAddress(address);
  return formatCellAddress(row, col);
}
