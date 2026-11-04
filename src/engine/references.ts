/**
 * Reference helpers for the formula editor. F4 cycles the reference at the caret:
 * A1 → $A$1 → A$1 → $A1 → A1 (build plan §6.6).
 */
import { tokenizeFormula, type Token } from "./tokenizer";
import type { FormatMode } from "./types";

export interface CaretEdit {
  readonly formula: string;
  /** Where the caret goes after the edit: right after the changed reference. */
  readonly caret: number;
}

/** A reference the caret can act on: one cell (`B2`) or a whole range (`B2:B6`, `A:C`, `1:3`). */
interface ReferenceUnit {
  readonly parts: readonly Token[];
  readonly start: number;
  readonly end: number;
}

function referenceUnits(tokens: readonly Token[]): ReferenceUnit[] {
  const units: ReferenceUnit[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const first = tokens[i]!;
    if (first.type !== "cell" && first.type !== "column" && first.type !== "row") continue;
    const colon = tokens[i + 1];
    const second = tokens[i + 2];
    if (colon?.type === "colon" && second?.type === first.type) {
      units.push({ parts: [first, second], start: first.start, end: second.end });
      i += 2;
    } else {
      units.push({ parts: [first], start: first.start, end: first.end });
    }
  }
  return units;
}

/** Which reference F4 acts on: the one containing the caret, else the one just before it, else after. */
function unitAtCaret(units: readonly ReferenceUnit[], caret: number): ReferenceUnit | undefined {
  return (
    units.find((u) => caret > u.start && caret < u.end) ??
    units.find((u) => u.end === caret) ??
    units.find((u) => u.start === caret)
  );
}

const CELL_PARTS_RE = /^(\$?)([A-Za-z]+)(\$?)([0-9]+)$/;

interface Anchoring {
  column: boolean;
  row: boolean;
}

/** A1 → $A$1 → A$1 → $A1 → A1 */
function nextCellAnchoring({ column, row }: Anchoring): Anchoring {
  if (!column && !row) return { column: true, row: true };
  if (column && row) return { column: false, row: true };
  if (!column && row) return { column: true, row: false };
  return { column: false, row: false };
}

function anchoringOf(part: Token): Anchoring {
  if (part.type === "cell") {
    const m = CELL_PARTS_RE.exec(part.text);
    return { column: m?.[1] === "$", row: m?.[3] === "$" };
  }
  const anchored = part.text.startsWith("$");
  return { column: anchored, row: anchored };
}

function applyAnchoring(part: Token, anchoring: Anchoring): string {
  const bare = part.text.replaceAll("$", "");
  if (part.type === "cell") {
    const m = CELL_PARTS_RE.exec(part.text);
    if (!m) return part.text;
    return `${anchoring.column ? "$" : ""}${m[2]}${anchoring.row ? "$" : ""}${m[4]}`;
  }
  // Whole columns (A:C) and rows (1:3) have two states only: relative and absolute.
  return part.type === "column" ? `${anchoring.column ? "$" : ""}${bare}` : `${anchoring.row ? "$" : ""}${bare}`;
}

/**
 * F4: cycles the absolute/relative state of the reference at or next to the caret.
 * For a range (`B2:B6`) both ends change together, like in Excel.
 * Text inside "quotes" and function names such as LOG10 are never touched.
 * Returns the formula unchanged (and the same caret) when there is no reference at the caret
 * or the text is not a formula.
 */
export function cycleAbsolute(formula: string, caretIndex: number, mode: FormatMode = "en"): CaretEdit {
  const caret = Math.max(0, Math.min(caretIndex, formula.length));
  if (!formula.startsWith("=")) return { formula, caret };
  const unit = unitAtCaret(referenceUnits(tokenizeFormula(formula, mode)), caret);
  if (!unit) return { formula, caret };

  const [first] = unit.parts;
  const current = anchoringOf(first!);
  const next =
    first!.type === "cell"
      ? nextCellAnchoring(current)
      : { column: !current.column, row: !current.row };

  let rewritten = "";
  let cursor = unit.start;
  for (const part of unit.parts) {
    rewritten += formula.slice(cursor, part.start) + applyAnchoring(part, next);
    cursor = part.end;
  }
  return {
    formula: formula.slice(0, unit.start) + rewritten + formula.slice(unit.end),
    caret: unit.start + rewritten.length,
  };
}
