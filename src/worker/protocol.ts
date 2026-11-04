/**
 * Messages between a web page and the engine worker (`engine-worker.js`).
 *
 * The page never runs HyperFormula itself: it sends these plain JSON messages with `postMessage`
 * and gets plain JSON back. Every request carries an `id`; the reply carries the same `id`.
 * Requests are handled one at a time, in the order they were sent.
 */
import type { GradeResult } from "../engine/grading";
import type { GradableTask, Sheet } from "../engine/schema";
import type { FormatMode } from "../engine/types";
import type { CellView, EditResult } from "../engine/workbook";

/** Bumped whenever a message changes shape. The client checks it on start-up. */
export const PROTOCOL_VERSION = 1;

/** Everything a page needs to draw one sheet, sent after every change. Sheets are small (≤ 12 × 6). */
export interface SheetSnapshot {
  readonly version: number;
  readonly mode: FormatMode;
  readonly rowCount: number;
  readonly columnCount: number;
  /** Row by row: `cells[row * columnCount + col]`. */
  readonly cells: readonly CellView[];
  /** `{ C2: "=B2*0,784" }`: what the learner typed in editable cells (non-empty only). */
  readonly learnerInputs: Readonly<Record<string, string>>;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
}

export type EditOp =
  | { readonly op: "set"; readonly cell: string; readonly raw: string }
  | { readonly op: "clear"; readonly range: string }
  | { readonly op: "copy"; readonly range: string }
  | { readonly op: "paste"; readonly target: string }
  | { readonly op: "pasteText"; readonly target: string; readonly text: string }
  | { readonly op: "fillDown"; readonly range: string }
  | { readonly op: "fillRight"; readonly range: string }
  | { readonly op: "undo" }
  | { readonly op: "redo" }
  | { readonly op: "setMode"; readonly mode: FormatMode };

export type EngineRequest =
  | { readonly type: "hello" }
  | {
      readonly type: "open";
      readonly sheet: Sheet;
      readonly mode: FormatMode;
      /** Inputs to restore (cell → raw input in `mode`); locked cells are skipped. */
      readonly inputs?: Readonly<Record<string, string>>;
    }
  | { readonly type: "edit"; readonly session: number; readonly edit: EditOp }
  | { readonly type: "close"; readonly session: number }
  | { readonly type: "grade"; readonly task: GradableTask; readonly inputs: Readonly<Record<string, string>>; readonly mode: FormatMode }
  | { readonly type: "suggestFormat"; readonly raw: string; readonly mode: FormatMode };

export type EngineReply =
  | { readonly type: "hello"; readonly protocol: number }
  | { readonly type: "opened"; readonly session: number; readonly snapshot: SheetSnapshot }
  | {
      readonly type: "edited";
      /** null for copy, undo, redo and setMode. */
      readonly result: EditResult | null;
      readonly snapshot: SheetSnapshot;
    }
  | { readonly type: "closed" }
  | { readonly type: "graded"; readonly result: GradeResult }
  | { readonly type: "suggested"; readonly mode: FormatMode | null }
  | { readonly type: "error"; readonly message: string };

export interface RequestEnvelope {
  readonly id: number;
  readonly request: EngineRequest;
}

export interface ReplyEnvelope {
  readonly id: number;
  readonly reply: EngineReply;
}
