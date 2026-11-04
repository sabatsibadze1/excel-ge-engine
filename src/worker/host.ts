/**
 * The engine side of the protocol: keeps the open sheets and answers requests. Runs inside the
 * worker (worker.ts); tests call it directly in Node.
 */
import { formatCellAddress } from "../engine/addresses";
import { suggestFormatSwitch } from "../engine/format";
import { gradeTask } from "../engine/grading";
import { SheetSession, type CellView, type EditResult } from "../engine/workbook";
import { PROTOCOL_VERSION, type EditOp, type EngineReply, type EngineRequest, type SheetSnapshot } from "./protocol";

export interface EngineHost {
  handle(request: EngineRequest): EngineReply;
  /** Open sheets, for tests. */
  readonly openSessions: number;
}

export function snapshotOf(session: SheetSession): SheetSnapshot {
  const rows = session.rowCount;
  const cols = session.columnCount;
  const cells: CellView[] = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) cells.push(session.getCell(formatCellAddress(r, c)));
  }
  return {
    version: session.version,
    mode: session.mode,
    rowCount: rows,
    columnCount: cols,
    cells,
    learnerInputs: session.getLearnerInputs(),
    canUndo: session.canUndo(),
    canRedo: session.canRedo(),
  };
}

function applyEdit(session: SheetSession, edit: EditOp): EditResult | null {
  switch (edit.op) {
    case "set":
      return session.setCellInput(edit.cell, edit.raw);
    case "clear":
      return session.clear(edit.range);
    case "copy":
      session.copy(edit.range);
      return null;
    case "paste":
      return session.paste(edit.target);
    case "pasteText":
      return session.pasteText(edit.target, edit.text);
    case "fillDown":
      return session.fillDown(edit.range);
    case "fillRight":
      return session.fillRight(edit.range);
    case "undo":
      session.undo();
      return null;
    case "redo":
      session.redo();
      return null;
    case "setMode":
      session.setMode(edit.mode);
      return null;
  }
}

export function createEngineHost(): EngineHost {
  const sessions = new Map<number, SheetSession>();
  let nextSession = 1;

  const sessionFor = (id: number): SheetSession => {
    const session = sessions.get(id);
    if (!session) throw new Error(`No open sheet with id ${id}`);
    return session;
  };

  const handle = (request: EngineRequest): EngineReply => {
    switch (request.type) {
      case "hello":
        return { type: "hello", protocol: PROTOCOL_VERSION };
      case "open": {
        const session = new SheetSession(request.sheet, { mode: request.mode });
        for (const [cell, raw] of Object.entries(request.inputs ?? {})) {
          if (session.isEditable(cell)) session.setCellInput(cell, raw);
        }
        const id = nextSession++;
        sessions.set(id, session);
        return { type: "opened", session: id, snapshot: snapshotOf(session) };
      }
      case "edit": {
        const session = sessionFor(request.session);
        const result = applyEdit(session, request.edit);
        return { type: "edited", result, snapshot: snapshotOf(session) };
      }
      case "close": {
        sessions.get(request.session)?.destroy();
        sessions.delete(request.session);
        return { type: "closed" };
      }
      case "grade":
        return { type: "graded", result: gradeTask(request.task, request.inputs, request.mode) };
      case "suggestFormat":
        return { type: "suggested", mode: suggestFormatSwitch(request.raw, request.mode) };
    }
  };

  return {
    handle(request) {
      try {
        return handle(request);
      } catch (e) {
        return { type: "error", message: e instanceof Error ? e.message : String(e) };
      }
    },
    get openSessions() {
      return sessions.size;
    },
  };
}
