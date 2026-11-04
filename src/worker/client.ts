/**
 * The page side of the protocol. No HyperFormula here: it only sends messages to the engine worker
 * and keeps the latest snapshot of each sheet, so the page can draw synchronously.
 *
 *   const engine = await EngineClient.start("/engine-worker.js");
 *   const sheet = await engine.openSheet(lesson.sheet, { mode: "ka" });
 *   sheet.getCell("C2").display;              // from the latest snapshot
 *   await sheet.setCellInput("C2", "=B2*0,784");
 *   const result = await engine.grade(lesson, sheet.getLearnerInputs(), sheet.mode);
 *
 * Edits are asynchronous (a round trip to the worker, usually a few milliseconds). Reads return the
 * state after the last edit that has come back; `await sheet.idle()` waits for every edit sent so far.
 */
import { formatCellAddress, parseCellAddress, parseRange } from "../engine/addresses";
import type { GradeResult } from "../engine/grading";
import type { GradableTask, Sheet } from "../engine/schema";
import type { CellValue, FormatMode } from "../engine/types";
import type { CellView, EditResult } from "../engine/workbook";
import {
  PROTOCOL_VERSION,
  type EditOp,
  type EngineReply,
  type EngineRequest,
  type ReplyEnvelope,
  type RequestEnvelope,
  type SheetSnapshot,
} from "./protocol";

/** The part of a Web Worker the client uses (so tests can pass a fake). */
export interface WorkerLike {
  postMessage(message: RequestEnvelope): void;
  onmessage: ((event: { data: ReplyEnvelope }) => void) | null;
  onerror: ((event: unknown) => void) | null;
  terminate(): void;
}

export class EngineError extends Error {
  override name = "EngineError";
}

type ReplyOf<T extends EngineReply["type"]> = Extract<EngineReply, { type: T }>;

export class EngineClient {
  private readonly pending = new Map<number, { resolve: (reply: EngineReply) => void; reject: (error: Error) => void }>();
  private nextId = 1;
  private failure: EngineError | null = null;

  constructor(private readonly worker: WorkerLike) {
    worker.onmessage = (event) => {
      const { id, reply } = event.data;
      const waiter = this.pending.get(id);
      if (!waiter) return;
      this.pending.delete(id);
      if (reply.type === "error") waiter.reject(new EngineError(reply.message));
      else waiter.resolve(reply);
    };
    worker.onerror = () => this.fail(new EngineError("The engine worker stopped."));
  }

  /** Starts the worker from `url` and checks it speaks this protocol version. */
  static async start(url: string, options: { timeoutMs?: number } = {}): Promise<EngineClient> {
    const client = new EngineClient(new Worker(url) as unknown as WorkerLike);
    await client.hello(options.timeoutMs ?? 15_000);
    return client;
  }

  async hello(timeoutMs: number): Promise<void> {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new EngineError("The engine worker did not answer.")), timeoutMs);
    });
    try {
      const reply = await Promise.race([this.send("hello", { type: "hello" }), timeout]);
      if (reply.protocol !== PROTOCOL_VERSION) {
        throw new EngineError(`Engine worker speaks protocol ${reply.protocol}, this page needs ${PROTOCOL_VERSION}.`);
      }
    } finally {
      clearTimeout(timer);
    }
  }

  send<T extends EngineReply["type"]>(expect: T, request: EngineRequest): Promise<ReplyOf<T>> {
    if (this.failure) return Promise.reject(this.failure);
    const id = this.nextId++;
    return new Promise<ReplyOf<T>>((resolve, reject) => {
      this.pending.set(id, {
        resolve: (reply) =>
          reply.type === expect ? resolve(reply as ReplyOf<T>) : reject(new EngineError(`Expected ${expect}, got ${reply.type}`)),
        reject,
      });
      this.worker.postMessage({ id, request });
    });
  }

  async openSheet(sheet: Sheet, options: { mode: FormatMode; inputs?: Readonly<Record<string, string>> }): Promise<RemoteSheet> {
    const reply = await this.send("opened", { type: "open", sheet, mode: options.mode, inputs: options.inputs });
    return new RemoteSheet(this, reply.session, reply.snapshot);
  }

  async grade(task: GradableTask, inputs: Readonly<Record<string, string>>, mode: FormatMode): Promise<GradeResult> {
    // Send only what grading needs (a lesson also carries theory, hints…).
    const reply = await this.send("graded", { type: "grade", task: { sheet: task.sheet, check: task.check }, inputs, mode });
    return reply.result;
  }

  async suggestFormatSwitch(raw: string, mode: FormatMode): Promise<FormatMode | null> {
    return (await this.send("suggested", { type: "suggestFormat", raw, mode })).mode;
  }

  terminate(): void {
    this.fail(new EngineError("The engine worker was stopped."));
    this.worker.terminate();
  }

  private fail(error: EngineError): void {
    this.failure ??= error;
    for (const waiter of this.pending.values()) waiter.reject(error);
    this.pending.clear();
  }
}

const EMPTY_CELL = (address: string, row: number, col: number): CellView => ({
  address,
  row,
  col,
  value: null,
  display: "",
  input: "",
  isFormula: false,
  isEditable: false,
  isSpilled: false,
  numberFormat: null,
  error: null,
  align: "left",
});

/** One open sheet in the worker, mirrored by its latest snapshot. */
export class RemoteSheet {
  private snap: SheetSnapshot;
  /** Counts the changes seen by this page (the worker's own version restarts on a format switch). */
  private changes = 0;
  private readonly listeners = new Set<() => void>();
  private queue: Promise<unknown> = Promise.resolve();
  private closed = false;

  constructor(
    private readonly client: EngineClient,
    readonly id: number,
    snapshot: SheetSnapshot,
  ) {
    this.snap = snapshot;
  }

  get version(): number {
    return this.changes;
  }
  get mode(): FormatMode {
    return this.snap.mode;
  }
  get rowCount(): number {
    return this.snap.rowCount;
  }
  get columnCount(): number {
    return this.snap.columnCount;
  }

  getCell(address: string): CellView {
    const { row, col } = parseCellAddress(address);
    if (row >= this.snap.rowCount || col >= this.snap.columnCount) return EMPTY_CELL(formatCellAddress(row, col), row, col);
    return this.snap.cells[row * this.snap.columnCount + col]!;
  }
  getValue(address: string): CellValue {
    return this.getCell(address).value;
  }
  getDisplayValue(address: string): string {
    return this.getCell(address).display;
  }
  getRawInput(address: string): string {
    return this.getCell(address).input;
  }
  isEditable(address: string): boolean {
    return this.getCell(address).isEditable;
  }
  getLearnerInputs(): Record<string, string> {
    return { ...this.snap.learnerInputs };
  }
  canUndo(): boolean {
    return this.snap.canUndo;
  }
  canRedo(): boolean {
    return this.snap.canRedo;
  }

  /**
   * Ctrl+C: returns the tab-separated text for the system clipboard right away (from the snapshot,
   * as the engine would) and tells the worker to remember the cells for `paste`.
   */
  copy(range: string): string {
    const { start, end } = parseRange(range);
    const lines: string[] = [];
    for (let row = start.row; row <= end.row; row++) {
      const values: string[] = [];
      for (let col = start.col; col <= end.col; col++) {
        values.push(this.getCell(formatCellAddress(row, col)).display);
      }
      lines.push(values.join("\t"));
    }
    void this.edit({ op: "copy", range }).catch(() => {});
    return lines.join("\n");
  }

  setCellInput(address: string, raw: string): Promise<EditResult> {
    return this.editResult({ op: "set", cell: address, raw });
  }
  clear(range: string): Promise<EditResult> {
    return this.editResult({ op: "clear", range });
  }
  paste(target: string): Promise<EditResult> {
    return this.editResult({ op: "paste", target });
  }
  pasteText(target: string, text: string): Promise<EditResult> {
    return this.editResult({ op: "pasteText", target, text });
  }
  fillDown(range: string): Promise<EditResult> {
    return this.editResult({ op: "fillDown", range });
  }
  fillRight(range: string): Promise<EditResult> {
    return this.editResult({ op: "fillRight", range });
  }
  async undo(): Promise<void> {
    await this.edit({ op: "undo" });
  }
  async redo(): Promise<void> {
    await this.edit({ op: "redo" });
  }
  async setMode(mode: FormatMode): Promise<void> {
    await this.edit({ op: "setMode", mode });
  }

  /** Resolves when every edit sent so far has come back. */
  async idle(): Promise<void> {
    await this.queue.catch(() => {});
  }

  /** Change notification (for React's useSyncExternalStore). */
  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  destroy(): void {
    if (this.closed) return;
    this.closed = true;
    this.listeners.clear();
    void this.client.send("closed", { type: "close", session: this.id }).catch(() => {});
  }

  private async editResult(edit: EditOp): Promise<EditResult> {
    return (await this.edit(edit))!;
  }

  private edit(edit: EditOp): Promise<EditResult | null> {
    if (this.closed) return Promise.reject(new EngineError("This sheet is closed."));
    const sent = this.client.send("edited", { type: "edit", session: this.id, edit }).then((reply) => {
      const changed = reply.snapshot.version !== this.snap.version || reply.snapshot.mode !== this.snap.mode;
      this.snap = reply.snapshot;
      if (changed) {
        this.changes++;
        for (const listener of this.listeners) listener();
      }
      return reply.result;
    });
    this.queue = sent.catch(() => {});
    return sent;
  }
}
