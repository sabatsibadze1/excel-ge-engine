import { describe, expect, it } from "vitest";
import type { Sheet } from "../engine/schema";
import { createEngineHost } from "./host";
import { PROTOCOL_VERSION } from "./protocol";

const sheet: Sheet = {
  cells: [
    ["თანამშრომელი", "დარიცხული", "ხელზე"],
    ["ნათია", 2500, null],
    ["ზურაბ", 1500, null],
  ],
  editable: ["C2:C3"],
};
const check = { cells: "C2:C3", solution: "=B2*0.784", vary: ["B2:B3"], mustUse: [] };

describe("engine host", () => {
  it("answers hello with the protocol version", () => {
    expect(createEngineHost().handle({ type: "hello" })).toEqual({ type: "hello", protocol: PROTOCOL_VERSION });
  });

  it("opens a sheet, restores inputs, edits it and sends snapshots", () => {
    const host = createEngineHost();
    const opened = host.handle({ type: "open", sheet, mode: "ka", inputs: { C2: "=B2*0,784", A1: "locked, skipped" } });
    if (opened.type !== "opened") throw new Error(opened.type);
    const snap = opened.snapshot;
    expect([snap.rowCount, snap.columnCount, snap.mode]).toEqual([3, 3, "ka"]);
    expect(snap.cells[1 * 3 + 2]!.display).toBe("1960");
    expect(snap.cells[0]!.display).toBe("თანამშრომელი");
    expect(snap.learnerInputs).toEqual({ C2: "=B2*0,784" });

    const edited = host.handle({ type: "edit", session: opened.session, edit: { op: "set", cell: "C3", raw: "=B3*0,784" } });
    if (edited.type !== "edited") throw new Error(edited.type);
    expect(edited.result).toEqual({ ok: true, changed: ["C3"] });
    expect(edited.snapshot.cells[2 * 3 + 2]!.display).toBe("1176");

    const locked = host.handle({ type: "edit", session: opened.session, edit: { op: "set", cell: "B2", raw: "1" } });
    expect(locked.type === "edited" && locked.result?.ok).toBe(false);

    const mode = host.handle({ type: "edit", session: opened.session, edit: { op: "setMode", mode: "en" } });
    expect(mode.type === "edited" && mode.snapshot.learnerInputs).toEqual({ C2: "=B2*0.784", C3: "=B3*0.784" });

    expect(host.openSessions).toBe(1);
    expect(host.handle({ type: "close", session: opened.session })).toEqual({ type: "closed" });
    expect(host.openSessions).toBe(0);
  });

  it("copies and pastes with relative references", () => {
    const host = createEngineHost();
    const opened = host.handle({ type: "open", sheet, mode: "en", inputs: { C2: "=B2*0.784" } });
    if (opened.type !== "opened") throw new Error(opened.type);
    host.handle({ type: "edit", session: opened.session, edit: { op: "copy", range: "C2" } });
    const pasted = host.handle({ type: "edit", session: opened.session, edit: { op: "paste", target: "C3" } });
    expect(pasted.type === "edited" && pasted.snapshot.learnerInputs.C3).toBe("=B3*0.784");
  });

  it("grades with swap-the-inputs", () => {
    const host = createEngineHost();
    const good = host.handle({ type: "grade", task: { sheet, check }, inputs: { C2: "=B2*0.784", C3: "=B3*0.784" }, mode: "en" });
    expect(good.type === "graded" && good.result.pass).toBe(true);
    const hardcoded = host.handle({ type: "grade", task: { sheet, check }, inputs: { C2: "=2500*0.784", C3: "=1500*0.784" }, mode: "en" });
    expect(hardcoded.type === "graded" && !hardcoded.result.pass && hardcoded.result.reason).toBe("hardcoded");
  });

  it("suggests a format switch", () => {
    expect(createEngineHost().handle({ type: "suggestFormat", raw: "=ROUND(B2*0,784;2)", mode: "en" })).toEqual({ type: "suggested", mode: "ka" });
  });

  it("returns an error reply instead of throwing", () => {
    const reply = createEngineHost().handle({ type: "edit", session: 99, edit: { op: "undo" } });
    expect(reply).toEqual({ type: "error", message: "No open sheet with id 99" });
  });
});
