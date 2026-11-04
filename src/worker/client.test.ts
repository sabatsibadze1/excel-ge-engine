import { describe, expect, it } from "vitest";
import type { Sheet } from "../engine/schema";
import { EngineClient, EngineError, type WorkerLike } from "./client";
import { createEngineHost } from "./host";
import type { ReplyEnvelope, RequestEnvelope } from "./protocol";

/** An in-process stand-in for `new Worker(...)`: same messages, answered asynchronously. */
function fakeWorker(): WorkerLike & { terminated: boolean } {
  const host = createEngineHost();
  const worker = {
    onmessage: null as ((event: { data: ReplyEnvelope }) => void) | null,
    onerror: null as ((event: unknown) => void) | null,
    terminated: false,
    postMessage(message: RequestEnvelope) {
      const copy = structuredClone(message);
      setTimeout(() => worker.onmessage?.({ data: structuredClone({ id: copy.id, reply: host.handle(copy.request) }) }), 0);
    },
    terminate() {
      worker.terminated = true;
    },
  };
  return worker;
}

const sheet: Sheet = {
  cells: [
    ["name", "gross", "net"],
    ["a", 2500, null],
    ["b", 1500, null],
  ],
  editable: ["C2:C3"],
};

describe("EngineClient + RemoteSheet", () => {
  it("says hello, opens a sheet, edits it and grades it", async () => {
    const client = new EngineClient(fakeWorker());
    await client.hello(1000);
    const remote = await client.openSheet(sheet, { mode: "en" });
    expect(remote.getCell("B2").display).toBe("2500");
    expect(remote.isEditable("C2")).toBe(true);

    let notified = 0;
    remote.subscribe(() => notified++);
    const result = await remote.setCellInput("C2", "=B2*0.784");
    expect(result).toEqual({ ok: true, changed: ["C2"] });
    expect(remote.getDisplayValue("C2")).toBe("1960");
    expect(remote.version).toBe(1);
    expect(notified).toBe(1);

    expect(remote.copy("C2")).toBe("1960");
    void remote.paste("C3");
    await remote.idle();
    expect(remote.getRawInput("C3")).toBe("=B3*0.784");

    const grade = await client.grade(
      { sheet, check: { cells: "C2:C3", solution: "=B2*0.784", vary: ["B2:B3"], mustUse: [] } },
      remote.getLearnerInputs(),
      remote.mode,
    );
    expect(grade.pass).toBe(true);

    await remote.setMode("ka");
    expect(remote.mode).toBe("ka");
    expect(remote.getRawInput("C2")).toBe("=B2*0,784");
    expect(await client.suggestFormatSwitch("2500.5", "ka")).toBe("en");

    remote.destroy();
    await expect(remote.setCellInput("C2", "1")).rejects.toBeInstanceOf(EngineError);
  });

  it("returns an empty locked cell outside the sheet", async () => {
    const client = new EngineClient(fakeWorker());
    const remote = await client.openSheet(sheet, { mode: "en" });
    expect(remote.getCell("F9")).toMatchObject({ display: "", isEditable: false });
  });

  it("rejects pending and later requests when the worker fails", async () => {
    const worker = fakeWorker();
    const client = new EngineClient(worker);
    worker.postMessage = () => {};
    const pending = client.hello(1000);
    worker.onerror?.(new Event("error"));
    await expect(pending).rejects.toThrow("The engine worker stopped.");
    await expect(client.suggestFormatSwitch("1", "en")).rejects.toBeInstanceOf(EngineError);
  });

  it("times out when the worker never answers", async () => {
    const worker = fakeWorker();
    worker.postMessage = () => {};
    await expect(new EngineClient(worker).hello(20)).rejects.toThrow("did not answer");
  });
});
