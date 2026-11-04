import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import { afterAll, describe, expect, it } from "vitest";
import type { ReplyEnvelope } from "../src/worker/protocol";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dir = mkdtempSync(path.join(tmpdir(), "engine-worker-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

describe("build-worker", () => {
  it("builds a standalone worker that starts with the license notice and answers messages", () => {
    const out = path.join(dir, "engine-worker.js");
    execFileSync(process.execPath, ["scripts/build-worker.mjs", out], { cwd: root, encoding: "utf8" });
    const code = readFileSync(out, "utf8");
    expect(code.startsWith("/*!\n * excel-ge-engine")).toBe(true);
    expect(code).toContain("GNU General Public License v3.0");
    expect(code).toContain("https://github.com/sabatsibadze1/excel-ge-engine");

    // Run it like a browser worker: a global scope with onmessage/postMessage, no imports.
    const replies: ReplyEnvelope[] = [];
    const scope: Record<string, unknown> = { postMessage: (m: ReplyEnvelope) => replies.push(m), console };
    scope.self = scope;
    vm.runInNewContext(code, scope);
    const send = (id: number, request: unknown) => (scope.onmessage as (e: { data: unknown }) => void)({ data: { id, request } });
    send(1, { type: "hello" });
    send(2, { type: "open", sheet: { cells: [["a", 2500, null]], editable: ["C1"] }, mode: "en", inputs: { C1: "=B1*0.784" } });
    expect(replies[0]).toEqual({ id: 1, reply: { type: "hello", protocol: 1 } });
    const opened = replies[1]!.reply;
    expect(opened.type === "opened" && opened.snapshot.cells[2]!.display).toBe("1960");
  });
});
