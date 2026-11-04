/**
 * Entry point of `engine-worker.js`: a standalone program (HyperFormula + this engine) that a web
 * page starts with `new Worker("/engine-worker.js")` and talks to only through JSON messages
 * (see protocol.ts). Build it with `node scripts/build-worker.mjs <outfile>`.
 */
import { createEngineHost } from "./host";
import type { ReplyEnvelope, RequestEnvelope } from "./protocol";

interface WorkerScope {
  onmessage: ((event: { data: RequestEnvelope }) => void) | null;
  postMessage(message: ReplyEnvelope): void;
}

const scope = globalThis as unknown as WorkerScope;
const host = createEngineHost();

scope.onmessage = (event) => {
  const { id, request } = event.data;
  scope.postMessage({ id, reply: host.handle(request) });
};
