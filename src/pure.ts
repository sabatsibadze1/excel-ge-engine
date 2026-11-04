/**
 * Everything a web page may use WITHOUT loading HyperFormula: the worker client, the lesson
 * format, addresses, the formula converter and display helpers. A test (pure.test.ts) fails if
 * anything imported from here pulls in HyperFormula.
 *
 * The spreadsheet itself (HyperFormula) only runs inside `engine-worker.js` (src/worker/worker.ts).
 */
export { EngineClient, EngineError, RemoteSheet, type WorkerLike } from "./worker/client";
export { PROTOCOL_VERSION, type EditOp, type EngineReply, type EngineRequest, type SheetSnapshot } from "./worker/protocol";

export * from "./engine/types";
export * from "./engine/addresses";
export * from "./engine/schema";
export * from "./engine/facts";
export * from "./engine/convert";
export * from "./engine/display";
export { cycleAbsolute, type CaretEdit } from "./engine/references";
export { functionNames, tokenizeFormula, type Token, type TokenType } from "./engine/tokenizer";
export { ERROR_CODES, explainError, unsupportedFunctionTip, type ErrorExplanation } from "./engine/errors";

// Types only (erased at build time).
export type { CellGrade, GradeFail, GradeFailReason, GradePass, GradeResult } from "./engine/grading";
export type { CellView, EditResult } from "./engine/workbook";
