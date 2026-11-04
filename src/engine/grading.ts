/**
 * Swap-the-inputs grading (build plan §6.1, SCHEMA.md "Grading rules").
 *
 * For every graded cell, in reading order:
 *   1. it must hold a formula (starts with "=");
 *   2. if `mustUse` is set, at least one alternative must be satisfied (whole function names before
 *      "(", outside quoted text; an alternative may be a list of names that must all appear);
 *   3. the learner's formulas and the solution are computed on the lesson data and must match:
 *      numbers within the RELATIVE tolerance (default 1e-9), text equal after trimming,
 *      errors only the same error type;
 *   4. swap test: 3 seeded trials change every `vary` number (never to the original, never to 0)
 *      and date (±1–60 days); then every `cases` dataset is applied on top of the lesson data.
 *      Both sides must still match. Failing only here gives the "hardcoded" message.
 *
 * Grading never touches the live grid: it builds one throwaway engine per task with two sheets,
 * `sol` (the solution, filled across the range like Ctrl+D) and `lrn` (the learner's inputs), in the
 * learner's format, and changes inputs with `batch()`.
 */
import type { HyperFormula } from "hyperformula";
import { expandRange, expandRanges, parseCellAddress, parseRange } from "./addresses";
import { formatValue } from "./display";
import { explainError } from "./errors";
import { excelRound } from "./excel-functions";
import { convertFormula } from "./format";
import {
  createEngine,
  currencySymbolAt,
  hfAddress,
  hfAddressOf,
  hfRange,
  isoDateToSerial,
  isSpilledCell,
  numberFormatAt,
  serialToIsoDate,
  toCellValue,
} from "./hyperformula";
import { createRng, perturbDateSerial, perturbNumber } from "./random";
import { isDateCell, type Case, type GradableTask, type MustUseAlternative, type SheetCell } from "./schema";
import { buildSheetData, editableCellSet, learnerInputToRaw, lessonValueToRaw, sheetSize } from "./sheet-data";
import { functionNames } from "./tokenizer";
import { isCellError, type CellValue, type ExcelErrorType, type FormatMode } from "./types";

/** Relative tolerance for numbers (SCHEMA.md): |a-b| <= tolerance × max(1, |a|, |b|). */
export const DEFAULT_TOLERANCE = 1e-9;
export const DEFAULT_TRIALS = 3;
/** Fixed default seed: learners get exactly the swap trials the lesson checker verified. */
export const DEFAULT_SEED = 20_261_102;

export interface GradeOptions {
  /** Seed for the swap test's random changes. Default `DEFAULT_SEED`. */
  readonly seed?: number;
  /** Number of swap trials. Default 3. */
  readonly trials?: number;
}

export type GradeFailReason =
  | "empty"
  | "not_formula"
  | "missing_function"
  | "excel_error"
  | "wrong_value"
  | "hardcoded";

/** The verdict for one graded cell. */
export interface CellGrade {
  readonly cell: string;
  readonly pass: boolean;
  readonly reason?: GradeFailReason;
  /** The solution's value, formatted for the learner's mode (for "hardcoded": in that dataset). */
  readonly expected?: string;
  /** The learner's value, formatted for the learner's mode. */
  readonly actual?: string;
  readonly errorType?: ExcelErrorType;
  /** For "hardcoded": the 1-based swap trial that exposed it… */
  readonly trial?: number;
  /** …or the 1-based `check.cases` dataset that exposed it. */
  readonly caseIndex?: number;
}

export interface GradePass {
  readonly pass: true;
  readonly messageKa: string;
  /** One entry per graded cell, in reading order. */
  readonly cells: CellGrade[];
}

export interface GradeFail {
  readonly pass: false;
  readonly reason: GradeFailReason;
  /** The first failing cell in reading order. */
  readonly cell: string;
  readonly expected?: string;
  readonly actual?: string;
  readonly errorType?: ExcelErrorType;
  /** For "missing_function": the alternatives the task accepts. */
  readonly mustUse?: MustUseAlternative[];
  /** For "hardcoded": which swap trial or case exposed it, and the changed inputs. */
  readonly trial?: number;
  readonly caseIndex?: number;
  readonly trialInputs?: Case;
  /** Short, friendly Georgian message for the learner. */
  readonly messageKa: string;
  readonly cells: CellGrade[];
}

export type GradeResult = GradePass | GradeFail;

/**
 * Two values match: numbers within the relative tolerance (or, with `compareDecimals`, equal after
 * Excel-style rounding to that many decimals), text equal after trimming, errors of the same type.
 */
export function valuesMatch(
  expected: CellValue,
  actual: CellValue,
  tolerance = DEFAULT_TOLERANCE,
  compareDecimals?: number | null,
): boolean {
  if (typeof expected === "number" && typeof actual === "number") {
    if (typeof compareDecimals === "number") {
      return excelRound(expected, compareDecimals) === excelRound(actual, compareDecimals);
    }
    return Math.abs(expected - actual) <= tolerance * Math.max(1, Math.abs(expected), Math.abs(actual));
  }
  if (typeof expected === "string" && typeof actual === "string") return expected.trim() === actual.trim();
  if (isCellError(expected) || isCellError(actual)) {
    return isCellError(expected) && isCellError(actual) && expected.type === actual.type;
  }
  return expected === actual;
}

/** True if the formula satisfies at least one `mustUse` alternative (or there are none). */
export function satisfiesMustUse(
  formula: string,
  mustUse: readonly MustUseAlternative[] | undefined,
  mode: FormatMode,
): boolean {
  if (!mustUse || mustUse.length === 0) return true;
  const used = new Set(functionNames(formula, mode));
  return mustUse.some((alternative) =>
    (Array.isArray(alternative) ? alternative : [alternative]).every((name) => used.has(name.toUpperCase())),
  );
}

/** [["INDEX","MATCH"], "XLOOKUP", "VLOOKUP"] → "INDEX და MATCH, XLOOKUP ან VLOOKUP" */
export function describeMustUseKa(mustUse: readonly MustUseAlternative[]): string {
  const options = mustUse.map((alternative) => {
    const names = (Array.isArray(alternative) ? alternative : [alternative]).map((n) => n.toUpperCase());
    return names.length > 1 ? `${names.slice(0, -1).join(", ")} და ${names.at(-1)}` : names[0] ?? "";
  });
  return options.length > 1 ? `${options.slice(0, -1).join(", ")} ან ${options.at(-1)}` : options[0] ?? "";
}

/** The numeric (or date) `vary` cells and their lesson values, in order. */
function varyCells(task: GradableTask): { cell: string; base: number | { date: string } }[] {
  const cells: { cell: string; base: number | { date: string } }[] = [];
  for (const cell of expandRanges(task.check.vary)) {
    const { row, col } = parseCellAddress(cell);
    const value = task.sheet.cells[row]?.[col];
    if (typeof value === "number" || isDateCell(value)) cells.push({ cell, base: value });
  }
  return cells;
}

/**
 * The swap test's changed inputs: one record per trial, mapping every `vary` number or date to its
 * new value. Deterministic for a seed, so the checker sees exactly the trials learners get.
 */
export function swapTrialInputs(task: GradableTask, options: GradeOptions = {}): Case[] {
  const rng = createRng(options.seed ?? DEFAULT_SEED);
  const cells = varyCells(task);
  return Array.from({ length: options.trials ?? DEFAULT_TRIALS }, () => {
    const inputs: Case = {};
    for (const { cell, base } of cells) {
      inputs[cell] =
        typeof base === "number"
          ? perturbNumber(base, rng)
          : { date: serialToIsoDate(perturbDateSerial(isoDateToSerial(base.date), rng)) };
    }
    return inputs;
  });
}

function gridSize(task: GradableTask) {
  const caseCells = (task.check.cases ?? []).flatMap((c) => Object.keys(c));
  return sheetSize(task.sheet, [task.check.cells, ...task.check.vary, ...caseCells]);
}

/** Writes the solution into the graded range of one sheet, filled like Ctrl+D. */
function fillSolution(hf: HyperFormula, sheet: number, task: GradableTask, mode: FormatMode): void {
  const range = parseRange(task.check.cells);
  const topLeft = hfAddress(range.start, sheet);
  hf.setCellContents(topLeft, convertFormula(task.check.solution, "en", mode));
  if (range.start.row !== range.end.row || range.start.col !== range.end.col) {
    const source = hfRange({ start: range.start, end: range.start }, sheet);
    hf.setCellContents(topLeft, hf.getFillRangeData(source, hfRange(range, sheet)));
  }
}

/**
 * A one-sheet engine with the lesson data and the solution filled across `check.cells`
 * (English format unless `mode` says otherwise). The caller must `destroy()` it.
 */
export function buildSolutionEngine(task: GradableTask, mode: FormatMode = "en"): HyperFormula {
  const graded = new Set(expandRange(task.check.cells));
  const data = buildSheetData(task.sheet, mode, gridSize(task), (cell) => (graded.has(cell) ? null : undefined));
  const hf = createEngine(data, mode, { todayIso: task.sheet.today });
  fillSolution(hf, 0, task, mode);
  return hf;
}

/** The solution formula of every graded cell (English format), after filling across the range. */
export function solutionFormulas(task: GradableTask): Record<string, string> {
  const hf = buildSolutionEngine(task);
  try {
    return Object.fromEntries(
      expandRange(task.check.cells).map((cell) => [cell, hf.getCellFormula(hfAddressOf(cell)) ?? ""]),
    );
  } finally {
    hf.destroy();
  }
}

/** Sets cells (lesson-style values) in the given sheets, as one batch. */
export function applyValues(hf: HyperFormula, sheets: readonly number[], values: Case, mode: FormatMode): void {
  hf.batch(() => {
    for (const [cell, value] of Object.entries(values)) {
      for (const sheet of sheets) hf.setCellContents(hfAddressOf(cell, sheet), lessonValueToRaw(value, mode));
    }
  });
}

/** The lesson's own values for some cells (to undo a trial or case). */
export function baseValues(task: GradableTask, cells: Iterable<string>): Case {
  const values: Case = {};
  for (const cell of cells) {
    const { row, col } = parseCellAddress(cell);
    values[cell] = (task.sheet.cells[row]?.[col] ?? null) as SheetCell;
  }
  return values;
}

function valueAt(hf: HyperFormula, sheet: number, cell: string): CellValue {
  return toCellValue(hf.getCellValue(hfAddressOf(cell, sheet)));
}

function shownAt(hf: HyperFormula, sheet: number, cell: string, mode: FormatMode): string {
  const at = hfAddressOf(cell, sheet);
  const kind = numberFormatAt(hf, at);
  const currency = kind === "currency" ? currencySymbolAt(hf, at) : undefined;
  return formatValue(valueAt(hf, sheet, cell), mode, { kind, currency });
}

/** Normalizes input keys ("c2" → "C2") and keeps only editable cells. */
function normalizeInputs(task: GradableTask, inputs: Readonly<Record<string, string>>): Record<string, string> {
  const editable = editableCellSet(task.sheet);
  const normalized: Record<string, string> = {};
  for (const [key, raw] of Object.entries(inputs)) {
    const cell = key.trim().toUpperCase();
    if (editable.has(cell) && typeof raw === "string") normalized[cell] = raw;
  }
  return normalized;
}

interface CellCheck extends CellGrade {
  /** The graded cell is empty but shows a value spilled from another cell's formula. */
  readonly spilled?: boolean;
}

/**
 * Grades one task. `learnerInputs` are the learner's raw inputs as typed, in `mode`, keyed by cell
 * (pass `session.getLearnerInputs()` from a `SheetSession`). Never throws for learner input.
 */
export function gradeTask(
  task: GradableTask,
  learnerInputs: Readonly<Record<string, string>>,
  mode: FormatMode,
  options: GradeOptions = {},
): GradeResult {
  const { check } = task;
  const matches = (expected: CellValue, actual: CellValue) =>
    valuesMatch(expected, actual, check.tolerance ?? DEFAULT_TOLERANCE, check.compareDecimals);
  const gradedCells = expandRange(check.cells);
  const graded = new Set(gradedCells);
  const editable = editableCellSet(task.sheet);
  const inputs = normalizeInputs(task, learnerInputs);
  const size = gridSize(task);

  const hf = createEngine(
    {
      sol: buildSheetData(task.sheet, mode, size, (cell) => (graded.has(cell) ? null : undefined)),
      lrn: buildSheetData(task.sheet, mode, size, (cell) =>
        editable.has(cell) ? learnerInputToRaw(inputs[cell]) : undefined,
      ),
    },
    mode,
    { todayIso: task.sheet.today },
  );
  const sol = hf.getSheetId("sol")!;
  const lrn = hf.getSheetId("lrn")!;
  try {
    fillSolution(hf, sol, task, mode);

    // Rules 1–3, on the lesson data.
    const baseChecks = gradedCells.map((cell): CellCheck => {
      const raw = inputs[cell];
      if (raw === undefined || raw.trim() === "") {
        return isSpilledCell(hf, hfAddressOf(cell, lrn))
          ? { cell, pass: false, reason: "not_formula", spilled: true, actual: shownAt(hf, lrn, cell, mode) }
          : { cell, pass: false, reason: "empty" };
      }
      const formula = raw.trimStart();
      if (!formula.startsWith("=")) {
        return { cell, pass: false, reason: "not_formula", actual: shownAt(hf, lrn, cell, mode) };
      }
      if (!satisfiesMustUse(formula, check.mustUse, mode)) return { cell, pass: false, reason: "missing_function" };
      const actual = valueAt(hf, lrn, cell);
      const expected = valueAt(hf, sol, cell);
      if (matches(expected, actual)) return { cell, pass: true };
      const shown = { expected: shownAt(hf, sol, cell, mode), actual: shownAt(hf, lrn, cell, mode) };
      return isCellError(actual)
        ? { cell, pass: false, reason: "excel_error", errorType: actual.type, ...shown }
        : { cell, pass: false, reason: "wrong_value", ...shown };
    });
    const baseFailure = baseChecks.find((c) => !c.pass);
    if (baseFailure) return failure(baseFailure, baseChecks, { task, hf, lrn, mode });

    // Rule 4: swap trials, then the extra `cases` datasets, each on top of the lesson data.
    const datasets: { inputs: Case; trial?: number; caseIndex?: number }[] = [
      ...swapTrialInputs(task, options).map((trialInputs, i) => ({ inputs: trialInputs, trial: i + 1 })),
      ...(check.cases ?? []).map((caseInputs, i) => ({ inputs: caseInputs, caseIndex: i + 1 })),
    ];
    const failures = new Map<string, CellCheck & { inputs: Case }>();
    const touched = new Set<string>();
    for (const dataset of datasets) {
      applyValues(hf, [sol, lrn], { ...baseValues(task, touched), ...dataset.inputs }, mode);
      for (const cell of Object.keys(dataset.inputs)) touched.add(cell);
      for (const cell of gradedCells) {
        if (failures.has(cell)) continue;
        const actual = valueAt(hf, lrn, cell);
        if (matches(valueAt(hf, sol, cell), actual)) continue;
        failures.set(cell, {
          cell,
          pass: false,
          reason: "hardcoded",
          trial: dataset.trial,
          caseIndex: dataset.caseIndex,
          expected: shownAt(hf, sol, cell, mode),
          actual: shownAt(hf, lrn, cell, mode),
          errorType: isCellError(actual) ? actual.type : undefined,
          inputs: dataset.inputs,
        });
      }
    }
    const cells = gradedCells.map((cell): CellCheck => failures.get(cell) ?? { cell, pass: true });
    const firstFailure = gradedCells.map((cell) => failures.get(cell)).find((f) => f !== undefined);
    if (firstFailure) return failure(firstFailure, cells, { task, hf, lrn, mode, trialInputs: firstFailure.inputs });
    return {
      pass: true,
      messageKa: "სწორია! ფორმულა სხვა მონაცემებზეც სწორად მუშაობს.",
      cells: cells.map(publicGrade),
    };
  } finally {
    hf.destroy();
  }
}

/** Drops keys whose value is undefined, so results stay tidy (and compare cleanly in tests). */
function compact<T extends object>(value: T): T {
  return Object.fromEntries(Object.entries(value).filter(([, v]) => v !== undefined)) as T;
}

function publicGrade(check: CellCheck): CellGrade {
  const { cell, pass, reason, expected, actual, errorType, trial, caseIndex } = check;
  return compact({ cell, pass, reason, expected, actual, errorType, trial, caseIndex });
}

interface FailureContext {
  readonly task: GradableTask;
  readonly hf: HyperFormula;
  readonly lrn: number;
  readonly mode: FormatMode;
  readonly trialInputs?: Case;
}

function failure(check: CellCheck, all: CellCheck[], ctx: FailureContext): GradeFail {
  const reason = check.reason!;
  return compact<GradeFail>({
    pass: false,
    reason,
    cell: check.cell,
    expected: check.expected,
    actual: check.actual,
    errorType: check.errorType,
    mustUse: reason === "missing_function" ? ctx.task.check.mustUse : undefined,
    trial: check.trial,
    caseIndex: check.caseIndex,
    trialInputs: reason === "hardcoded" ? ctx.trialInputs : undefined,
    messageKa: failureMessageKa(check, ctx),
    cells: all.map(publicGrade),
  });
}

function failureMessageKa(check: CellCheck, ctx: FailureContext): string {
  const { cell } = check;
  switch (check.reason) {
    case "empty":
      return `${cell}: უჯრა ცარიელია. ჩაწერეთ მასში ფორმულა.`;
    case "not_formula":
      return check.spilled
        ? `${cell}: ეს მნიშვნელობა სხვა უჯრის ფორმულამ ავტომატურად გადმოიტანა. ამ დავალებაში თითოეულ უჯრაში საკუთარი ფორმულა ჩაწერეთ (მაგ. Ctrl+D-ით).`
        : `${cell}: ჩაწერეთ ფორმულა და არა მზა რიცხვი. ფორმულა იწყება „=“ ნიშნით.`;
    case "missing_function":
      return `${cell}: ამ დავალებაში გამოიყენეთ ${describeMustUseKa(ctx.task.check.mustUse ?? [])}.`;
    case "excel_error": {
      const value = valueAt(ctx.hf, ctx.lrn, cell);
      const detail = isCellError(value) ? value.detail : "";
      const { code, messageKa } = explainError(check.errorType!, detail, ctx.mode);
      return `${cell}: შეცდომა ${code}. ${messageKa}`;
    }
    case "wrong_value":
      return `${cell}: მიღებული შედეგი (${check.actual}) სწორ პასუხს არ ემთხვევა. გადაამოწმეთ ფორმულა.`;
    case "hardcoded":
      return check.caseIndex !== undefined
        ? `${cell}: ფორმულა ყველა შემთხვევაში სწორ შედეგს არ იძლევა (მაგ. ზღვრულ მნიშვნელობაზე ან განსხვავებულ მონაცემებზე). გადაამოწმეთ პირობები და უჯრების მისამართები.`
        : `${cell}: ფორმულა სწორ პასუხს მხოლოდ ამ კონკრეტულ რიცხვებზე იძლევა. რიცხვების ნაცვლად გამოიყენეთ უჯრების მისამართები (მაგ. B2), რომ მონაცემების შეცვლისას შედეგიც განახლდეს.`;
    default:
      return "";
  }
}
