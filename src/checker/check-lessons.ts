/**
 * The lesson checker's logic (no file system access, so it is easy to test).
 * `scripts/check-lessons.ts` reads the files and prints the report.
 *
 * Every lesson must pass before it is published (CLAUDE.md, build plan §9, SCHEMA.md).
 * Errors block publishing; warnings are worth reading.
 */
import { expandRange, expandRanges, formatCellAddress, parseCellAddress } from "../engine/addresses";
import { unsupportedFunctionTip } from "../engine/errors";
import { lessonTexts, scanPlaceholders } from "../engine/facts";
import { enToKa } from "../engine/format";
import {
  DEFAULT_TOLERANCE,
  applyValues,
  baseValues,
  buildSolutionEngine,
  gradeTask,
  solutionFormulas,
  swapTrialInputs,
  valuesMatch,
  type GradeOptions,
} from "../engine/grading";
import { hfAddressOf, isValidFormula, supportedFunctionNames, toCellValue } from "../engine/hyperformula";
import {
  isDateCell,
  parseFactsJson,
  parseLessonJson,
  tasksOf,
  type Case,
  type Fact,
  type GradableTask,
  type LessonFile,
  type SheetCell,
} from "../engine/schema";
import { editableCellSet } from "../engine/sheet-data";
import { functionNames, normalizeFunctionName, tokenizeFormula, type Token } from "../engine/tokenizer";
import { isCellError, type CellValue } from "../engine/types";

export type IssueLevel = "error" | "warning";

export interface CheckIssue {
  readonly level: IssueLevel;
  readonly message: string;
}

export interface SourceFile {
  /** Path shown in the report. */
  readonly path: string;
  /** The lesson folder it came from: ids and orders must be unique within a folder. */
  readonly folder: string;
  readonly text: string;
  /** True for the public app repo's own lessons: paid lessons must never be there. */
  readonly isPublic?: boolean;
}

export interface FileReport {
  readonly path: string;
  readonly folder: string;
  readonly id?: string;
  readonly kind?: LessonFile["kind"];
  readonly issues: CheckIssue[];
}

export interface FactsReport {
  readonly path: string | null;
  readonly count: number;
  readonly issues: CheckIssue[];
}

export interface CheckReport {
  readonly files: FileReport[];
  readonly facts: FactsReport;
  readonly errors: number;
  readonly warnings: number;
}

export interface CheckOptions {
  /** Unverified or stale facts used by the checked lessons become errors (needed to publish). */
  readonly strict?: boolean;
  /** "Today" for the 12-month facts rule. Default: now. */
  readonly today?: Date;
  /** Swap-test seed and trial count; defaults match what learners get. */
  readonly seed?: number;
  readonly trials?: number;
}

const MAX_THEORY_WORDS = 150;
const FACT_MAX_AGE_MONTHS = 12;
const LOOKUP_FUNCTIONS = ["VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP", "MATCH", "INDEX"];
const ALLOWED_TEXT_FORMAT_RE = /^(?:0+|dd\.mm\.yyyy)$/i;
const GEORGIAN_RE = /[Ⴀ-ჿᲐ-Ჿ]/;
const MTAVRULI_RE = /[Ა-Ჿ]/;

const error = (message: string): CheckIssue => ({ level: "error", message });
const warning = (message: string): CheckIssue => ({ level: "warning", message });
const hasErrors = (issues: readonly CheckIssue[]) => issues.some((issue) => issue.level === "error");

function describeCell(value: SheetCell | undefined): string {
  if (value === undefined || value === null) return "is empty";
  if (isDateCell(value)) return `holds the date ${value.date}`;
  if (typeof value === "string") return value.startsWith("=") ? `holds a formula (${value})` : `holds text ${JSON.stringify(value)}`;
  return `holds ${value}`;
}

function showValue(value: CellValue): string {
  if (isCellError(value)) return value.code;
  return typeof value === "string" ? JSON.stringify(value) : String(value);
}

function showInputs(inputs: Case): string {
  return Object.entries(inputs)
    .map(([cell, v]) => `${cell}=${isDateCell(v) ? v.date : JSON.stringify(v)}`)
    .join(", ");
}

/** `=1960`, `="text"`, `=TRUE()`: a solution value typed in as a constant formula. */
function constantFormula(value: CellValue): string | null {
  if (typeof value === "number") return `=${String(value)}`;
  if (typeof value === "string") return value.includes('"') ? null : `="${value}"`;
  if (typeof value === "boolean") return value ? "=TRUE()" : "=FALSE()";
  return null;
}

function lessonValue(task: GradableTask, cell: string): SheetCell | undefined {
  const { row, col } = parseCellAddress(cell);
  return task.sheet.cells[row]?.[col];
}

/** Every pre-filled formula of the sheet, with its cell. */
function dataFormulas(task: GradableTask): { cell: string; formula: string }[] {
  const out: { cell: string; formula: string }[] = [];
  task.sheet.cells.forEach((row, r) =>
    row.forEach((value, c) => {
      if (typeof value === "string" && value.startsWith("=")) out.push({ cell: formatCellAddress(r, c), formula: value });
    }),
  );
  return out;
}

// ------------------------------------------------------------------ formula lint

/** The second argument of every TEXT(…) call: its literal text, or null if it is not a literal. */
function textFormats(tokens: readonly Token[]): (string | null)[] {
  const formats: (string | null)[] = [];
  tokens.forEach((token, i) => {
    if (token.type !== "function" || normalizeFunctionName(token.text) !== "TEXT") return;
    let depth = 0;
    let arg = 0;
    const second: Token[] = [];
    for (let j = i + 1; j < tokens.length; j++) {
      const t = tokens[j]!;
      if (t.type === "lparen" || t.type === "lbrace") depth++;
      else if (t.type === "rparen" || t.type === "rbrace") {
        depth--;
        if (depth === 0) break;
      } else if (t.type === "argSeparator" && depth === 1) {
        arg++;
        continue;
      }
      if (arg === 1 && depth >= 1 && t.type !== "whitespace") second.push(t);
    }
    const only = second.length === 1 ? second[0] : undefined;
    formats.push(only?.type === "string" ? only.text.slice(1, -1) : null);
  });
  return formats;
}

/** Problems HyperFormula or the format converter would have with a formula (SCHEMA.md rule 8). */
function lintFormula(formula: string, where: string, task: GradableTask): CheckIssue[] {
  const issues: CheckIssue[] = [];
  const tokens = tokenizeFormula(formula, "en");
  for (const token of tokens) {
    if (token.type === "string" && token.text.slice(1, -1).includes('""')) {
      issues.push(error(`${where}: "" inside a text string is not supported (${token.text})`));
    }
    if (token.type === "number" && /e/i.test(token.text)) {
      issues.push(error(`${where}: write ${token.text} as a plain number (no 1E+3-style numbers)`));
    }
  }
  if (!isValidFormula(formula, "en")) {
    if (issues.length > 0) return issues;
    const hint = isValidFormula(formula, "ka") ? " (it looks like Georgian format; lesson files use English format)" : "";
    return [error(`${where}: ${formula} is not a valid English-format formula${hint}`)];
  }
  if (!isValidFormula(enToKa(formula), "ka")) {
    issues.push(error(`${where}: ${formula} cannot be converted to Georgian format (${enToKa(formula)})`));
  }
  const supported = supportedFunctionNames();
  for (const name of functionNames(formula, "en")) {
    if (name === "NOW") {
      issues.push(error(`${where}: NOW() is not allowed (results change every second); use TODAY() with sheet.today`));
    } else if (name === "TODAY" && !task.sheet.today) {
      issues.push(error(`${where} uses TODAY(): add sheet.today (the date TODAY() returns) and state it in the task text`));
    } else if (!supported.has(name)) {
      const tip = unsupportedFunctionTip(name);
      issues.push(error(`${where}: ${name} is not supported by our engine${tip ? ` (tip: ${tip})` : ""}`));
    }
  }
  for (const format of textFormats(tokens)) {
    if (format === null || !ALLOWED_TEXT_FORMAT_RE.test(format)) {
      issues.push(
        error(
          `${where}: TEXT format ${format === null ? "must be a literal" : JSON.stringify(format)} is not allowed; use zero padding ("00000000000"), "0" or "dd.mm.yyyy"`,
        ),
      );
    }
  }
  return issues;
}

/** The cells and ranges a formula refers to, e.g. ["B2", "$E$1", "F2:H6"]. */
function references(formula: string): string[] {
  const tokens = tokenizeFormula(formula, "en");
  const refs: string[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]!;
    if (token.type !== "cell") continue;
    const end = tokens[i + 1]?.type === "colon" && tokens[i + 2]?.type === "cell" ? tokens[i + 2]! : undefined;
    refs.push(end ? `${token.text}:${end.text}` : token.text);
    if (end) i += 2;
  }
  return refs;
}

/** References in the filled solution that point to empty cells (usually a missing $). */
function emptyReferences(task: GradableTask, formulas: Record<string, string>, at: (s: string) => string): CheckIssue[] {
  const graded = new Set(Object.keys(formulas));
  const isEmpty = (cell: string) => {
    if (graded.has(cell)) return false;
    const value = lessonValue(task, cell);
    return value === undefined || value === null;
  };
  const reported = new Set<string>();
  const issues: CheckIssue[] = [];
  for (const [gradedCell, formula] of Object.entries(formulas)) {
    for (const ref of references(formula)) {
      const plain = ref.replaceAll("$", "");
      if (reported.has(plain) || !expandRange(plain).every(isEmpty)) continue;
      reported.add(plain);
      issues.push(
        error(`${at("check.solution")} filled into ${gradedCell} (${formula}) refers to ${ref}, which is empty; often a missing $ (e.g. $E$1)`),
      );
    }
  }
  return issues;
}

/** Numbers the solution reads directly that neither `vary` nor any case changes (SCHEMA.md rule 5). */
function unvariedInputs(task: GradableTask, formulas: Record<string, string>, at: (s: string) => string): CheckIssue[] {
  const covered = new Set([
    ...expandRanges(task.check.vary),
    ...(task.check.cases ?? []).flatMap((c) => Object.keys(c)),
    ...Object.keys(formulas),
  ]);
  const missing = new Set<string>();
  for (const formula of Object.values(formulas)) {
    for (const ref of references(formula)) {
      for (const cell of expandRange(ref.replaceAll("$", ""))) {
        const value = lessonValue(task, cell);
        if ((typeof value === "number" || isDateCell(value)) && !covered.has(cell)) missing.add(cell);
      }
    }
  }
  if (missing.size === 0) return [];
  const list = [...missing].slice(0, 6).join(", ") + (missing.size > 6 ? ", …" : "");
  return [warning(`${at("check.vary")}: the solution reads ${list}, but neither vary nor any case changes them; a learner could type those numbers in`)];
}

// ------------------------------------------------------------------ one task

function checkStructure(task: GradableTask, at: (s: string) => string): CheckIssue[] {
  const issues: CheckIssue[] = [];
  const graded = new Set(expandRange(task.check.cells));
  const editable = editableCellSet(task.sheet);
  for (const cell of graded) {
    if (!editable.has(cell)) issues.push(error(`${at("check.cells")}: ${cell} is graded but not editable (add it to sheet.editable)`));
  }
  for (const cell of expandRanges(task.check.vary)) {
    const value = lessonValue(task, cell);
    if (typeof value !== "number" && !isDateCell(value)) {
      issues.push(error(`${at("check.vary")}: ${cell} must hold a number or a date, but it ${describeCell(value)}`));
    }
    if (graded.has(cell)) issues.push(error(`${at("check.vary")}: ${cell} is also a graded cell`));
    else if (editable.has(cell)) issues.push(error(`${at("check.vary")}: ${cell} is editable; vary cells must be locked lesson data`));
  }
  (task.check.cases ?? []).forEach((c, i) => {
    for (const cell of Object.keys(c)) {
      if (graded.has(cell) || editable.has(cell)) {
        issues.push(error(`${at(`check.cases[${i}]`)}: ${cell} is a graded or editable cell; cases may only change lesson data`));
      }
    }
  });
  if (task.check.vary.length === 0 && !(task.check.cases ?? []).length) {
    issues.push(warning(`${at("check.vary")} is empty and there are no cases: a typed-in answer may pass (SCHEMA.md rule 7)`));
  }
  return issues;
}

function checkMustUse(task: GradableTask, at: (s: string) => string): CheckIssue[] {
  const mustUse = task.check.mustUse ?? [];
  if (mustUse.length === 0) return [];
  const issues: CheckIssue[] = [];
  const supported = supportedFunctionNames();
  for (const name of mustUse.flat()) {
    if (!supported.has(name.toUpperCase())) issues.push(error(`${at("check.mustUse")}: "${name}" is not a function our engine knows`));
  }
  const used = new Set(functionNames(task.check.solution, "en"));
  const satisfied = mustUse.some((alt) => (Array.isArray(alt) ? alt : [alt]).every((n) => used.has(n.toUpperCase())));
  if (!satisfied) issues.push(error(`${at("check.mustUse")}: the solution satisfies none of the alternatives`));
  return issues;
}

/** Content-design rules from SCHEMA.md 5–6 and 8 that are advice, not hard errors. */
function checkDesign(task: GradableTask, base: readonly CellValue[], at: (s: string) => string): CheckIssue[] {
  const issues: CheckIssue[] = [];
  const used = functionNames(task.check.solution, "en");
  const noCases = !(task.check.cases ?? []).length;
  if (typeof task.check.compareDecimals === "number" && used.includes("ROUND")) {
    issues.push(warning(`${at("check.compareDecimals")} is set on a task whose solution uses ROUND: rounding would no longer be checked`));
  }
  if (noCases && used.some((f) => f === "IF" || f === "IFS")) {
    issues.push(warning(`${at("check.cases")}: IF/IFS tasks need cases at the boundary and just below it (SCHEMA.md rule 5)`));
  }
  if (noCases && used.some((f) => f === "IFERROR" || f === "IFNA")) {
    issues.push(warning(`${at("check.cases")}: IFERROR tasks need a case that produces the error (SCHEMA.md rule 5)`));
  }
  if (noCases && used.some((f) => LOOKUP_FUNCTIONS.includes(f))) {
    issues.push(warning(`${at("check.cases")}: lookup tasks need cases that put other existing keys into the key cell (SCHEMA.md rule 6)`));
  }
  if (used.some((f) => f === "PROPER" || f === "UPPER")) {
    const georgianData = task.sheet.cells.flat().some((v) => typeof v === "string" && !v.startsWith("=") && GEORGIAN_RE.test(v));
    const mtavruli = base.some((v) => typeof v === "string" && MTAVRULI_RE.test(v));
    if (georgianData || mtavruli) {
      issues.push(warning(`${at("check.solution")}: PROPER/UPPER turn Georgian letters into Mtavruli; use them only on Latin-script text (SCHEMA.md rule 8)`));
    }
  }
  if (base.some((v) => typeof v === "string" && /\d[.,]\d/.test(v))) {
    issues.push(warning(`${at("check.solution")}: the answer is text built from a decimal number; Excel's separator may differ (SCHEMA.md rule 8)`));
  }
  return issues;
}

/** Runs the solution on the lesson data, each vary cell alone, the swap trials and the cases. */
function checkSolutionValues(
  task: GradableTask,
  at: (s: string) => string,
  grading: GradeOptions,
): { issues: CheckIssue[]; base: CellValue[] | null } {
  const issues: CheckIssue[] = [];
  const graded = expandRange(task.check.cells);
  const same = (a: CellValue, b: CellValue) =>
    valuesMatch(a, b, task.check.tolerance ?? DEFAULT_TOLERANCE, task.check.compareDecimals);
  const hf = buildSolutionEngine(task, "en");
  const read = () => graded.map((cell) => toCellValue(hf.getCellValue(hfAddressOf(cell))));
  const changedFromBase = (values: CellValue[], base: CellValue[]) => values.some((v, i) => !same(base[i] ?? null, v));
  try {
    const base = read();
    base.forEach((value, i) => {
      if (isCellError(value)) {
        issues.push(error(`${at("check.solution")} gives ${value.code} in ${graded[i]} on the lesson data${value.detail ? ` (${value.detail})` : ""}`));
      }
    });
    for (const { cell, formula } of dataFormulas(task)) {
      const shown = toCellValue(hf.getCellValue(hfAddressOf(cell)));
      if (isCellError(shown)) issues.push(warning(`${at("sheet.cells")}: pre-filled ${cell} ${formula} shows ${shown.code}`));
    }
    if (hasErrors(issues)) return { issues, base: null };

    // With an empty vary the trials change nothing, so they are not datasets.
    const trials = swapTrialInputs(task, grading).filter((trial) => Object.keys(trial).length > 0);
    // Each vary cell on its own must change the result (SCHEMA.md rule 5).
    for (const cell of Object.keys(trials[0] ?? {})) {
      let changes = false;
      for (const trial of trials) {
        applyValues(hf, [0], { [cell]: trial[cell]! }, "en");
        changes = changedFromBase(read(), base);
        applyValues(hf, [0], baseValues(task, [cell]), "en");
        if (changes) break;
      }
      if (!changes) {
        issues.push(error(`${at("check.vary")}: changing ${cell} on its own never changes the result; remove it from vary or check the solution uses it`));
      }
    }

    // Swap trials (all vary cells together), then every case, each on top of the lesson data.
    const datasets = [
      ...trials.map((inputs, i) => ({ inputs, label: `swap trial ${i + 1}`, isCase: false })),
      ...(task.check.cases ?? []).map((inputs, i) => ({ inputs, label: `case ${i + 1}`, isCase: true })),
    ];
    const everChanged = new Set<string>();
    const touched = new Set<string>();
    for (const { inputs, label, isCase } of datasets) {
      applyValues(hf, [0], { ...baseValues(task, touched), ...inputs }, "en");
      for (const cell of Object.keys(inputs)) touched.add(cell);
      const values = read();
      values.forEach((value, i) => {
        const cell = graded[i]!;
        if (isCellError(value)) {
          issues.push(error(`${at("check.solution")} gives ${value.code} in ${cell} in ${label} (${showInputs(inputs)})`));
        } else if (!same(base[i] ?? null, value)) {
          everChanged.add(cell);
        }
      });
      if (isCase && !changedFromBase(values, base)) {
        issues.push(warning(`${at("check.cases")}: ${label} (${showInputs(inputs)}) gives the same result as the lesson data, so it tests nothing`));
      }
    }
    if (datasets.length > 0) {
      for (const [i, cell] of graded.entries()) {
        if (!everChanged.has(cell)) {
          issues.push(error(`swap test: no trial or case changes the result in ${cell} (${showValue(base[i] ?? null)}), so a typed-in answer would pass`));
        }
      }
    }
    return { issues, base };
  } finally {
    hf.destroy();
  }
}

function checkSelfGrading(task: GradableTask, base: readonly CellValue[], at: (s: string) => string, grading: GradeOptions): CheckIssue[] {
  const issues: CheckIssue[] = [];
  const formulas = solutionFormulas(task);

  const en = gradeTask(task, formulas, "en", grading);
  if (!en.pass) {
    issues.push(error(`${at("check.solution")} fails its own grading in English format: ${en.reason} in ${en.cell} (${en.messageKa})`));
  }
  const kaInputs = Object.fromEntries(Object.entries(formulas).map(([cell, f]) => [cell, enToKa(f)]));
  const ka = gradeTask(task, kaInputs, "ka", grading);
  if (!ka.pass) {
    issues.push(error(`${at("check.solution")} fails its own grading in Georgian format (${kaInputs[ka.cell] ?? ""}): ${ka.reason} in ${ka.cell} (${ka.messageKa})`));
  }

  // Anti-hardcode self-test: the solution's own values typed in as constants must fail.
  const constants: Record<string, string> = {};
  expandRange(task.check.cells).forEach((cell, i) => {
    const constant = constantFormula(base[i] ?? null);
    if (constant !== null) constants[cell] = constant;
  });
  if (Object.keys(constants).length === 0) return issues;
  // With vary or cases, test them alone (mustUse would hide a weak swap test).
  const protectedByData = task.check.vary.length > 0 || (task.check.cases ?? []).length > 0;
  const probe = protectedByData ? { ...task, check: { ...task.check, mustUse: [] } } : task;
  if (gradeTask(probe, constants, "en", grading).pass) {
    const example = Object.entries(constants).slice(0, 2).map(([cell, f]) => `${cell}: ${f}`).join(", ");
    issues.push(
      error(
        `anti-hardcode self-test: typing the solution's values as constants (${example}) passes grading; put every number the answer depends on in ${at("check.vary")} or add cases`,
      ),
    );
  }
  return issues;
}

/** Each `mustFail` formula (a typical wrong answer) must fail grading. */
function checkMustFail(task: GradableTask, at: (s: string) => string, grading: GradeOptions): CheckIssue[] {
  const issues: CheckIssue[] = [];
  (task.check.mustFail ?? []).forEach((formula, i) => {
    const where = at(`check.mustFail[${i}]`);
    if (!isValidFormula(formula, "en")) {
      issues.push(error(`${where}: ${formula} is not a valid English-format formula, so it tests nothing`));
      return;
    }
    const wrong = solutionFormulas({ ...task, check: { ...task.check, solution: formula } });
    if (gradeTask(task, wrong, "en", grading).pass) {
      issues.push(error(`${where}: the wrong answer ${formula} passes grading; add vary cells, cases or mustUse that reject it`));
    }
  });
  return issues;
}

function checkTask(task: GradableTask, prefix: string, options: CheckOptions): CheckIssue[] {
  const at = (path: string) => `${prefix}${path}`;
  const grading: GradeOptions = { seed: options.seed, trials: options.trials };
  const issues = [...checkStructure(task, at)];
  for (const { cell, formula } of dataFormulas(task)) issues.push(...lintFormula(formula, `${at("sheet.cells")} ${cell}`, task));

  const solutionLint = lintFormula(task.check.solution, at("check.solution"), task);
  issues.push(...solutionLint, ...checkMustUse(task, at));
  if (hasErrors(solutionLint)) return issues;

  const filled = solutionFormulas(task);
  issues.push(...emptyReferences(task, filled, at), ...unvariedInputs(task, filled, at));
  const values = checkSolutionValues(task, at, grading);
  issues.push(...values.issues);
  if (!values.base) return issues;
  issues.push(...checkDesign(task, values.base, at));
  if (!hasErrors(issues)) issues.push(...checkSelfGrading(task, values.base, at, grading), ...checkMustFail(task, at, grading));
  return issues;
}

// ------------------------------------------------------------------ facts

function factIdsUsedBy(file: LessonFile): Set<string> {
  const ids = new Set<string>(file.kind === "lesson" ? file.facts : file.tasks.flatMap((t) => t.facts));
  for (const { text } of lessonTexts(file)) for (const id of scanPlaceholders(text).factIds) ids.add(id);
  return ids;
}

function checkFactReferences(file: LessonFile, facts: ReadonlyMap<string, Fact> | null): CheckIssue[] {
  const issues: CheckIssue[] = [];
  const declared = file.kind === "lesson" ? [file.facts] : file.tasks.map((task) => task.facts);
  const declaredIds = new Set(declared.flat());
  const prefixes = file.kind === "lesson" ? [""] : file.tasks.map((_, i) => `tasks[${i}].`);
  declared.forEach((ids, i) => {
    for (const id of ids) {
      if (facts && !facts.has(id)) issues.push(error(`${prefixes[i]}facts: unknown fact "${id}" (not in facts.json)`));
    }
  });
  for (const { where, text } of lessonTexts(file)) {
    const { factIds, malformed } = scanPlaceholders(text);
    for (const bad of malformed) issues.push(error(`${where}: malformed placeholder ${bad} (use {{fact:fact-id}})`));
    for (const id of factIds) {
      const fact = facts?.get(id);
      if (facts && !fact) issues.push(error(`${where}: {{fact:${id}}} refers to an unknown fact`));
      else if (!declaredIds.has(id)) issues.push(warning(`${where}: {{fact:${id}}} is used but "${id}" is not listed in facts`));
      if (fact && fact.display === undefined && typeof fact.value === "number" && !Number.isInteger(fact.value)) {
        issues.push(warning(`${where}: {{fact:${id}}} will show as ${fact.value}; give the fact a display text such as "18%"`));
      }
    }
  }
  if (!facts && declaredIds.size > 0) {
    issues.push(error("facts are referenced but no valid facts.json was found (use --facts <path>)"));
  }
  return issues;
}

function addMonthsUtc(date: Date, months: number): Date {
  const result = new Date(date.getTime());
  result.setUTCMonth(result.getUTCMonth() + months);
  return result;
}

function checkFacts(facts: readonly Fact[], usedBy: ReadonlyMap<string, string[]>, options: CheckOptions): CheckIssue[] {
  const issues: CheckIssue[] = [];
  const now = options.today ?? new Date();
  const today = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  const seen = new Set<string>();
  for (const fact of facts) {
    if (seen.has(fact.id)) issues.push(error(`${fact.id}: duplicate id`));
    seen.add(fact.id);
    const users = usedBy.get(fact.id) ?? [];
    // --strict fails on facts that checked lessons use (SCHEMA.md: needed to publish them).
    const level: IssueLevel = options.strict && users.length > 0 ? "error" : "warning";
    const usage = users.length > 0 ? ` (used by ${users.join(", ")})` : "";
    if (fact.checked === null) {
      issues.push({ level, message: `${fact.id}: not verified yet (checked: null)${usage}` });
      continue;
    }
    const checked = new Date(`${fact.checked}T00:00:00Z`);
    if (checked > today) {
      issues.push(warning(`${fact.id}: checked date ${fact.checked} is in the future`));
    } else if (addMonthsUtc(checked, FACT_MAX_AGE_MONTHS) < today) {
      issues.push({ level, message: `${fact.id}: last checked ${fact.checked}, more than ${FACT_MAX_AGE_MONTHS} months ago${usage}` });
    }
  }
  return issues;
}

// ------------------------------------------------------------------ files

function checkFile(file: LessonFile, source: SourceFile, facts: ReadonlyMap<string, Fact> | null, options: CheckOptions): CheckIssue[] {
  const issues: CheckIssue[] = [];
  if (source.isPublic && !file.free) {
    issues.push(error("paid lesson (free: false) in the public app repo: paid content must stay in the private content repo"));
  }
  tasksOf(file).forEach((task, i) => {
    issues.push(...checkTask(task, file.kind === "mock" ? `tasks[${i}].` : "", options));
  });
  issues.push(...checkFactReferences(file, facts));
  if (file.kind === "lesson") {
    const words = file.theory.split(/\s+/).filter((word) => /[\p{L}\p{N}]/u.test(word)).length;
    if (words > MAX_THEORY_WORDS) issues.push(warning(`theory has ${words} words (SCHEMA.md: at most ${MAX_THEORY_WORDS})`));
    if (file.hints.length !== 3) issues.push(warning(`hints: ${file.hints.length} hint(s); SCHEMA.md asks for exactly 3`));
  }
  return issues;
}

function fileName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path;
}

/** Checks lesson files (and the facts registry) and returns a report. Never throws on bad content. */
export function checkLessons(
  files: readonly SourceFile[],
  factsFile: { path: string; text: string } | null,
  options: CheckOptions = {},
): CheckReport {
  let facts: Map<string, Fact> | null = null;
  let factList: Fact[] = [];
  const factsIssues: CheckIssue[] = [];
  if (factsFile) {
    const parsed = parseFactsJson(factsFile.text);
    if (parsed.ok) {
      factList = parsed.value;
      facts = new Map(factList.map((fact) => [fact.id, fact]));
    } else {
      factsIssues.push(...parsed.errors.map(error));
    }
  }

  const usedBy = new Map<string, string[]>();
  const reports: (FileReport & { order?: number })[] = files.map((source) => {
    const parsed = parseLessonJson(source.text);
    if (!parsed.ok) return { path: source.path, folder: source.folder, issues: parsed.errors.map(error) };
    const file = parsed.value;
    for (const id of factIdsUsedBy(file)) usedBy.set(id, [...(usedBy.get(id) ?? []), file.id]);
    let issues: CheckIssue[];
    try {
      issues = checkFile(file, source, facts, options);
    } catch (e) {
      issues = [error(`checker crashed on this file: ${(e as Error).message}`)];
    }
    const prefix = /^(\d+)-/.exec(fileName(source.path))?.[1];
    if (prefix !== undefined && Number(prefix) !== file.order) {
      issues.push(warning(`file name starts with ${prefix} but order is ${file.order}`));
    }
    return { path: source.path, folder: source.folder, id: file.id, kind: file.kind, order: file.order, issues };
  });

  // Ids and orders must be unique within each folder.
  for (const report of reports) {
    if (report.id === undefined) continue;
    for (const other of reports) {
      if (other === report || other.folder !== report.folder || other.id === undefined) continue;
      if (other.id === report.id) report.issues.push(error(`id "${report.id}" is also used by ${fileName(other.path)}`));
      if (other.order === report.order) report.issues.push(error(`order ${report.order} is also used by ${fileName(other.path)}`));
    }
  }

  if (facts) factsIssues.push(...checkFacts(factList, usedBy, options));
  const all = [...reports.flatMap((r) => r.issues), ...factsIssues];
  return {
    files: reports.map(({ order: _order, ...report }) => report),
    facts: { path: factsFile?.path ?? null, count: factList.length, issues: factsIssues },
    errors: all.filter((issue) => issue.level === "error").length,
    warnings: all.filter((issue) => issue.level === "warning").length,
  };
}
