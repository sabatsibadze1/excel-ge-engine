/**
 * The lesson and facts format, exactly as in the canonical contract `content/SCHEMA.md`
 * (in the content repo). If you change one, change the other.
 *
 * Use `parseLessonFile` / `parseFacts` (or their `…Json` variants for raw text) to get either a
 * typed value or a list of friendly error messages such as
 * `check.cells: expected a cell like "C2" or a range like "C2:C6", got "c2-c6"`.
 */
import * as z from "zod";

const CELL = "[A-Z]{1,3}[1-9][0-9]{0,6}";
const CELL_RE = new RegExp(`^${CELL}$`);
const RANGE_RE = new RegExp(`^${CELL}(?::${CELL})?$`);
const KEBAB_CASE_RE = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const FUNCTION_NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]*$/;

const nonEmptyText = z.string().trim().min(1, { error: "must not be empty" });

function isRealDate(value: string): boolean {
  const date = new Date(`${value}T00:00:00Z`);
  return !Number.isNaN(date.getTime()) && date.toISOString().startsWith(value);
}

/** A calendar date written as "YYYY-MM-DD". */
export const IsoDateSchema = z
  .string()
  .regex(ISO_DATE_RE, { error: 'expected a date like "2026-11-02"' })
  .refine(isRealDate, { error: "is not a real calendar date" });

/** A single cell address: "C2". Capital letters, no `$`. */
export const CellAddressSchema = z.string().regex(CELL_RE, {
  error: (issue) => `expected a cell like "C2", got ${JSON.stringify(issue.input)}`,
});

/** A cell or a range: "C2" or "C2:C6". Capital letters, no `$`. */
export const RangeAddressSchema = z.string().regex(RANGE_RE, {
  error: (issue) =>
    `expected a cell like "C2" or a range like "C2:C6", got ${JSON.stringify(issue.input)}`,
});

const IdSchema = z.string().regex(KEBAB_CASE_RE, {
  error: (issue) => `expected a kebab-case id like "payroll-gross-net", got ${JSON.stringify(issue.input)}`,
});

/** A date cell: `{ "date": "2026-02-01" }`. Loaded as a real date, shown as 01.02.2026. */
export const DateCellSchema = z.strictObject({ date: IsoDateSchema });

/**
 * One cell of the starting data (SCHEMA.md "Cell values and addresses"):
 * a number (dot decimal), null (empty), a string starting with "=" (an English-format formula),
 * any other string (always text, even "01001012345"), or a date `{ "date": "YYYY-MM-DD" }`.
 */
export const SheetCellSchema = z.union([z.number(), z.string(), z.null(), DateCellSchema]);

export const SheetSchema = z.strictObject({
  /** Rows of cells starting at A1. */
  cells: z.array(z.array(SheetCellSchema)).min(1, { error: "needs at least one row" }),
  /** Cells or ranges the learner may edit. Everything else is locked. */
  editable: z.array(RangeAddressSchema).min(1, { error: "needs at least one editable cell" }),
  /** The date TODAY() returns in this task (required when a formula uses TODAY()). */
  today: IsoDateSchema.nullable().optional(),
});

/** A function name, or a list of names that must ALL appear (e.g. ["INDEX", "MATCH"]). */
export const MustUseAlternativeSchema = z.union([
  z.string().regex(FUNCTION_NAME_RE, { error: 'expected a function name like "VLOOKUP"' }),
  z
    .array(z.string().regex(FUNCTION_NAME_RE, { error: 'expected a function name like "MATCH"' }))
    .min(1, { error: "a list of functions must not be empty" }),
]);

/** One extra dataset for grading: cell → new value, applied on top of the lesson data. */
export const CaseSchema = z
  .record(CellAddressSchema, SheetCellSchema)
  .refine((data) => Object.keys(data).length > 0, { error: "a case must change at least one cell" });

export const CheckSchema = z.strictObject({
  /** The graded cell or range. */
  cells: RangeAddressSchema,
  /** English-format formula for the top-left graded cell; filled across a range like Ctrl+D. */
  solution: z.string().regex(/^=/, { error: 'must be a formula starting with "="' }),
  /** Cells/ranges holding the numbers (or dates) the answer depends on; changed in the swap test. */
  vary: z.array(RangeAddressSchema),
  /** Alternatives: at least one must be satisfied by the learner's formula. */
  mustUse: z.array(MustUseAlternativeSchema).optional(),
  /** RELATIVE tolerance for numbers: |a-b| <= tolerance × max(1, |a|, |b|). Default 1e-9. */
  tolerance: z.number().nonnegative().optional(),
  /**
   * If set, numbers are rounded to this many decimals before comparing (instead of the tolerance):
   * "rounded to tetri or not" are both accepted. Never on a lesson that teaches ROUND.
   */
  compareDecimals: z.number().int().min(0).max(15).nullable().optional(),
  /** Extra datasets graded after the swap trials (boundaries, other lookup keys, error cases). */
  cases: z.array(CaseSchema).optional(),
  /** Typical wrong formulas for the first graded cell (English format); the checker asserts each fails. */
  mustFail: z.array(z.string().regex(/^=/, { error: 'must be a formula starting with "="' })).optional(),
});

export const SourceSchema = z.enum(["ai-draft", "community-test", "expert"]);

export const VerifiedSchema = z.strictObject({
  checker: z.boolean(),
  claudeAudit: IsoDateSchema.nullable(),
  founderPlayed: IsoDateSchema.nullable(),
});

const common = {
  id: IdSchema,
  module: z.number().int().min(1),
  order: z.number().int().min(1),
  free: z.boolean(),
  title: nonEmptyText,
  source: SourceSchema,
  verified: VerifiedSchema,
};

export const LessonSchema = z.strictObject({
  ...common,
  kind: z.literal("lesson"),
  skills: z.array(nonEmptyText),
  theory: nonEmptyText,
  task: nonEmptyText,
  hints: z.array(nonEmptyText).min(1, { error: "needs at least one hint" }),
  sheet: SheetSchema,
  check: CheckSchema,
  facts: z.array(IdSchema),
});

/** One task of a mock test: its own small sheet, no hints, no theory. */
export const TaskSchema = z.strictObject({
  task: nonEmptyText,
  sheet: SheetSchema,
  check: CheckSchema,
  facts: z.array(IdSchema),
});

export const MockTestSchema = z.strictObject({
  ...common,
  kind: z.literal("mock"),
  intro: nonEmptyText,
  minutes: z.number().int().positive(),
  tasks: z.array(TaskSchema).min(1, { error: "needs at least one task" }),
});

export const LessonFileSchema = z.discriminatedUnion("kind", [LessonSchema, MockTestSchema]);

export const FactSchema = z.strictObject({
  id: IdSchema,
  text_ka: nonEmptyText,
  value: z.union([z.number(), z.string(), z.boolean(), z.null()]),
  source: nonEmptyText,
  /** When someone opened the source and confirmed the fact, or null if not yet verified. */
  checked: IsoDateSchema.nullable(),
  checkedBy: z.enum(["ai-web", "founder", "expert"]).nullable(),
  /** How {{fact:id}} is shown in lesson texts, e.g. "18%" (default: the value as written). */
  display: z.string().optional(),
  notes: z.string().optional(),
});

export const FactsFileSchema = z.array(FactSchema);

export type DateCell = z.infer<typeof DateCellSchema>;
export type SheetCell = z.infer<typeof SheetCellSchema>;
export type MustUseAlternative = z.infer<typeof MustUseAlternativeSchema>;
export type Case = z.infer<typeof CaseSchema>;
export type Sheet = z.infer<typeof SheetSchema>;
export type Check = z.infer<typeof CheckSchema>;
export type Lesson = z.infer<typeof LessonSchema>;
export type Task = z.infer<typeof TaskSchema>;
export type MockTest = z.infer<typeof MockTestSchema>;
/** What one file in `content/lessons/` holds. */
export type LessonFile = Lesson | MockTest;
export type Fact = z.infer<typeof FactSchema>;

/** Anything that can be graded: a lesson or one mock-test task. */
export type GradableTask = Pick<Task, "sheet" | "check">;

export type ParseResult<T> = { ok: true; value: T } | { ok: false; errors: string[] };

const KNOWN_KEYS = [
  ...new Set([
    ...Object.keys(LessonSchema.shape),
    ...Object.keys(MockTestSchema.shape),
    ...Object.keys(TaskSchema.shape),
    ...Object.keys(SheetSchema.shape),
    ...Object.keys(CheckSchema.shape),
    ...Object.keys(VerifiedSchema.shape),
    ...Object.keys(DateCellSchema.shape),
    ...Object.keys(FactSchema.shape),
  ]),
];

function formatPath(path: readonly PropertyKey[]): string {
  let out = "";
  for (const key of path) {
    if (typeof key === "number") out += `[${key}]`;
    else out += out ? `.${String(key)}` : String(key);
  }
  return out;
}

function describeIssue(issue: z.core.$ZodIssue): string {
  if (issue.code === "unrecognized_keys") {
    return issue.keys
      .map((key) => {
        const match = KNOWN_KEYS.find((k) => k.toLowerCase() === key.toLowerCase());
        return match ? `unknown field "${key}" (did you mean "${match}"?)` : `unknown field "${key}"`;
      })
      .join("; ");
  }
  if (issue.code === "invalid_type" && issue.input === undefined) {
    return `is required (expected ${issue.expected})`;
  }
  return issue.message;
}

/** Turns zod issues into short "path: problem" lines. */
export function formatIssues(issues: readonly z.core.$ZodIssue[]): string[] {
  return issues.map((issue) => {
    const where = formatPath(issue.path);
    const what = describeIssue(issue);
    return where ? `${where}: ${what}` : what;
  });
}

function parseWith<T>(schema: z.ZodType<T>, data: unknown): ParseResult<T> {
  const result = schema.safeParse(data, { reportInput: true });
  return result.success
    ? { ok: true, value: result.data }
    : { ok: false, errors: formatIssues(result.error.issues) };
}

/** Validates one lesson or mock test (already parsed from JSON). */
export function parseLessonFile(data: unknown): ParseResult<LessonFile> {
  const kind = (data as { kind?: unknown } | null)?.kind;
  if (typeof data !== "object" || data === null || Array.isArray(data)) {
    return { ok: false, errors: ["expected a JSON object"] };
  }
  if (kind === "lesson") return parseWith(LessonSchema, data);
  if (kind === "mock") return parseWith(MockTestSchema, data);
  return { ok: false, errors: [`kind: expected "lesson" or "mock", got ${JSON.stringify(kind)}`] };
}

/** Validates the facts registry (already parsed from JSON). */
export function parseFacts(data: unknown): ParseResult<Fact[]> {
  return parseWith(FactsFileSchema, data);
}

function parseJsonText(text: string): ParseResult<unknown> {
  try {
    return { ok: true, value: JSON.parse(text) as unknown };
  } catch (error) {
    return { ok: false, errors: [`invalid JSON: ${(error as Error).message}`] };
  }
}

/** Parses and validates the text of a lesson file. */
export function parseLessonJson(text: string): ParseResult<LessonFile> {
  const json = parseJsonText(text);
  return json.ok ? parseLessonFile(json.value) : json;
}

/** Parses and validates the text of `facts.json`. */
export function parseFactsJson(text: string): ParseResult<Fact[]> {
  const json = parseJsonText(text);
  return json.ok ? parseFacts(json.value) : json;
}

/** True for a `{ "date": "YYYY-MM-DD" }` cell. */
export function isDateCell(value: unknown): value is DateCell {
  return typeof value === "object" && value !== null && typeof (value as DateCell).date === "string";
}

/** The gradable tasks of a lesson file: the lesson itself, or each task of a mock test. */
export function tasksOf(file: LessonFile): GradableTask[] {
  return file.kind === "lesson" ? [file] : file.tasks;
}
