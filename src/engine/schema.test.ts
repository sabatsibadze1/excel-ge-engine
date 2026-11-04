import { describe, expect, it } from "vitest";
import { parseFacts, parseLessonFile, parseLessonJson, tasksOf } from "./schema";

const lesson = {
  id: "payroll-gross-net",
  kind: "lesson",
  module: 1,
  order: 5,
  free: true,
  title: "ხელზე ასაღები ხელფასი",
  skills: ["arithmetic"],
  theory: "Net = Gross × 0.784. {{fact:income-tax-rate}}",
  task: "C2 უჯრაში გამოთვალეთ ხელზე ასაღები ხელფასი.",
  hints: ["`=B2*0.98*0.8`"],
  sheet: {
    cells: [
      ["თანამშრომელი", "დარიცხული (₾)", "ხელზე ასაღები (₾)", "თარიღი"],
      ["გიორგი ბერიძე", 2500, null, { date: "2026-02-01" }],
    ],
    editable: ["C2"],
    today: null,
  },
  check: {
    cells: "C2",
    solution: "=B2*0.98*0.8",
    vary: ["B2"],
    mustUse: [["INDEX", "MATCH"], "XLOOKUP"],
    tolerance: 1e-9,
    cases: [{ B2: 2000 }, { D2: { date: "2026-03-01" } }],
  },
  facts: ["income-tax-rate"],
  source: "ai-draft",
  verified: { checker: false, claudeAudit: null, founderPlayed: "2026-11-11" },
};

const mock = {
  id: "mini-mock-foundations",
  kind: "mock",
  module: 1,
  order: 8,
  free: true,
  title: "მინი ტესტი",
  intro: "5 დავალება, 10 წუთი.",
  minutes: 10,
  tasks: [{ task: "B8-ში ჯამი.", sheet: { cells: [[1], [2]], editable: ["B8"] }, check: { cells: "B8", solution: "=SUM(A1:A2)", vary: ["A1:A2"] }, facts: [] }],
  source: "ai-draft",
  verified: { checker: false, claudeAudit: null, founderPlayed: null },
};

describe("parseLessonFile", () => {
  it("accepts a lesson with dates, today, cases and mustUse alternatives", () => {
    const result = parseLessonFile(lesson);
    expect(result.ok).toBe(true);
    if (result.ok) expect(tasksOf(result.value)).toHaveLength(1);
  });

  it("accepts a mock test", () => {
    const result = parseLessonFile(mock);
    expect(result.ok).toBe(true);
    if (result.ok) expect(tasksOf(result.value)).toHaveLength(1);
  });

  it("explains problems with paths and friendly messages", () => {
    const bad = {
      ...lesson,
      id: "Payroll Gross",
      check: { ...lesson.check, cells: "c2-c6", mustuse: [] },
      sheet: { ...lesson.sheet, cells: [[{ date: "2026-02-30" }]] },
    };
    const result = parseLessonFile(bad);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.errors).toEqual(
      expect.arrayContaining([
        expect.stringMatching(/^id: expected a kebab-case id/),
        'check.cells: expected a cell like "C2" or a range like "C2:C6", got "c2-c6"',
        'check: unknown field "mustuse" (did you mean "mustUse"?)',
        expect.stringMatching(/^sheet\.cells\[0\]\[0\]/),
      ]),
    );
  });

  it("reports missing fields and a bad kind", () => {
    const { task: _task, ...noTask } = lesson;
    const missing = parseLessonFile(noTask);
    expect(missing.ok || missing.errors).toContain("task: is required (expected string)");
    const kind = parseLessonFile({ ...lesson, kind: "quiz" });
    expect(kind.ok || kind.errors).toEqual(['kind: expected "lesson" or "mock", got "quiz"']);
  });

  it("reports invalid JSON", () => {
    const result = parseLessonJson("{ not json");
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.errors[0]).toMatch(/^invalid JSON/);
  });
});

describe("parseFacts", () => {
  it("accepts the facts registry format", () => {
    const facts = [
      { id: "vat-rate", text_ka: "დღგ — 18%", value: 0.18, source: "https://matsne.gov.ge", checked: null, checkedBy: null },
      { id: "iban-country-code", text_ka: "GE", value: "GE", source: "https://nbg.gov.ge", checked: "2026-11-02", checkedBy: "founder", notes: "x" },
    ];
    expect(parseFacts(facts).ok).toBe(true);
    const bad = parseFacts([{ ...facts[0], checkedBy: "someone" }]);
    expect(bad.ok).toBe(false);
  });
});
