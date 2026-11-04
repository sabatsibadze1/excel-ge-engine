import { execFileSync } from "node:child_process";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { checkLessons, type CheckReport, type SourceFile } from "./check-lessons";

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, "../..");
const fixtures = path.join(here, "fixtures");
const factsFile = { path: "facts.json", text: readFileSync(path.join(fixtures, "facts.json"), "utf8") };
const TODAY = new Date("2026-11-02T00:00:00Z");

function folder(dir: string, extra: Partial<SourceFile> = {}): SourceFile[] {
  return readdirSync(dir)
    .filter((name) => name.endsWith(".json"))
    .map((name) => ({ path: path.join(dir, name), folder: dir, text: readFileSync(path.join(dir, name), "utf8"), ...extra }));
}

const baseLesson = {
  id: "test-lesson",
  kind: "lesson",
  module: 1,
  order: 1,
  free: true,
  title: "ტესტი",
  skills: ["test"],
  theory: "თეორია.",
  task: "C2 უჯრაში ჩაწერეთ ფორმულა.",
  hints: ["პირველი", "მეორე", "მესამე"],
  sheet: { cells: [["სახელი", "დარიცხული", "ხელზე"], ["გიორგი", 2500, null]], editable: ["C2"] },
  check: { cells: "C2", solution: "=B2*0.98*0.8", vary: ["B2"] },
  facts: [],
  source: "ai-draft",
  verified: { checker: false, claudeAudit: null, founderPlayed: null },
};

type LessonPatch = Record<string, unknown> & { check?: Record<string, unknown>; sheet?: Record<string, unknown> };

function lesson(patch: LessonPatch = {}) {
  return {
    ...baseLesson,
    ...patch,
    sheet: { ...baseLesson.sheet, ...patch.sheet },
    check: { ...baseLesson.check, ...patch.check },
  };
}

function check(lessons: object[], options: Parameters<typeof checkLessons>[2] = {}, facts = factsFile): CheckReport {
  const files = lessons.map((l, i) => ({ path: `0${i + 1}-test.json`, folder: "lessons", text: JSON.stringify(l) }));
  return checkLessons(files, facts, { today: TODAY, ...options });
}

function messages(report: CheckReport, level: "error" | "warning" = "error"): string[] {
  return report.files.flatMap((f) => f.issues).filter((i) => i.level === level).map((i) => i.message);
}

describe("checkLessons: fixtures", () => {
  it("passes a good lesson", () => {
    const report = checkLessons(folder(path.join(fixtures, "good")), factsFile, { today: TODAY });
    expect(messages(report)).toEqual([]);
    expect(report.errors).toBe(0);
  });

  it("(a) catches a lesson whose vary does not affect the result", () => {
    const report = checkLessons(folder(path.join(fixtures, "bad-vary")), factsFile, { today: TODAY });
    expect(report.errors).toBeGreaterThan(0);
    expect(messages(report).join("\n")).toContain("changing D2 on its own never changes the result");
    expect(messages(report).join("\n")).toContain("no trial or case changes the result in C2");
  });

  it("(b) catches a lesson whose solution gives an Excel error", () => {
    const report = checkLessons(folder(path.join(fixtures, "bad-solution")), factsFile, { today: TODAY });
    expect(messages(report).join("\n")).toContain("check.solution gives #DIV/0! in C2 on the lesson data");
  });
});

describe("checkLessons: structure", () => {
  it("requires graded cells to be editable and vary cells to hold locked numbers or dates", () => {
    const report = check([
      lesson({
        sheet: { cells: [["a", "b", "c", "d"], ["x", 2500, null, "text"]], editable: ["B2"] },
        check: { vary: ["B2", "D2", "C2"] },
      }),
    ]);
    const all = messages(report).join("\n");
    expect(all).toContain("C2 is graded but not editable");
    expect(all).toContain('D2 must hold a number or a date, but it holds text "text"');
    expect(all).toContain("B2 is editable; vary cells must be locked lesson data");
    expect(all).toContain("C2 is also a graded cell");
  });

  it("rejects cases that touch graded cells", () => {
    expect(messages(check([lesson({ check: { cases: [{ C2: 5 }] } })])).join()).toContain("cases may only change lesson data");
  });

  it("warns about an empty vary without cases, and the anti-hardcode test fails it", () => {
    const report = check([lesson({ check: { vary: [] } })]);
    expect(messages(report, "warning").join()).toContain("vary is empty and there are no cases");
    expect(messages(report).join()).toContain("anti-hardcode self-test");
  });

  it("checks ids and orders are unique within a folder, not across folders", () => {
    const text = JSON.stringify(lesson());
    const same = checkLessons(
      [
        { path: "01-a.json", folder: "x", text },
        { path: "01-b.json", folder: "x", text },
        { path: "01-c.json", folder: "y", text },
      ],
      factsFile,
      { today: TODAY },
    );
    expect(messages(same).filter((m) => m.includes("also used by"))).toHaveLength(4);
    expect(same.files[2]!.issues).toEqual([]);
  });

  it("rejects a paid lesson in the public app repo", () => {
    const report = checkLessons(
      [{ path: "01-paid.json", folder: "content/lessons", text: JSON.stringify(lesson({ free: false })), isPublic: true }],
      factsFile,
      { today: TODAY },
    );
    expect(messages(report).join()).toContain("paid lesson (free: false) in the public app repo");
  });

  it("reports invalid JSON and schema errors without crashing", () => {
    const report = checkLessons(
      [
        { path: "01-x.json", folder: "f", text: "{ nope" },
        { path: "02-y.json", folder: "f", text: JSON.stringify({ ...baseLesson, order: 2, kind: "quiz" }) },
      ],
      factsFile,
    );
    expect(messages(report)).toEqual([expect.stringMatching(/^invalid JSON/), 'kind: expected "lesson" or "mock", got "quiz"']);
  });

  it("warns when the file name and order disagree", () => {
    expect(messages(check([lesson({ order: 7 })]), "warning").join()).toContain("file name starts with 01 but order is 7");
  });
});

describe("checkLessons: solution rules (SCHEMA.md 5–8)", () => {
  it("errors when the filled solution refers to empty cells (missing $)", () => {
    const report = check([
      lesson({
        sheet: { cells: [["a", "b", "c", "d", 0.18], ["x", 100, null], ["y", 200, null]], editable: ["C2:C3"] },
        check: { cells: "C2:C3", solution: "=B2*(1+E1)", vary: ["B2:B3", "E1"] },
      }),
    ]);
    expect(messages(report).join()).toContain("filled into C3 (=B3*(1+E2)) refers to E2, which is empty");
  });

  it("rejects TEX formats other than zero padding, \"0\" and dd.mm.yyyy", () => {
    const bad = check([lesson({ check: { solution: '=TEXT(B2*0.784,"#,##0.00")' } })]);
    expect(messages(bad).join()).toContain('TEXT format "#,##0.00" is not allowed');
    const good = check([
      lesson({
        sheet: { cells: [["a", "b", "c"], ["x", 1001012345, null]], editable: ["C2"] },
        check: { solution: '=TEXT(B2,"00000000000")', vary: ["B2"] },
      }),
    ]);
    expect(messages(good)).toEqual([]);
  });

  it('rejects "" inside text and 1E+3-style numbers', () => {
    expect(messages(check([lesson({ check: { solution: '=B2&" ""x"""' } })])).join()).toContain('"" inside a text string');
    expect(messages(check([lesson({ check: { solution: "=B2*1e-3" } })])).join()).toContain("no 1E+3-style numbers");
  });

  it("requires sheet.today for TODAY() and rejects NOW()", () => {
    const dated = { cells: [["a", "b", "c"], ["x", { date: "2026-10-01" }, null]], editable: ["C2"] };
    expect(messages(check([lesson({ sheet: dated, check: { solution: "=TODAY()-B2" } })])).join()).toContain("add sheet.today");
    expect(messages(check([lesson({ sheet: { ...dated, today: "2026-11-02" }, check: { solution: "=TODAY()-B2" } })]))).toEqual([]);
    expect(messages(check([lesson({ sheet: dated, check: { solution: "=NOW()-B2" } })])).join()).toContain("NOW() is not allowed");
  });

  it("rejects unsupported functions and unsatisfiable mustUse", () => {
    expect(messages(check([lesson({ check: { solution: "=AVERAGEIFS(B2,B2,1)" } })])).join()).toContain("AVERAGEIFS is not supported");
    const report = check([lesson({ check: { mustUse: ["ROUND", ["VLOKUP"]] } })]);
    expect(messages(report).join()).toContain('"VLOKUP" is not a function our engine knows');
    expect(messages(report).join()).toContain("the solution satisfies none of the alternatives");
  });

  it("warns about IF without boundary cases, cases that test nothing, and unvaried inputs", () => {
    const sheet = { cells: [["a", "b", "c", "goal"], ["x", 2500, null, 2000]], editable: ["C2"] };
    const noCases = check([lesson({ sheet, check: { solution: "=IF(B2>=D2,B2*0.05,0)", vary: ["B2"] } })]);
    const warnings = messages(noCases, "warning").join("\n");
    expect(warnings).toContain("IF/IFS tasks need cases");
    expect(warnings).toContain("the solution reads D2, but neither vary nor any case changes them");
    const idle = check([lesson({ sheet, check: { solution: "=IF(B2>=D2,B2*0.05,0)", vary: ["B2"], cases: [{ D2: 2001 }] } })]);
    expect(messages(idle, "warning").join()).toContain("gives the same result as the lesson data");
  });

  it("warns about PROPER/UPPER on Georgian text", () => {
    const report = check([
      lesson({
        sheet: { cells: [["a", "b", "c"], ["გიორგი", 2500, null]], editable: ["C2"] },
        check: { solution: '=PROPER(A2)&" "&B2' },
      }),
    ]);
    expect(messages(report, "warning").join()).toContain("Mtavruli");
  });

  it("asserts that every mustFail formula fails", () => {
    const ok = check([lesson({ check: { mustFail: ["=B2*0.78", "=2500*0.784"] } })]);
    expect(messages(ok)).toEqual([]);
    const report = check([lesson({ check: { mustFail: ["=B2*0.784", "=ROUND(B2*0,784;2)"] } })]);
    const all = messages(report).join("\n");
    expect(all).toContain("check.mustFail[0]: the wrong answer =B2*0.784 passes grading");
    expect(all).toContain("check.mustFail[1]: =ROUND(B2*0,784;2) is not a valid English-format formula");
  });

  it("warns about compareDecimals on a ROUND task and about hint counts", () => {
    const sheet = { cells: [["a", "b", "c"], ["x", 1180, null]], editable: ["C2"] };
    const report = check([
      lesson({ sheet, hints: ["ერთი"], check: { solution: "=ROUND(B2/1.18*0.18,2)", compareDecimals: 2 } }),
    ]);
    const warnings = messages(report, "warning").join("\n");
    expect(warnings).toContain("compareDecimals is set on a task whose solution uses ROUND");
    expect(warnings).toContain("hints: 1 hint(s); SCHEMA.md asks for exactly 3");
  });

  it("checks every mock task with a tasks[i] prefix", () => {
    const mock = {
      ...baseLesson,
      kind: "mock",
      intro: "ტესტი",
      minutes: 10,
      tasks: [
        { task: "C2", sheet: baseLesson.sheet, check: baseLesson.check, facts: [] },
        { task: "C2", sheet: baseLesson.sheet, check: { ...baseLesson.check, solution: "=B2/0" }, facts: [] },
      ],
    };
    for (const key of ["skills", "theory", "task", "hints", "sheet", "check", "facts"] as const) delete (mock as Record<string, unknown>)[key];
    expect(messages(check([mock]))).toEqual([expect.stringMatching(/^tasks\[1\]\.check\.solution gives #DIV\/0!/)]);
  });
});

describe("checkLessons: facts", () => {
  it("errors on unknown facts and placeholders, warns when a placeholder's fact is not listed", () => {
    const report = check([
      lesson({ facts: ["no-such-fact"], theory: "{{fact:income-tax-rate}} {{fact:missing}} {{fact:VAT}}" }),
    ]);
    const all = messages(report).join("\n");
    expect(all).toContain('unknown fact "no-such-fact"');
    expect(all).toContain("{{fact:missing}} refers to an unknown fact");
    expect(all).toContain("malformed placeholder {{fact:VAT}}");
    expect(messages(report, "warning").join()).toContain('"income-tax-rate" is not listed in facts');
  });

  it("warns about unverified and stale facts; --strict fails the ones lessons use", () => {
    const facts = {
      path: "facts.json",
      text: JSON.stringify([
        { id: "used-unverified", text_ka: "ა", value: 1, source: "https://x.ge", checked: null, checkedBy: null },
        { id: "used-stale", text_ka: "ბ", value: 2, source: "https://x.ge", checked: "2025-10-01", checkedBy: "founder" },
        { id: "unused", text_ka: "გ", value: 3, source: "https://x.ge", checked: null, checkedBy: null },
        { id: "fresh", text_ka: "დ", value: 4, source: "https://x.ge", checked: "2026-10-01", checkedBy: "expert" },
      ]),
    };
    const lessons = [lesson({ facts: ["used-unverified", "used-stale", "fresh"] })];
    const normal = check(lessons, {}, facts);
    expect(normal.errors).toBe(0);
    expect(normal.facts.issues.map((i) => i.level)).toEqual(["warning", "warning", "warning"]);
    const strict = check(lessons, { strict: true }, facts);
    expect(strict.facts.issues.map((i) => [i.level, i.message.split(":")[0]])).toEqual([
      ["error", "used-unverified"],
      ["error", "used-stale"],
      ["warning", "unused"],
    ]);
  });
});

describe("check-lessons CLI", () => {
  const run = (...args: string[]) => {
    try {
      const stdout = execFileSync(process.execPath, ["--import", "tsx", "scripts/check-lessons.ts", ...args], {
        cwd: repoRoot,
        encoding: "utf8",
        env: { ...process.env, NO_COLOR: "1" },
      });
      return { code: 0, stdout };
    } catch (e) {
      const err = e as { status: number; stdout: string; stderr: string };
      return { code: err.status, stdout: err.stdout + err.stderr };
    }
  };

  it("exits 0 for good lessons and 1 for bad ones, with a summary", () => {
    const facts = ["--facts", path.join(fixtures, "facts.json")];
    const good = run(path.join(fixtures, "good"), ...facts);
    expect(good.code).toBe(0);
    expect(good.stdout).toContain("PASSED");
    const bad = run(path.join(fixtures, "bad-vary"), path.join(fixtures, "bad-solution"), ...facts);
    expect(bad.code).toBe(1);
    expect(bad.stdout).toContain("FAIL 05-vary-ignored.json");
    expect(bad.stdout).toContain("FAILED");
  });

  it("exits 2 for a missing folder", () => {
    expect(run("no/such/folder").code).toBe(2);
  });
});
