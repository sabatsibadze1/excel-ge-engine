/**
 * Lesson checker CLI.
 *
 *   npm run check-lessons -- [folders or files...] [--facts <path>] [--strict]
 *
 * Run it from the app folder. Defaults: `content/lessons` there, plus `../content/lessons` when it
 * exists (the private content repo). Facts: `../content/facts.json` when it exists, else `content/facts.json`.
 * --strict: unverified or stale facts used by the checked lessons are errors (needed to publish).
 * Exit code: 0 = no errors, 1 = errors found, 2 = bad arguments or missing files.
 * The checks themselves live in src/checker/check-lessons.ts.
 */
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { parseArgs } from "node:util";
import { checkLessons, type CheckIssue, type CheckReport, type SourceFile } from "../src/checker/check-lessons";

/** Run from the app folder: its `content/` ends up in the website build, so paid lessons must never be there. */
const APP_ROOT = process.cwd();
const PUBLIC_CONTENT = path.join(APP_ROOT, "content");

const USAGE = `Usage: npm run check-lessons -- [folders or files...] [--facts <path>] [--strict]

  folders/files  lesson folders or single .json files
                 (default: content/lessons, plus ../content/lessons if it exists)
  --facts <path> facts registry (default: ../content/facts.json if it exists, else content/facts.json)
  --strict       unverified (checked: null) or stale (> 12 months) facts used by the
                 checked lessons are errors, not warnings (required before publishing)
  --help         show this help`;

const useColor = process.stdout.isTTY && !process.env.NO_COLOR;
const paint = (code: number) => (text: string) => (useColor ? `\u001b[${code}m${text}\u001b[0m` : text);
const red = paint(31);
const yellow = paint(33);
const green = paint(32);
const dim = paint(2);
const bold = paint(1);

function display(p: string): string {
  const relative = path.relative(process.cwd(), p);
  return relative === "" ? "." : relative;
}

function fail(message: string): never {
  console.error(`check-lessons: ${message}\n\n${USAGE}`);
  process.exit(2);
}

function isInside(child: string, parent: string): boolean {
  const relative = path.relative(parent, child);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function source(file: string, folder: string): SourceFile {
  return {
    path: display(file),
    folder: display(folder),
    text: readFileSync(file, "utf8"),
    isPublic: isInside(file, PUBLIC_CONTENT),
  };
}

function collect(target: string): SourceFile[] {
  const absolute = path.resolve(target);
  if (!existsSync(absolute)) fail(`not found: ${target}`);
  if (statSync(absolute).isFile()) return [source(absolute, path.dirname(absolute))];
  return readdirSync(absolute)
    .filter((name) => name.endsWith(".json"))
    .sort()
    .map((name) => source(path.join(absolute, name), absolute));
}

function printIssue(issue: CheckIssue): void {
  const label = issue.level === "error" ? red("error  ") : yellow("warning");
  console.log(`      ${label} ${issue.message}`);
}

function print(report: CheckReport, folders: string[], factsPath: string | null, strict: boolean): void {
  console.log(
    bold(`Checking ${report.files.length} lesson file(s) in ${folders.join(", ")}`) +
      dim(` · facts: ${factsPath ?? "none"}${strict ? " · strict" : ""}`),
  );
  let currentFolder = "";
  for (const file of report.files) {
    if (file.folder !== currentFolder) {
      currentFolder = file.folder;
      console.log(`\n  ${bold(currentFolder)}`);
    }
    const errors = file.issues.filter((i) => i.level === "error").length;
    const status = errors > 0 ? red("FAIL") : file.issues.length > 0 ? yellow("WARN") : green("OK  ");
    const what = file.id ? dim(` ${file.kind} ${file.id}`) : "";
    console.log(`    ${status} ${path.basename(file.path)}${what}`);
    file.issues.forEach(printIssue);
  }
  if (report.facts.path) {
    const errors = report.facts.issues.filter((i) => i.level === "error").length;
    const status = errors > 0 ? red("FAIL") : report.facts.issues.length > 0 ? yellow("WARN") : green("OK  ");
    console.log(`\n  ${bold("Facts")}\n    ${status} ${report.facts.path}${dim(` ${report.facts.count} facts`)}`);
    report.facts.issues.forEach(printIssue);
  }
  const verdict = report.errors > 0 ? red("FAILED") : green("PASSED");
  console.log(
    `\n${bold("Summary:")} ${report.files.length} file(s), ${report.errors} error(s), ${report.warnings} warning(s). ${verdict}`,
  );
}

function main(): void {
  let parsed;
  try {
    parsed = parseArgs({
      allowPositionals: true,
      options: { facts: { type: "string" }, strict: { type: "boolean", default: false }, help: { type: "boolean" } },
    });
  } catch (e) {
    fail((e as Error).message);
  }
  if (parsed.values.help) {
    console.log(USAGE);
    return;
  }

  const targets =
    parsed.positionals.length > 0
      ? parsed.positionals
      : [path.join(APP_ROOT, "content", "lessons"), path.join(APP_ROOT, "..", "content", "lessons")].filter(
          (folder, i) => i === 0 || existsSync(folder),
        );

  const contentFacts = path.join(APP_ROOT, "..", "content", "facts.json");
  const factsPath = parsed.values.facts
    ? path.resolve(parsed.values.facts)
    : existsSync(contentFacts)
      ? contentFacts
      : path.join(APP_ROOT, "content", "facts.json");
  if (parsed.values.facts && !existsSync(factsPath)) fail(`facts file not found: ${parsed.values.facts}`);

  const files = targets.flatMap(collect);
  const facts = existsSync(factsPath) ? { path: display(factsPath), text: readFileSync(factsPath, "utf8") } : null;
  const report = checkLessons(files, facts, { strict: parsed.values.strict });
  print(report, targets.map(display), facts?.path ?? null, parsed.values.strict ?? false);
  process.exitCode = report.errors > 0 ? 1 : 0;
}

main();
