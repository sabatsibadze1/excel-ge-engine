/**
 * `{{fact:income-tax-rate}}` placeholders in lesson texts (SCHEMA.md). Facts live in
 * `content/facts.json`, each with an official source.
 */
import type { Fact, LessonFile } from "./schema";

/** Any `{{…}}` placeholder, well-formed or not. */
const ANY_PLACEHOLDER_RE = /\{\{([^{}]*)\}\}/g;
const FACT_PLACEHOLDER_RE = /^\s*fact:([a-z0-9]+(?:-[a-z0-9]+)*)\s*$/;

export interface PlaceholderScan {
  /** Fact ids referenced as `{{fact:id}}`, unique, in order of appearance. */
  readonly factIds: string[];
  /** `{{…}}` texts that are not valid fact placeholders, e.g. `{{fact:VAT}}` or `{{vat}}`. */
  readonly malformed: string[];
}

export function scanPlaceholders(text: string): PlaceholderScan {
  const factIds: string[] = [];
  const malformed: string[] = [];
  for (const match of text.matchAll(ANY_PLACEHOLDER_RE)) {
    const id = FACT_PLACEHOLDER_RE.exec(match[1]!)?.[1];
    if (id === undefined) malformed.push(match[0]);
    else if (!factIds.includes(id)) factIds.push(id);
  }
  return { factIds, malformed };
}

/** Fact ids referenced as `{{fact:id}}` in a text. */
export function findFactPlaceholders(text: string): string[] {
  return scanPlaceholders(text).factIds;
}

/**
 * Replaces every `{{fact:id}}` with `render(id)`. The caller decides how a fact looks
 * (its `text_ka`, its value, a link to the source…). Malformed placeholders are left as they are.
 */
export function replaceFactPlaceholders(text: string, render: (factId: string) => string): string {
  return text.replace(ANY_PLACEHOLDER_RE, (whole: string, inner: string) => {
    const id = FACT_PLACEHOLDER_RE.exec(inner)?.[1];
    return id === undefined ? whole : render(id);
  });
}

/** How a fact appears in lesson texts: its `display` text, else its value as written (SCHEMA.md). */
export function renderFact(fact: Fact): string {
  if (fact.display !== undefined) return fact.display;
  return fact.value === null ? fact.text_ka : String(fact.value);
}

/** Replaces `{{fact:id}}` with each fact's display text; unknown ids are left as they are. */
export function renderFactPlaceholders(text: string, facts: readonly Fact[]): string {
  const byId = new Map(facts.map((fact) => [fact.id, fact]));
  return replaceFactPlaceholders(text, (id) => {
    const fact = byId.get(id);
    return fact ? renderFact(fact) : `{{fact:${id}}}`;
  });
}

/** Every learner-facing text of a lesson file that may contain fact placeholders. */
export function lessonTexts(file: LessonFile): { where: string; text: string }[] {
  if (file.kind === "lesson") {
    return [
      { where: "title", text: file.title },
      { where: "theory", text: file.theory },
      { where: "task", text: file.task },
      ...file.hints.map((text, i) => ({ where: `hints[${i}]`, text })),
    ];
  }
  return [
    { where: "title", text: file.title },
    { where: "intro", text: file.intro },
    ...file.tasks.map((task, i) => ({ where: `tasks[${i}].task`, text: task.task })),
  ];
}
