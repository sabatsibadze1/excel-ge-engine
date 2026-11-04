import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const src = path.dirname(fileURLToPath(import.meta.url));

/** Value imports and re-exports of a file (type-only ones are erased by the compiler). */
function valueImports(file: string): string[] {
  const text = readFileSync(file, "utf8");
  const specs: string[] = [];
  const re = /^\s*(import|export)\s+(type\s+)?([^;]*?)\s+from\s+"([^"]+)";?/gms;
  for (const m of text.matchAll(re)) {
    if (m[2]) continue; // import type / export type
    const clause = m[3]!.trim();
    if (/^\{[^}]*\}$/s.test(clause) && clause.slice(1, -1).split(",").every((part) => /^\s*(type\s|$)/.test(part))) continue;
    specs.push(m[4]!);
  }
  for (const m of text.matchAll(/^\s*import\s+"([^"]+)";?/gm)) specs.push(m[1]!);
  return specs;
}

function graph(entry: string): { files: Set<string>; packages: Set<string> } {
  const files = new Set<string>();
  const packages = new Set<string>();
  const visit = (file: string) => {
    if (files.has(file)) return;
    files.add(file);
    for (const spec of valueImports(file)) {
      if (spec.startsWith(".")) visit(path.resolve(path.dirname(file), `${spec}.ts`));
      else packages.add(spec);
    }
  };
  visit(entry);
  return { files, packages };
}

describe("excel-ge-engine/pure", () => {
  it("never loads HyperFormula", () => {
    const { files, packages } = graph(path.join(src, "pure.ts"));
    const relative = [...files].map((f) => path.relative(src, f)).sort();
    expect([...packages]).toEqual(["zod"]);
    for (const banned of ["engine/hyperformula.ts", "engine/workbook.ts", "engine/grading.ts", "engine/format.ts", "engine/excel-functions.ts", "engine/sheet-data.ts", "worker/host.ts"]) {
      expect(relative).not.toContain(banned);
    }
    expect(relative).toContain("worker/client.ts");
  });

  it("the full engine does load HyperFormula (sanity check of the scan)", () => {
    expect([...graph(path.join(src, "engine/index.ts")).packages]).toContain("hyperformula");
  });
});
