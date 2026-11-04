# excel-ge-engine

The spreadsheet and grading engine of a Georgian-language Excel practice website. Learners type real
formulas; the engine calculates them with [HyperFormula](https://hyperformula.handsontable.com/) and
checks answers by **swapping the inputs** (any correct formula passes, typed-in numbers fail).

Licensed under the **GNU General Public License v3.0 only** ([LICENSE](LICENSE)), because it includes
HyperFormula under its GPLv3 option (`licenseKey: 'gpl-v3'`).

## What's here

| Path | Contents |
|---|---|
| `src/engine/` | The engine: HyperFormula setup, the live sheet (`SheetSession`), grading, the English ↔ Georgian formula format (`,` vs `;`), Georgian error explanations, the lesson format (zod). See its [README](src/engine/README.md). |
| `src/checker/` | The lesson checker: validates a lesson file and proves its solution, swap test and cases. |
| `src/worker/` | The **engine worker**: `worker.ts` (entry), `protocol.ts` (the JSON messages), `host.ts` (worker side), `client.ts` (page side, no HyperFormula). |
| `src/pure.ts` | What a web page may import without loading HyperFormula: the client, lesson format, addresses, converter, display helpers. A test enforces it. |
| `scripts/build-worker.mjs` | Builds `engine-worker.js`, one standalone file for `new Worker(url)`. |
| `scripts/check-lessons.ts` | Lesson checker CLI. |

## How a website uses it

The website runs the engine as a separate program in the browser and only exchanges JSON messages
with it:

```ts
import { EngineClient } from "excel-ge-engine/pure";

const engine = await EngineClient.start("/engine-worker.js");
const sheet = await engine.openSheet(lesson.sheet, { mode: "en" });
await sheet.setCellInput("C2", "=B2*0.784");
sheet.getCell("C2").display;                                   // "1960"
const result = await engine.grade(lesson, sheet.getLearnerInputs(), sheet.mode);
```

Build the worker into the site's static folder: `node node_modules/excel-ge-engine/scripts/build-worker.mjs public/engine-worker.js`.
The file starts with a notice naming this license and this repository.

## Commands

Node 22 (`.nvmrc`), npm:

| Command | What it does |
|---|---|
| `npm test` | Vitest: engine, grading, checker, worker protocol, purity and worker-build tests. |
| `npm run typecheck` | TypeScript. |
| `npm run build` | Writes `dist/engine-worker.js`. |
| `npm run check-lessons -- <folders> --facts <file> [--strict]` | Checks lesson files (run it from a folder with `content/lessons`, or pass paths). |

## Rules for changes

- HyperFormula is pinned to exactly 3.5.0. Build instances only through `createEngine()`; never call
  `updateConfig()` on a live engine.
- Grading is swap-the-inputs, never regex. Format conversion uses `enToKa`/`kaToEn`, never a naive
  `;` → `,` replace.
- Nothing reachable from `src/pure.ts` may import HyperFormula.
- Bump `PROTOCOL_VERSION` in `src/worker/protocol.ts` when a message changes shape.
- No lesson content in this repo except the small test fixtures.
