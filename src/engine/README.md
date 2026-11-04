# Spreadsheet & grading engine

Framework-free TypeScript (no React, no DOM, no Node APIs) around **HyperFormula 3.5.0**
(`licenseKey: 'gpl-v3'`, pinned exactly). It runs inside the engine worker in the browser, in Node
(tests, `check-lessons`) and later in Deno (B2B grading; Deno needs an import map for `hyperformula` and `zod`).

The lesson format is the contract in the content repo's `content/SCHEMA.md` (private); `schema.ts` implements it.

## Quick start

Web pages never import this folder directly: they use `excel-ge-engine/pure` and talk to the
standalone `engine-worker.js` (see the repo README). In Node (tests, the lesson checker) or inside
the worker, use the engine directly:

```ts
import { SheetSession, gradeTask } from "excel-ge-engine";

const session = new SheetSession(lesson.sheet, { mode: "ka" }); // or "en"
session.setCellInput("C2", "=ROUND(B2*0,784;2)");   // what the learner typed, in their format
session.getDisplayValue("C2");                       // "1960"  (decimal comma in ka: "1960,5")
session.getRawInput("C2");                           // "=ROUND(B2*0,784;2)"  (formula bar)

const result = gradeTask(lesson, session.getLearnerInputs(), session.mode);
if (!result.pass) show(result.messageKa);            // short Georgian message
session.destroy();
```

## `SheetSession` (workbook.ts): the live grid of one task

| Method | What it does |
|---|---|
| `new SheetSession(sheet, { mode })` | Loads the lesson data. Strings stay text, `{date}` cells become dates, pre-filled formulas are converted for Georgian mode, `TODAY()` returns `sheet.today`. |
| `rowCount`, `columnCount` | Grid size (covers the data and every editable cell). |
| `getCell(addr)` | Everything to draw a cell: `value`, `display`, `input`, `isFormula`, `isEditable`, `isSpilled`, `numberFormat`, `error` (`{ type, code, messageKa }`), `align` (Excel defaults: numbers right, text left, booleans/errors centre). |
| `getDisplayValue(addr)` | Shown value: `1960,5` (ka), `18%`, `01.02.2026`, `#N/A`, `TRUE`. |
| `getRawInput(addr)` | Formula-bar text in the learner's format; values as they would be typed (`'01001012345` for numeric-looking text). |
| `getValue(addr)` | `number \| string \| boolean \| null \| CellError`. |
| `isEditable(addr)`, `editableCells()` | Locked vs editable (`sheet.editable`). |
| `setCellInput(addr, raw)` | Enter what was typed. `""` clears. Returns `EditResult`: `{ ok: true, changed }` or `{ ok: false, reason: "locked", cells, messageKa }`. |
| `clear(range)` | Delete key. |
| `copy(range)` / `paste(target)` | Ctrl+C / Ctrl+V with Excel-style relative-reference adjustment (done by HyperFormula's `getFillRangeData`). A range target that is a multiple of the copied block is tiled. `copy` returns tab-separated text for the system clipboard. |
| `pasteText(target, text)` | Paste from outside the grid (tab/newline separated), as if typed. |
| `fillDown(range)` / `fillRight(range)` | Ctrl+D / Ctrl+R. On a single cell, copies from above / from the left. |
| `undo()`, `redo()`, `canUndo()`, `canRedo()` | Every edit above is one undo step (limit 100). |
| `setMode(mode)` | Excel-format switch: rebuilds the engine (never `updateConfig()`), converts learner formulas without changing their meaning. A formula typed in the wrong format is kept as typed, so switching fixes it. Clears undo history and the copy clipboard. |
| `getLearnerInputs()` | `{ C2: "=B2*0,784", … }` for `gradeTask`. |
| `subscribe(fn)`, `version` | Change notification. |

Every edit that touches a locked cell is refused as a whole, like a protected sheet in Excel.

## Grading (grading.ts), SCHEMA.md rules

```ts
gradeTask(task, learnerInputs, mode, { seed?, trials? }): GradeResult
```

- `task` is a lesson or one mock-test task (`{ sheet, check }`). Never uses the live grid: it builds a
  throwaway engine with sheets `sol` (solution filled across `check.cells` like Ctrl+D) and `lrn`.
- Order per graded cell: `empty` → `not_formula` → `missing_function` (`mustUse` alternatives:
  `"VLOOKUP"` or `["INDEX","MATCH"]` = all of them; whole names before `(`, outside quotes) →
  `excel_error` → `wrong_value`. Then 3 seeded swap trials and every `check.cases` dataset; failing
  only there is `hardcoded` (`trial` or `caseIndex` tells which).
- Numbers match with a RELATIVE tolerance (default `1e-9`), or after rounding to `compareDecimals`
  places when set. Text must be equal after trimming; errors match only the same error type.
- Swap values are never the original and never 0 (decimals kept; dates move ±1–60 days).
- The default seed is fixed (`DEFAULT_SEED`), so learners get exactly the trials the checker verified.

Result (`GradePass | GradeFail`): `pass`, `messageKa`, `cells` (one verdict per graded cell, for
colouring), and when failing: `reason`, `cell` (first failing, reading order), `expected`/`actual`
(display strings in the learner's format), `errorType`, `mustUse`, `trial`/`caseIndex`, `trialInputs`.

## Excel format (format.ts)

- `enToKa(f)` / `kaToEn(f)` / `convertFormula(f, from, to)`: the reviewed quote-aware converter:
  `=ROUND(B2*0.784,2)` ↔ `=ROUND(B2*0,784;2)`, arrays `{1,2;3,4}` ↔ `{1;2|3;4}`. Text in quotes is
  never touched.
- `suggestFormatSwitch(raw, mode)` → `"ka" | "en" | null`, `detectLikelyKaFormat(raw, mode)`, and the
  hint text `FORMAT_SWITCH_MESSAGES_KA.ka` ("როგორც ჩანს, წერტილმძიმეს (;) იყენებთ. გადავრთოთ ქართულ ფორმატზე?").
  After switching because of the hint, a number typed as `2500,5` in English mode is still text:
  call `setCellInput` again with the same text.
- `convertFormulasInText(text, mode)`: theory and hints write formulas in backticks in English
  format; this converts them for Georgian mode.

## Other helpers

- `cycleAbsolute(formula, caret, mode)` → `{ formula, caret }`: F4 on the reference at or next to the
  caret (`A1 → $A$1 → A$1 → $A1 → A1`; both ends of a range together).
- `explainError(type, detail?, mode?)` → `{ code, messageKa }` for `#N/A`, `#VALUE!`, `#NAME?`,
  `#DIV/0!`, `#REF!`, `#NUM!`, `#CYCLE!`, `#ERROR!` (parse error), `#SPILL!`. Uses HyperFormula's
  message to be precise (misspelled function name, unquoted text, wrong number of arguments).
- `renderFactPlaceholders(text, facts)`: `{{fact:vat-rate}}` → the fact's `display` (else its value).
- Addresses: `parseCellAddress("C2")` → `{ row: 1, col: 2 }`, `formatCellAddress(1, 2)`, `expandRange("C2:C6")`.
- Schema: `parseLessonFile(json)` / `parseLessonJson(text)` / `parseFacts(...)` → `{ ok, value }` or `{ ok: false, errors }`.

## Files

| File | Role |
|---|---|
| `hyperformula.ts` | `createEngine()`: the only place HyperFormula instances are built (configs, TRUE/FALSE names, `sheet.today`). |
| `excel-functions.ts` | Custom functions, registered once: RANK, RANK.EQ, CONCAT, Excel-style ROUND, frozen TODAY, any-case DATEDIF. |
| `convert.ts`, `format.ts`, `tokenizer.ts`, `references.ts` | Format conversion (`convert.ts`, no HyperFormula) and format-switch suggestions (`format.ts`), a lossless formula tokenizer (function names, F4), F4. |
| `workbook.ts`, `sheet-data.ts`, `display.ts` | Live grid, lesson data loading, value formatting. |
| `grading.ts`, `random.ts` | Swap-the-inputs grading and its seeded randomness. |
| `schema.ts`, `facts.ts`, `errors.ts` | Lesson format, fact placeholders, Georgian error texts. |

Import HyperFormula only through ESM `import` (the package is `"type": "module"`): a mix of
`require` and `import` loads two copies, and the custom functions apply to one only.

## Known differences vs Excel (found by testing HyperFormula 3.5.0)

Fixed or worked around in this engine:

- **`TRUE`/`FALSE`:** HyperFormula only knows `TRUE()`/`FALSE()`; `=VLOOKUP(…,FALSE)` gave `#NAME?`. Defined as names.
- **Missing functions:** RANK, RANK.EQ and CONCAT are added as custom functions. Still missing:
  RANK.AVG, AVERAGEIFS, XMATCH, SORTBY, INDIRECT, LET, LAMBDA, NUMBERVALUE, FIXED, DOLLAR,
  TEXTSPLIT/TEXTBEFORE/TEXTAFTER (learners get a Georgian "not available yet" tip).
- **ROUND:** HyperFormula gives `ROUND(1.25*0.18,2) = 0.22` and `ROUND(1.005,2) = 1`; ours rounds the
  15-digit decimal like Excel (0.23, 1.01).
- **DATEDIF:** HyperFormula only accepts upper-case units (`"D"`); ours accepts `"d"`, `"m"`, `"y"`, `"md"`…
- **Separators:** `updateConfig()` with other separators turns formulas into `#ERROR!`; the format switch rebuilds instead.

Still different (lessons must avoid these; the checker enforces most of them):

- **TEXT formats:** only zero padding (`"00000000000"`), `"0"`, `"0.00"` and dates (`"dd.mm.yyyy"`) work.
  `"#,##0.00"` gives `"1235,##0.00"`, `"0%"` gives `"0%"`. Format strings always use `.` for decimals,
  even in Georgian mode.
- **Numbers turned into text** (`=A1&""`, `="x"&1.5`) always use `.`, even in Georgian mode.
- **Criteria strings** like `">1.5"` in COUNTIF/SUMIF are text, so they are not converted: in Georgian
  mode they must be written `">1,5"`.
- **Scientific notation:** `=1E+3` and `=1E3` do not parse (`=1e3` does). `""` inside text (`"say ""hi"""`) does not parse.
- **Approximate lookups:** VLOOKUP/MATCH default to approximate match, like Excel, but give different
  answers on unsorted keys. Always use `FALSE`/`0`.
- **PROPER/UPPER** turn Georgian letters into Mtavruli (`"Გიორგი"`).
- **Large numbers:** results are rounded to about 10 significant digits (`=123456789012` → 123456789010).
  Keep IDs as text.
- **Errors Excel would refuse to accept** show in the cell instead: parse errors `#ERROR!`, circular
  references `#CYCLE!` (Excel shows 0 and a warning), a wrong number of arguments `#N/A`.
- **Dynamic arrays:** `=SUM(B2:B6*2)` works and `=B2:B6*2` spills like Excel 365, but `=A1:A3` alone gives `#VALUE!`.
- **Typed input:** no thousands separators (`1,234` is text in English mode); dates are typed as
  `DD/MM/YYYY` or `DD.MM.YYYY` and shown as `dd.mm.yyyy`; `$` is the only currency symbol recognised;
  spaces before `=` are ignored (Excel would store `" =B2"` as text). R1C1 references are not supported.

Same as Excel (checked): `=A1` on an empty cell gives 0; `="10"*2` gives 20; SUM ignores text in
ranges; `"A"="a"` is TRUE (EXACT is case-sensitive); `-2^2` is 4; dates are Excel serial numbers.
