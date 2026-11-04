/**
 * Public API of the spreadsheet & grading engine. Framework-free: no React, no DOM, no Node APIs.
 * See README.md in this folder for usage and the known differences vs Excel.
 *
 * This is the full engine, with HyperFormula: use it in Node (tests, the lesson checker) and inside
 * the engine worker. Web pages use `excel-ge-engine/pure` (no HyperFormula) and talk to
 * `engine-worker.js` through `EngineClient`.
 */

// Types shared by everything.
export {
  FORMAT_MODES,
  FORMAT_SEPARATORS,
  isCellError,
  type CellError,
  type CellValue,
  type ExcelErrorType,
  type FormatMode,
  type FormatSeparators,
  type NumberFormatKind,
} from "./types";

// Lesson & facts format (zod schemas, inferred types, parsing with friendly errors).
export {
  CaseSchema,
  CheckSchema,
  DateCellSchema,
  FactSchema,
  FactsFileSchema,
  LessonFileSchema,
  LessonSchema,
  MockTestSchema,
  MustUseAlternativeSchema,
  SheetSchema,
  TaskSchema,
  formatIssues,
  isDateCell,
  parseFacts,
  parseFactsJson,
  parseLessonFile,
  parseLessonJson,
  tasksOf,
  type Case,
  type Check,
  type DateCell,
  type Fact,
  type GradableTask,
  type Lesson,
  type LessonFile,
  type MockTest,
  type MustUseAlternative,
  type ParseResult,
  type Sheet,
  type SheetCell,
  type Task,
} from "./schema";

// The live grid.
export { SheetSession, type CellView, type EditResult, type SheetSessionOptions } from "./workbook";

// Grading.
export {
  DEFAULT_SEED,
  DEFAULT_TOLERANCE,
  DEFAULT_TRIALS,
  describeMustUseKa,
  gradeTask,
  satisfiesMustUse,
  solutionFormulas,
  swapTrialInputs,
  valuesMatch,
  type CellGrade,
  type GradeFail,
  type GradeFailReason,
  type GradeOptions,
  type GradePass,
  type GradeResult,
} from "./grading";

// Excel format (en ↔ ka).
export {
  FORMAT_SWITCH_MESSAGES_KA,
  convertFormula,
  convertFormulasInText,
  convertInput,
  detectLikelyKaFormat,
  enToKa,
  kaToEn,
  looksLikeNumber,
  otherMode,
  suggestFormatSwitch,
} from "./format";
export { isValidFormula, isoDateToSerial, serialToIsoDate, supportedFunctionNames } from "./hyperformula";
export { excelRound } from "./excel-functions";

// Display.
export { alignmentOf, formatDate, formatNumber, formatTime, formatValue, valueToInput } from "./display";

// Errors.
export { ERROR_CODES, explainError, unsupportedFunctionTip, type ErrorExplanation } from "./errors";

// Formula editing helpers.
export { cycleAbsolute, type CaretEdit } from "./references";
export { functionNames, tokenizeFormula, type Token, type TokenType } from "./tokenizer";

// Addresses.
export {
  columnToLetters,
  expandRange,
  expandRanges,
  formatCellAddress,
  formatRange,
  isCellAddress,
  isRangeAddress,
  lettersToColumn,
  parseCellAddress,
  parseRange,
  type CellCoords,
  type RangeCoords,
} from "./addresses";

// Facts placeholders in lesson texts.
export {
  findFactPlaceholders,
  lessonTexts,
  renderFact,
  renderFactPlaceholders,
  replaceFactPlaceholders,
  scanPlaceholders,
} from "./facts";
