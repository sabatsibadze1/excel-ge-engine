/**
 * Short, friendly Georgian explanations for Excel errors (learn mode, build plan §6.5).
 * HyperFormula's English message (`detail`) is used only to pick a more precise explanation.
 */
import { convertFormula } from "./convert";
import type { ExcelErrorType, FormatMode } from "./types";

/** What each error shows in a cell. */
export const ERROR_CODES: Readonly<Record<ExcelErrorType, string>> = {
  DIV_BY_ZERO: "#DIV/0!",
  NA: "#N/A",
  NAME: "#NAME?",
  NUM: "#NUM!",
  REF: "#REF!",
  VALUE: "#VALUE!",
  CYCLE: "#CYCLE!",
  ERROR: "#ERROR!",
  SPILL: "#SPILL!",
  LIC: "#LIC!",
};

export interface ErrorExplanation {
  readonly type: ExcelErrorType;
  /** "#N/A", "#DIV/0!"… */
  readonly code: string;
  /** One or two short Georgian sentences for the learner. */
  readonly messageKa: string;
}

/** Base explanations (wording from the content review). */
const MESSAGES_KA: Readonly<Record<ExcelErrorType, string>> = {
  NA: "მნიშვნელობა ვერ მოიძებნა. გადაამოწმეთ, საძიებო მნიშვნელობა ზუსტად ემთხვევა თუ არა ცხრილში ჩაწერილს. ხშირი მიზეზია ზედმეტი ჰარი (Space) ან ტექსტად ჩაწერილი რიცხვი.",
  VALUE: "გამოთვლაში რიცხვის ნაცვლად ტექსტი მოხვდა. შეამოწმეთ, რომელ უჯრებს იყენებს ფორმულა.",
  NAME: "ფუნქციის სახელი არასწორად წერია (ყველა ფუნქცია ინგლისურად იწერება) ან ტექსტი ბრჭყალებში არ არის ჩასმული.",
  DIV_BY_ZERO: "ნულზე გაყოფა. შეამოწმეთ გამყოფი უჯრა: ცარიელი ან 0 ხომ არ არის?",
  REF: "ფორმულა არარსებულ უჯრას მიუთითებს (მაგ. ცხრილის საზღვრებს გარეთ).",
  NUM: "რიცხვითი შეცდომა: ფუნქციას გადაეცა დაუშვებელი რიცხვი (მაგ. უარყოფითი რიცხვის ფესვი) ან შედეგი ზედმეტად დიდია.",
  CYCLE: "წრიული მითითება: ფორმულა პირდაპირ ან სხვა უჯრის გავლით საკუთარ თავს მიუთითებს.",
  ERROR: "ფორმულაში სინტაქსური შეცდომაა. შეამოწმეთ ფრჩხილები, ბრჭყალები და გამყოფები: ინგლისურ ფორმატში გამოიყენება „,“ და „.“, ქართულში — „;“ და „,“.",
  SPILL: "შედეგი რამდენიმე უჯრაზე უნდა გაიშალოს, მაგრამ ზოგი მათგანი უკვე შევსებულია. გაასუფთავეთ ეს უჯრები.",
  LIC: "ტექნიკური შეცდომა. გთხოვთ, შეგვატყობინოთ.",
};

/** A Georgian tip, optionally followed by an English-format example formula. */
interface UnsupportedFunctionTip {
  readonly tipKa: string;
  readonly example?: string;
}

/**
 * Excel functions our engine does not have (HyperFormula 3.5.0 plus our custom RANK, RANK.EQ and
 * CONCAT), with a Georgian tip. Learners get "this works in Excel but not here yet" instead of
 * "check the spelling".
 */
export const UNSUPPORTED_EXCEL_FUNCTIONS: Readonly<Record<string, UnsupportedFunctionTip>> = {
  "RANK.AVG": { tipKa: "გამოიყენეთ RANK.EQ:", example: "=RANK.EQ(B2,$B$2:$B$10)" },
  AVERAGEIFS: { tipKa: "გამოიყენეთ SUMIFS(...)/COUNTIFS(...)." },
  XMATCH: { tipKa: "გამოიყენეთ MATCH, ბოლო არგუმენტად 0-ით." },
  SORTBY: { tipKa: "გამოიყენეთ SORT." },
  INDIRECT: { tipKa: "მიუთითეთ უჯრა პირდაპირ, მაგ. B2." },
  LET: { tipKa: "ჩაწერეთ გამოთვლა პირდაპირ, LET-ის გარეშე." },
  LAMBDA: { tipKa: "ჩაწერეთ გამოთვლა პირდაპირ, LAMBDA-ს გარეშე." },
  NUMBERVALUE: { tipKa: "გამოიყენეთ VALUE." },
  FIXED: { tipKa: "გამოიყენეთ ROUND." },
  DOLLAR: { tipKa: "გამოიყენეთ ROUND." },
  TEXTSPLIT: { tipKa: "გამოიყენეთ LEFT, MID და FIND." },
  TEXTBEFORE: { tipKa: "გამოიყენეთ LEFT და FIND." },
  TEXTAFTER: { tipKa: "გამოიყენეთ MID და FIND." },
};

/** The tip for an Excel function our engine lacks, with its example in the learner's format. */
export function unsupportedFunctionTip(name: string, mode: FormatMode = "en"): string | undefined {
  const tip = UNSUPPORTED_EXCEL_FUNCTIONS[name.toUpperCase()];
  if (!tip) return undefined;
  return tip.example ? `${tip.tipKa} ${convertFormula(tip.example, "en", mode)}` : tip.tipKa;
}

const FUNCTION_NOT_RECOGNIZED_RE = /Function name (.+?) not recognized/i;
const NAME_NOT_RECOGNIZED_RE = /Named expression (.+?) not recognized/i;

function explainName(detail: string, mode: FormatMode): string | undefined {
  const fn = FUNCTION_NOT_RECOGNIZED_RE.exec(detail)?.[1];
  if (fn) {
    const name = fn.toUpperCase();
    const tip = unsupportedFunctionTip(name, mode);
    if (tip) return `${name} Excel-ის ფუნქციაა, მაგრამ ამ სიმულატორში ჯერ არ მუშაობს. ${tip}`;
    return `ფუნქცია „${fn}“ არ არსებობს. შეამოწმეთ მართლწერა: ფუნქციები ინგლისურად იწერება (მაგ. SUM, VLOOKUP).`;
  }
  const name = NAME_NOT_RECOGNIZED_RE.exec(detail)?.[1];
  if (name) {
    return `„${name}“ არც ფუნქციაა და არც უჯრის მისამართი. თუ ეს ტექსტია, ჩასვით ბრჭყალებში: "${name}".`;
  }
  return undefined;
}

function refine(type: ExcelErrorType, detail: string, mode: FormatMode): string | undefined {
  switch (type) {
    case "NAME":
      return explainName(detail, mode);
    case "NA":
      return /wrong number of arguments/i.test(detail)
        ? "ფუნქციას არგუმენტების არასწორი რაოდენობა აქვს. შეამოწმეთ გამყოფები (, ან ;) და ფრჩხილები."
        : undefined;
    case "VALUE":
      return /cell range not allowed/i.test(detail)
        ? "აქ დიაპაზონი (მაგ. B2:B6) არ გამოდგება. მიუთითეთ ერთი უჯრა ან გამოიყენეთ ფუნქცია, მაგ. SUM."
        : undefined;
    default:
      return undefined;
  }
}

/**
 * Georgian explanation for an error type. Pass HyperFormula's `detail` message (from
 * `CellError.detail`) to get a more precise text, e.g. the misspelled function's name, and the
 * learner's format mode so example formulas use the right separators.
 */
export function explainError(type: ExcelErrorType, detail = "", mode: FormatMode = "en"): ErrorExplanation {
  return {
    type,
    code: ERROR_CODES[type],
    messageKa: refine(type, detail, mode) ?? MESSAGES_KA[type],
  };
}
