/**
 * Internal: Excel-compatibility functions for HyperFormula 3.5.0, registered once per process.
 * Ported from the technical review's verified engine (review-tech/exp/engine.mjs).
 *
 * - RANK, RANK.EQ: missing in HyperFormula 3.5.0.
 * - CONCAT: missing in HyperFormula 3.5.0.
 * - ROUND: Excel rounds the 15-significant-digit decimal, half away from zero.
 *   HyperFormula's ROUND(1.25*0.18, 2) gives 0.22; Excel gives 0.23.
 * - TODAY: frozen to the task's `sheet.today`, passed in through `config.context`.
 * - DATEDIF: Excel accepts lower-case units ("d", "m", "y", "md"…); HyperFormula only upper case.
 *
 * Import HyperFormula only with ESM `import` (never `require`): two copies of the class would
 * each have their own function registry, and the custom functions would apply to one only.
 */
import {
  CellError,
  CellValueDetailedType,
  EmptyValue,
  ErrorType,
  FunctionArgumentType,
  FunctionPlugin,
  HyperFormula,
  type FunctionPluginDefinition,
  type SimpleRangeValue,
} from "hyperformula";

/** What HyperFormula passes to a plugin method (its AST and state types are not exported). */
type RunFunction = FunctionPlugin["runFunction"];
type PluginAst = { args: Parameters<RunFunction>[0] };
type PluginState = Parameters<RunFunction>[1];

/** Values `config.context` carries into the custom functions. */
export interface EngineContext {
  /** Excel date serial that TODAY() returns (from `sheet.today`). */
  readonly todaySerial?: number;
}

/** Excel-like ROUND: round the 15-significant-digit decimal representation, half away from zero. */
export function excelRound(value: number, digits: number): number {
  if (!Number.isFinite(value)) return value;
  const d = Math.trunc(digits);
  const rounded = Number(`${Math.round(Number(`${Math.abs(value).toPrecision(15)}e${d}`))}e${-d}`);
  return value < 0 ? -rounded : rounded;
}

/** Today's Excel date serial in UTC (used when a task has no `sheet.today`). */
function realTodaySerial(): number {
  return Math.floor(Date.now() / 86_400_000) + 25_569;
}

function textOf(value: unknown): string {
  if (value === EmptyValue || value === null || value === undefined || typeof value === "symbol") return "";
  if (value === true) return "TRUE";
  if (value === false) return "FALSE";
  return String(value);
}

function isRange(value: unknown): value is SimpleRangeValue {
  return typeof (value as SimpleRangeValue | null)?.valuesFromTopLeftCorner === "function";
}

class ExcelCompatPlugin extends FunctionPlugin {
  rankEq(ast: PluginAst, state: PluginState) {
    return this.runFunction(ast.args, state, this.metadata("RANK.EQ"), (n: number, range: SimpleRangeValue, order: number) => {
      const numbers = range.valuesFromTopLeftCorner().filter((v): v is number => typeof v === "number");
      if (!numbers.includes(n)) return new CellError(ErrorType.NA);
      return 1 + numbers.filter((v) => (order ? v < n : v > n)).length;
    });
  }

  concat(ast: PluginAst, state: PluginState) {
    return this.runFunction(ast.args, state, this.metadata("CONCAT"), (...args: unknown[]) =>
      args.flatMap((arg) => (isRange(arg) ? arg.valuesFromTopLeftCorner() : [arg])).map(textOf).join(""),
    );
  }

  round(ast: PluginAst, state: PluginState) {
    return this.runFunction(ast.args, state, this.metadata("ROUND"), (x: number, digits: number) => excelRound(x, digits));
  }

  today(ast: PluginAst, state: PluginState) {
    return this.runFunction(ast.args, state, this.metadata("TODAY"), () => {
      const context = this.config.context as EngineContext | undefined;
      return context?.todaySerial ?? realTodaySerial();
    });
  }
}

const NUMBER = { argumentType: FunctionArgumentType.NUMBER };
const RANGE = { argumentType: FunctionArgumentType.RANGE };

ExcelCompatPlugin.implementedFunctions = {
  "RANK.EQ": { method: "rankEq", parameters: [NUMBER, RANGE, { ...NUMBER, defaultValue: 0 }] },
  RANK: { method: "rankEq", parameters: [NUMBER, RANGE, { ...NUMBER, defaultValue: 0 }] },
  CONCAT: { method: "concat", parameters: [{ argumentType: FunctionArgumentType.ANY }], repeatLastArgs: 1, expandRanges: false },
  ROUND: { method: "round", parameters: [NUMBER, { ...NUMBER, defaultValue: 0 }] },
  TODAY: { method: "today", parameters: [], isVolatile: true, returnNumberType: CellValueDetailedType.NUMBER_DATE },
};

/** HyperFormula's own DATEDIF, with the unit argument upper-cased first. */
const BuiltInDatedif = HyperFormula.getFunctionPlugin("DATEDIF") as FunctionPluginDefinition;
type PluginMethod = (ast: PluginAst, state: PluginState) => ReturnType<RunFunction>;

class DatedifAnyCasePlugin extends BuiltInDatedif {
  constructor(...args: ConstructorParameters<FunctionPluginDefinition>) {
    super(...args);
    const run = this.runFunction;
    this.runFunction = (fnArgs, state, metadata, implementation) =>
      run(fnArgs, state, metadata, (start: unknown, end: unknown, unit: unknown, ...rest: unknown[]) =>
        implementation(start, end, typeof unit === "string" ? unit.toUpperCase() : unit, ...rest),
      );
  }

  // HyperFormula's registry needs the method on the plugin class itself.
  datedif(ast: PluginAst, state: PluginState) {
    return (BuiltInDatedif.prototype as unknown as { datedif: PluginMethod }).datedif.call(this, ast, state);
  }
}
DatedifAnyCasePlugin.implementedFunctions = { DATEDIF: BuiltInDatedif.implementedFunctions.DATEDIF! };

const CUSTOM_IDS = Object.keys(ExcelCompatPlugin.implementedFunctions);
let registered = false;

/** Registers the custom functions with HyperFormula. Safe to call many times. */
export function registerExcelFunctions(): void {
  if (registered) return;
  registered = true;
  const existing = new Set(HyperFormula.getRegisteredFunctionNames("enGB"));
  for (const id of [...CUSTOM_IDS, "DATEDIF"]) {
    if (existing.has(id)) HyperFormula.unregisterFunction(id);
  }
  const names = Object.fromEntries(CUSTOM_IDS.map((id) => [id, id]));
  HyperFormula.registerFunctionPlugin(ExcelCompatPlugin, { enGB: names });
  HyperFormula.registerFunction("DATEDIF", DatedifAnyCasePlugin, { enGB: { DATEDIF: "DATEDIF" } });
}
