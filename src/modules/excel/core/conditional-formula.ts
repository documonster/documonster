/**
 * The formula a conditional-formatting rule needs when the caller did not write one.
 *
 * Several rule types are *specified* in terms of a formula that the file must carry even though the caller
 * expressed the intent some other way: `containsText` with a `text` attribute still needs
 * `NOT(ISERROR(SEARCH("…",A1)))` in the file, and every `timePeriod` rule needs the arithmetic that decides
 * what "last week" means. Excel writes these; a file without them has a rule that never fires.
 *
 * ## Why this is its own module
 *
 * It was inside `cf-rule-xform.ts`, reachable only from the XML writer, and the XLSB writer therefore emitted
 * those rules with `cbFmla1 = 0`. Excel's own `BrtBeginCFRule` for the same rule is 82 bytes against this
 * library's 50, and the 32 missing bytes are exactly the formula — so a `containsText` rule written to XLSB
 * looked complete, matched on nothing, and reported no loss.
 *
 * One list, two containers. A second copy of "what `lastWeek` means" is the shape of mistake this repository
 * has paid for repeatedly: the two would agree until one of them was edited.
 */
import { rangeCreate, rangeTl } from "@excel/core/range";

/** The fields a rule has to expose for its formula to be derived. */
export interface ConditionalRuleLike {
  readonly type?: string;
  readonly operator?: string;
  readonly text?: string;
  readonly timePeriod?: string;
  readonly ref?: string;
  readonly formulae?: readonly unknown[];
}

/**
 * Whether a formula is absent: not a string or number, or a string with nothing but whitespace in it.
 *
 * Whitespace counts as absent because it is not a formula in either container: the XLSX writer used to write
 * `<formula> </formula>`, a condition Excel cannot evaluate, while the XLSB writer failed to tokenize the same text
 * and refused the rule — so one model produced a broken file in one container and a refusal in the other.
 */
function isBlankFormula(formula: unknown): boolean {
  return typeof formula === "string" ? formula.trim() === "" : typeof formula !== "number";
}

/** The rule's own formula, when it has one. */
function stated(rule: ConditionalRuleLike): string | undefined {
  const first = rule.formulae?.[0];
  return typeof first === "string" && !isBlankFormula(first) ? first : undefined;
}

/**
 * The eight conditions of the text-matching family, by the XLSX `type` that names each.
 *
 * The model spells them two ways, and both are accepted everywhere: as the rule's own `type` (`"beginsWith"`, which
 * is how XLSX names the rule, and how readers return `notContainsText`, `beginsWith` and `endsWith`) or as
 * `type: "containsText"` with the condition in `operator`, the shape `ContainsTextRuleType` declares.
 * `"notContains"` — XLSX's spelling of the *operator* attribute for `notContainsText` — is accepted as the same
 * condition.
 */
export type TextRuleKind =
  | "containsText"
  | "notContainsText"
  | "beginsWith"
  | "endsWith"
  | "containsBlanks"
  | "notContainsBlanks"
  | "containsErrors"
  | "notContainsErrors";

const TEXT_RULE_KINDS: ReadonlySet<string> = new Set<TextRuleKind>([
  "containsText",
  "notContainsText",
  "beginsWith",
  "endsWith",
  "containsBlanks",
  "notContainsBlanks",
  "containsErrors",
  "notContainsErrors"
]);

/** The text-family condition a rule states, or `undefined` when it is not one of them. */
export function textRuleKind(rule: ConditionalRuleLike): TextRuleKind | undefined {
  const asKind = (value: string | undefined): TextRuleKind | undefined =>
    value === "notContains"
      ? "notContainsText"
      : value !== undefined && TEXT_RULE_KINDS.has(value)
        ? (value as TextRuleKind)
        : undefined;
  if (rule.type === "containsText") {
    // A bare `containsText` with no operator is the plain condition its name says.
    return asKind(rule.operator) ?? "containsText";
  }
  return asKind(rule.type);
}

/**
 * The `type` and `operator` a reader returns for a text-family condition — the shape the XLSX reader has always
 * produced, so the binary reader gives the same model for the same rule.
 *
 * Five conditions fold onto `type: "containsText"` with the XLSX type in `operator`; the other three keep the XLSX
 * spelling of both, `notContainsText`'s operator attribute being `notContains`.
 */
export function textRuleReadShape(kind: TextRuleKind): { type: string; operator: string } {
  switch (kind) {
    case "notContainsText":
      return { type: "notContainsText", operator: "notContains" };
    case "beginsWith":
    case "endsWith":
      return { type: kind, operator: kind };
    default:
      return { type: "containsText", operator: kind };
  }
}

/** Whether a text-family condition carries a search string (`text`), which four of the eight do. */
export function textRuleHasSearchString(kind: TextRuleKind): boolean {
  return (
    kind === "containsText" ||
    kind === "notContainsText" ||
    kind === "beginsWith" ||
    kind === "endsWith"
  );
}

/** A string as a formula literal: quoted, with each quote doubled, as Excel escapes one. */
function stringLiteral(text: string): string {
  return `"${text.replaceAll('"', '""')}"`;
}

/**
 * The top-left cell of the rule's range, which every derived formula is written relative to.
 *
 * Excel writes the formula against the first cell and applies it to the rest by offset, which is why the
 * binary form stores it as a relative `PtgRefN`.
 */
function topLeft(rule: ConditionalRuleLike): string | undefined {
  return rule.ref === undefined ? undefined : rangeTl(rangeCreate(rule.ref));
}

/** The formula for the text-matching family, or `undefined` when the rule is not one of them. */
export function textRuleFormula(rule: ConditionalRuleLike): string | undefined {
  const own = stated(rule);
  if (own !== undefined) {
    return own;
  }
  const tl = topLeft(rule);
  if (tl === undefined) {
    return undefined;
  }
  // **Quotes in the search string are doubled.** Interpolated raw, `q"z` produced `SEARCH("q"z",C5)` — a formula
  // that does not parse, so the XLSX file carried a broken condition and the XLSB writer refused the rule.
  const text = stringLiteral(rule.text ?? "");
  switch (textRuleKind(rule)) {
    case "containsText":
      return `NOT(ISERROR(SEARCH(${text},${tl})))`;
    case "notContainsText":
      return `ISERROR(SEARCH(${text},${tl}))`;
    case "beginsWith":
      return `LEFT(${tl},LEN(${text}))=${text}`;
    case "endsWith":
      return `RIGHT(${tl},LEN(${text}))=${text}`;
    case "containsBlanks":
      return `LEN(TRIM(${tl}))=0`;
    case "notContainsBlanks":
      return `LEN(TRIM(${tl}))>0`;
    case "containsErrors":
      return `ISERROR(${tl})`;
    case "notContainsErrors":
      return `NOT(ISERROR(${tl}))`;
    default:
      return undefined;
  }
}

/** The formula for a `timePeriod` rule, or `undefined` when the period has none. */
export function timePeriodRuleFormula(rule: ConditionalRuleLike): string | undefined {
  const own = stated(rule);
  if (own !== undefined) {
    return own;
  }
  const tl = topLeft(rule);
  if (tl === undefined) {
    return undefined;
  }
  switch (rule.timePeriod) {
    case "thisWeek":
      return `AND(TODAY()-ROUNDDOWN(${tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${tl},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
    case "lastWeek":
      return `AND(TODAY()-ROUNDDOWN(${tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${tl},0)<(WEEKDAY(TODAY())+7))`;
    case "nextWeek":
      return `AND(ROUNDDOWN(${tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${tl},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
    case "yesterday":
      return `FLOOR(${tl},1)=TODAY()-1`;
    case "today":
      return `FLOOR(${tl},1)=TODAY()`;
    case "tomorrow":
      return `FLOOR(${tl},1)=TODAY()+1`;
    case "last7Days":
      return `AND(TODAY()-FLOOR(${tl},1)<=6,FLOOR(${tl},1)<=TODAY())`;
    case "lastMonth":
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0-1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0-1)))`;
    case "thisMonth":
      return `AND(MONTH(${tl})=MONTH(TODAY()),YEAR(${tl})=YEAR(TODAY()))`;
    case "nextMonth":
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0+1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0+1)))`;
    default:
      return undefined;
  }
}

/**
 * The formula for any rule that derives one, or `undefined`.
 *
 * The text family is recognised in either spelling — see {@link textRuleKind} — so a caller does not have to
 * normalise first.
 */
export function derivedRuleFormula(rule: ConditionalRuleLike): string | undefined {
  if (rule.type === "timePeriod") {
    return timePeriodRuleFormula(rule);
  }
  if (textRuleKind(rule) !== undefined) {
    return textRuleFormula(rule);
  }
  return stated(rule);
}

/**
 * Why a rule cannot be written, when it states its condition only through formulas and does not have them.
 *
 * A comparison (`cellIs`) and an `expression` have nothing else to derive a condition from: with no operand the
 * first compares against nothing and the second has no test, and a container cannot express either as a rule that
 * means anything. Both writers ask this one question, so they skip and report the same rules — the XLSX writer used to
 * fail on such a rule with a `TypeError`, and the XLSB writer wrote it with no formula at all.
 *
 * @returns `undefined` when the rule has what it needs.
 */
export function missingRuleFormula(rule: ConditionalRuleLike): string | undefined {
  if (rule.type !== "cellIs" && rule.type !== "expression") {
    return undefined;
  }
  const needed =
    rule.type === "cellIs" && (rule.operator === "between" || rule.operator === "notBetween")
      ? 2
      : 1;
  const present = (rule.formulae ?? []).filter(formula => !isBlankFormula(formula)).length;
  if (present >= needed) {
    return undefined;
  }
  return `${rule.type} rule needs ${needed} formula(s) and has ${present}`;
}
