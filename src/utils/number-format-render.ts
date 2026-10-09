/**
 * Render a value through an Excel number format.
 *
 * **One renderer for every caller.** A cell's display text and the `TEXT` worksheet function
 * format a number the same way in Excel, but this library used to carry two implementations:
 * `excel/utils/cell-format.ts` for display and a second copy inside the formula engine. They had
 * different rounding, different fraction search, different month/minute rules and different ideas
 * of what a literal is — `TEXT` read `0" hours"` as a date and dropped every literal from a
 * percentage, while the display path placed the minus sign after a currency symbol. The formula
 * engine sits below the Excel module and could not import the display renderer, so the shared one
 * lives here, at Layer 0, where both can reach it.
 *
 * It renders from the tokens of `./number-format`, so the question "is this character a code or a
 * literal" has one answer for rendering, for date detection on load, for date-kind classification
 * and for input parsing.
 */

import { serialToParts } from "./excel-serial";
import type { FormatCompareOp, FormatDatePart, FormatSection, FormatToken } from "./number-format";
import { DAY_NAMES, MONTH_NAMES, parseNumberFormat } from "./number-format";

export interface NumberFormatRenderOptions {
  /**
   * The workbook's epoch, which date codes read the serial in. Defaults to the 1900 epoch. A serial made
   * from a `Date` has to be made with the same epoch passed here: calendar fields would survive a mismatch,
   * but an elapsed format reads the whole serial and would be off by 1,462 days.
   */
  readonly date1904?: boolean;
}

/** What Excel shows for a value its format cannot display: a date past 9999 or a number past `double`. */
const NOT_DISPLAYABLE = "########";

/** Serial of 10000-01-01 under the 1900 epoch; a date at or past it has no display. */
const FIRST_SERIAL_PAST_9999 = 2958466;

/** Days between the epochs: the same calendar date is this much lower under the 1904 epoch. */
const EPOCH_1904_OFFSET = 1462;

/** Excel accepts at most three decimal places of a second (`ss.000`); more is an invalid format. */
const MAX_SECOND_DECIMALS = 3;

/**
 * Format `value` as Excel would display it under `fmt`.
 *
 * Text is shown through the format's text section (the fourth, or a trailing section carrying
 * `@`), or as is when there is none. Booleans are `TRUE`/`FALSE`. An empty format is `General`.
 * A value the format cannot show — a date past 9999 or before 1900, a negative time, a conditional
 * format no section of which matches — is shown as hashes, as Excel shows it in a cell.
 */
export function renderNumberFormat(
  fmt: string,
  value: number | string | boolean,
  options: NumberFormatRenderOptions = {}
): string {
  return tryRenderNumberFormat(fmt, value, options) ?? NOT_DISPLAYABLE;
}

/**
 * {@link renderNumberFormat}, but `undefined` for a value the format cannot show — which is where
 * Excel's `TEXT()` returns `#VALUE!` and a cell shows hashes.
 */
export function tryRenderNumberFormat(
  fmt: string,
  value: number | string | boolean,
  options: NumberFormatRenderOptions = {}
): string | undefined {
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  const parsed = parseNumberFormat(fmt.trim() === "" ? "General" : fmt);
  if (typeof value === "string") {
    return parsed.textSection ? renderTextSection(parsed.textSection.tokens, value) : value;
  }
  if (!Number.isFinite(value)) {
    return String(value);
  }

  const sections = parsed.numberSections;
  if (sections.length === 0) {
    // A text-only format such as `"n="@` says nothing about numbers, which are shown as General.
    return formatGeneral(value);
  }
  const choice = chooseSection(sections, value);
  if (choice === undefined) {
    return undefined;
  }
  const { section, signed } = choice;
  const shown = signed ? value : Math.abs(value);
  if (section.hasGeneral) {
    return renderGeneralSection(section.tokens, shown);
  }
  if (section.hasDate || section.hasTime || section.hasElapsed) {
    // A negative value has no date or clock reading, whichever section it reached. A lone elapsed count
    // (`[h]`, `[ss]`) is the exception: it is a plain signed number of units.
    if (value < 0 && !isLoneElapsed(section.tokens)) {
      return undefined;
    }
    return renderDateSection(section.tokens, shown, options.date1904 === true);
  }
  return renderNumberSection(section.tokens, shown);
}

function isLoneElapsed(tokens: readonly FormatToken[]): boolean {
  return (
    tokens.filter(t => t.kind === "elapsed").length === 1 &&
    !tokens.some(t => t.kind === "date" || t.kind === "ampm")
  );
}

// =============================================================================
// Section choice
// =============================================================================

function compare(value: number, op: FormatCompareOp, threshold: number): boolean {
  switch (op) {
    case "=":
      return value === threshold;
    case "<>":
      return value !== threshold;
    case "<":
      return value < threshold;
    case "<=":
      return value <= threshold;
    case ">":
      return value > threshold;
    case ">=":
      return value >= threshold;
  }
}

/** A condition no non-negative number can meet: `[<0]`, `[<-5]`, `[<=-1]`, `[=-5]`. */
function negativeOnly(c: { op: FormatCompareOp; value: number }): boolean {
  return (c.op === "<" && c.value <= 0) || ((c.op === "<=" || c.op === "=") && c.value < 0);
}

/**
 * Pick the section for a number, and say whether it keeps its sign; `undefined` when no section applies.
 *
 * Without conditions the roles are positional: a negative number takes the second section, which prints
 * its own sign, and zero the third. With a condition, the rules below reproduce Excel exactly over 160
 * measured combinations of condition, section count and value:
 *
 * - A conditioned section that matches is used. It drops the sign only when its condition admits no
 *   non-negative number (`[<0]`, `[<=-5]`); `[<100]` and `[<=0]` keep it.
 * - A bare first section beside a conditioned second means `[>0]`.
 * - With three or more sections a bare second one means `[<0]` and the third takes everything else.
 * - With two sections a bare second one takes everything else. A negative number there drops its sign when
 *   the first condition covered every positive number or only negatives (`[>0]`, `[>=-5]`, `[<-10]`), and
 *   keeps it otherwise (`[>100]`, `[=-5]`).
 * - Two conditions and no third section: a number matching neither has no display.
 */
function chooseSection(
  sections: readonly FormatSection[],
  value: number
): { section: FormatSection; signed: boolean } | undefined {
  const [first, second, third] = sections;
  if (first.condition === undefined && second?.condition === undefined) {
    if (sections.length === 1) {
      return { section: first, signed: true };
    }
    if (value < 0) {
      return { section: second, signed: false };
    }
    return { section: value === 0 && third ? third : first, signed: true };
  }

  const c1 = first.condition ?? { op: ">" as const, value: 0 };
  if (compare(value, c1.op, c1.value)) {
    return { section: first, signed: !negativeOnly(c1) };
  }
  if (second === undefined) {
    return { section: first, signed: true };
  }
  const c2 = second.condition;
  if (c2) {
    if (compare(value, c2.op, c2.value)) {
      return { section: second, signed: !negativeOnly(c2) };
    }
    return third ? { section: third, signed: true } : undefined;
  }
  if (third) {
    return value < 0 ? { section: second, signed: false } : { section: third, signed: true };
  }
  const coversPositives = (c1.op === ">" || c1.op === ">=") && c1.value <= 0;
  const dropsSign = value < 0 && (coversPositives || (c1.op !== "=" && negativeOnly(c1)));
  return { section: second, signed: !dropsSign };
}

// =============================================================================
// Shared pieces
// =============================================================================

/** What a token shows when it is not doing its job as a code — or is a pure layout token. */
function plainText(token: FormatToken): string {
  switch (token.kind) {
    case "literal":
    case "numeral":
      return token.text;
    case "pad":
      return " ";
    case "locale":
      return token.symbol;
    case "digit":
      return token.char === "0" ? "0" : token.char === "?" ? " " : "";
    case "decimal":
      return ".";
    case "thousands":
      return ",";
    case "percent":
      return "%";
    case "slash":
      return "/";
    case "exponent":
      return (token.upper ? "E" : "e") + token.sign;
    default:
      // fill, colour, condition, other brackets, and codes with no meaning here.
      return "";
  }
}

/**
 * Render text through a text section. The caller decides whether there is one; an *empty* text section
 * (`0;0;0;`) is a real section that shows nothing, not an absent one that shows the text as is.
 */
function renderTextSection(tokens: readonly FormatToken[], text: string): string {
  let out = "";
  for (const token of tokens) {
    out += token.kind === "text" ? text : plainText(token);
  }
  return out;
}

/** Trailing zeros of a decimal fraction, and the point itself when nothing follows it. */
function stripDecimal(s: string): string {
  return s.includes(".") ? s.replace(/(?:\.0*|(\.\d*[1-9])0+)$/, "$1") : s;
}

/**
 * Excel's `General`: at most eleven characters, not counting a minus sign, switching to scientific
 * notation with six significant digits when a number cannot fit — `0.333333333`, `1.23457E+11`,
 * `1234567891` for 1234567890.5. The shape of the algorithm is SheetJS's, which its authors calibrated
 * against Excel; the rounding is decimal, at fifteen significant digits, so a half like -501059677.45 rounds
 * away from zero as Excel does rather than as its nearest double would. It agreed with Excel on all 146
 * values measured.
 */
export function formatGeneral(value: number): string {
  if (value === 0) {
    return "0";
  }
  const sign = value < 0 ? "-" : "";
  const x = Math.abs(value);
  const fixed = (decimals: number): string => {
    const { int, frac } = toFixedDigits(x, decimals);
    return (int || "0") + (decimals > 0 ? "." + frac : "");
  };
  const magnitude = Math.floor(Math.log10(x));
  let s: string;
  if (magnitude >= -4 && magnitude <= -1) {
    s = fixed(9);
  } else if (Math.abs(magnitude) <= 9) {
    s = stripDecimal(fixed(12));
    if (s.length > 11) {
      s = fixed(9 - magnitude);
      if (s.split(".")[0].length > Math.max(magnitude, 0) + 1) {
        // Rounding carried into the next magnitude: 9999999999.9 is 10000000000.
        return sign + formatGeneral(Number(s));
      }
      if (stripDecimal(s).length > 11) {
        s = x.toExponential(5);
      }
    }
  } else if (magnitude === 10) {
    // Eleven integer digits fill the width: round to a whole number, which may carry to twelve digits.
    const rounded = Number(toFixedDigits(x, 0).int);
    return sign + (rounded >= 1e11 ? formatGeneral(rounded) : String(rounded));
  } else {
    s = stripDecimal(fixed(11));
    if (s.length > 11 || s === "0") {
      s = x.toExponential(5);
    }
  }
  s = s.toUpperCase();
  if (s.includes("E")) {
    // At least two exponent digits, and no trailing zeros in the mantissa: `1.5E+15`, `1E-12`.
    s = s.replace(/(?:\.0*|(\.\d*[1-9])0+)E/, "$1E").replace(/E([+-])(\d)$/, "E$10$2");
  }
  return sign + stripDecimal(s);
}

function renderGeneralSection(tokens: readonly FormatToken[], value: number): string {
  let out = "";
  for (const token of tokens) {
    out += token.kind === "general" ? formatGeneral(value) : plainText(token);
  }
  return out;
}

/**
 * A non-negative number as decimal digit strings, rounded half away from zero at `decimals`
 * places, and taken to fifteen significant digits first as Excel does — so `1.005` rounds to
 * `1.01` rather than to the `1.00` its binary value would give.
 *
 * `int` has no leading zeros and is empty for a value below one. With `shift`, the value is first divided
 * by `10^shift` exactly, by moving the decimal point.
 */
function toFixedDigits(x: number, decimals: number, shift = 0): { int: string; frac: string } {
  if (x === 0) {
    return { int: "", frac: "0".repeat(decimals) };
  }
  const [mantissa, exponent] = x.toExponential(14).split("e");
  let digits = mantissa.replace(".", "");
  // `shift` divides by 10^shift by moving the point in the digit string — never by dividing a `double`,
  // which underflows to 0 for the smallest values and turned 5e-324 into `0.InE-324`.
  let point = Number(exponent) + 1 - shift;
  const keep = point + decimals;
  if (keep < 0) {
    return { int: "", frac: "0".repeat(decimals) };
  }
  if (keep < digits.length) {
    const roundUp = digits.charCodeAt(keep) >= 53;
    const kept = digits.slice(0, keep).split("");
    if (roundUp) {
      let i = kept.length - 1;
      while (i >= 0 && kept[i] === "9") {
        kept[i] = "0";
        i--;
      }
      if (i >= 0) {
        kept[i] = String(Number(kept[i]) + 1);
      } else {
        kept.unshift("1");
        point++;
      }
    }
    digits = kept.join("");
  }
  const full = point <= 0 ? "0".repeat(-point) + digits : digits;
  const at = Math.max(point, 0);
  return {
    int: full.slice(0, at).padEnd(at, "0").replace(/^0+/, ""),
    frac: full.slice(at).padEnd(decimals, "0").slice(0, decimals)
  };
}

function isDigitChar(s: string | undefined): boolean {
  return s !== undefined && s >= "0" && s <= "9";
}

/**
 * Lay an integer's digits into a run of placeholders, right-aligned.
 *
 * Returns one string per placeholder. A position with no digit shows `0` for `0`, a space for `?`
 * and nothing for `#`; digits beyond the placeholders go in front of the first one, which is how
 * `0` displays `12345`. With `grouping`, a comma follows every third digit from the right.
 */
function layInteger(
  digits: string,
  slots: readonly ("0" | "#" | "?")[],
  grouping: boolean
): string[] {
  const count = slots.length;
  const charAt = (p: number): string => {
    if (p < digits.length) {
      return digits[digits.length - 1 - p];
    }
    if (p >= count) {
      return "";
    }
    const slot = slots[count - 1 - p];
    return slot === "0" ? "0" : slot === "?" ? " " : "";
  };
  const withComma = (p: number): string => {
    const ch = charAt(p);
    return grouping && p > 0 && p % 3 === 0 && isDigitChar(ch) && isDigitChar(charAt(p - 1))
      ? ch + ","
      : ch;
  };
  const parts: string[] = [];
  for (let j = 0; j < count; j++) {
    let part = "";
    if (j === 0) {
      for (let p = digits.length - 1; p >= count; p--) {
        part += withComma(p);
      }
    }
    part += withComma(count - 1 - j);
    parts.push(part);
  }
  return parts;
}

/**
 * Lay fraction digits into placeholders, left-aligned. Trailing zeros under `#` vanish and under
 * `?` become spaces; a `0` placeholder, or any non-zero digit, stops the trimming.
 */
function layFraction(digits: string, slots: readonly ("0" | "#" | "?")[]): string[] {
  const parts = slots.map((_, k) => digits[k] ?? "0");
  for (let k = slots.length - 1; k >= 0; k--) {
    if (parts[k] !== "0" || slots[k] === "0") {
      break;
    }
    parts[k] = slots[k] === "?" ? " " : "";
  }
  return parts;
}

function slotOf(token: FormatToken): "0" | "#" | "?" {
  return token.kind === "digit" ? token.char : "0";
}

// =============================================================================
// Numbers
// =============================================================================

function renderNumberSection(tokens: readonly FormatToken[], value: number): string | undefined {
  const negative = value < 0;
  let x = Math.abs(value);

  const expIndex = tokens.findIndex(t => t.kind === "exponent");
  const fraction = expIndex === -1 ? findFraction(tokens) : undefined;
  const numberEnd = expIndex === -1 ? tokens.length : expIndex;
  let decimalIndex = -1;
  for (let k = 0; k < numberEnd; k++) {
    if (tokens[k].kind === "decimal") {
      decimalIndex = k;
      break;
    }
  }
  const intEnd = decimalIndex === -1 ? numberEnd : decimalIndex;
  const lastDigitBefore = (end: number, start: number): number => {
    for (let k = end - 1; k >= start; k--) {
      if (tokens[k].kind === "digit") {
        return k;
      }
    }
    return -1;
  };
  const lastIntDigit = lastDigitBefore(intEnd, 0);
  const lastFracDigit = decimalIndex === -1 ? -1 : lastDigitBefore(numberEnd, decimalIndex + 1);

  // A comma that ends a run of digit placeholders divides by a thousand — at the end of the integer
  // part (`0,.0`) or of the fraction (`0.0,,"M"`). One with integer placeholders after it groups.
  let grouping = false;
  let scale = 0;
  for (let k = 0; k < numberEnd; k++) {
    if (tokens[k].kind !== "thousands") {
      continue;
    }
    if (fraction) {
      // Only the whole-number part of a fraction groups (`#,##0 ?/?`); a fraction does not scale.
      grouping ||= fraction.integer.length > 0 && k < fraction.numerator[0];
      continue;
    }
    const partEnd = k < intEnd ? lastIntDigit : lastFracDigit;
    if (partEnd !== -1 && k > partEnd) {
      scale++;
    } else if (k < intEnd) {
      grouping = true;
    }
  }
  // Each `%` multiplies by 100 and each scaling comma divides by 1,000. Zero stays zero, however many.
  if (x !== 0) {
    x =
      (x * Math.pow(100, tokens.filter(t => t.kind === "percent").length)) / Math.pow(1000, scale);
  }
  if (!Number.isFinite(x)) {
    return undefined;
  }

  if (fraction) {
    return renderFraction(tokens, fraction, x, negative, grouping);
  }
  if (expIndex !== -1) {
    return renderScientific(tokens, expIndex, decimalIndex, x, negative);
  }

  const intSlots: ("0" | "#" | "?")[] = [];
  const fracSlots: ("0" | "#" | "?")[] = [];
  tokens.forEach((t, k) => {
    if (t.kind === "digit") {
      (decimalIndex !== -1 && k > decimalIndex ? fracSlots : intSlots).push(t.char);
    }
  });
  const fixed = toFixedDigits(x, fracSlots.length);
  // A value that rounds to zero shows no sign. A section with no digit placeholder rounds nothing, so it
  // keeps the sign: `"x"` shows -5 as `-x`, as Excel does.
  const isZero =
    (intSlots.length > 0 || fracSlots.length > 0) && fixed.int === "" && !/[1-9]/.test(fixed.frac);
  const intParts = layInteger(fixed.int, intSlots, grouping);
  const fracParts = layFraction(fixed.frac, fracSlots);

  let out = negative && !isZero ? "-" : "";
  let intCursor = 0;
  let fracCursor = 0;
  tokens.forEach((token, k) => {
    switch (token.kind) {
      case "digit":
        out +=
          decimalIndex !== -1 && k > decimalIndex ? fracParts[fracCursor++] : intParts[intCursor++];
        break;
      case "decimal":
        if (k === decimalIndex && intSlots.length === 0) {
          // `.00` still shows the integer part of 12.5: there is nowhere else to put it.
          out += grouping ? fixed.int.replace(/\B(?=(\d{3})+$)/g, ",") : fixed.int;
        }
        out += ".";
        break;
      case "thousands":
        break;
      case "text":
        out += formatGeneral(Math.abs(value));
        break;
      default:
        out += plainText(token);
    }
  });
  return out;
}

function renderScientific(
  tokens: readonly FormatToken[],
  expIndex: number,
  decimalIndex: number,
  x: number,
  negative: boolean
): string {
  const intSlots: ("0" | "#" | "?")[] = [];
  const fracSlots: ("0" | "#" | "?")[] = [];
  const expSlots: ("0" | "#" | "?")[] = [];
  tokens.forEach((t, k) => {
    if (t.kind !== "digit") {
      return;
    }
    if (k > expIndex) {
      expSlots.push(t.char);
    } else if (decimalIndex !== -1 && decimalIndex < expIndex && k > decimalIndex) {
      fracSlots.push(t.char);
    } else {
      intSlots.push(t.char);
    }
  });

  // With several integer placeholders the exponent is a multiple of their count: `##0.0E+0`
  // is engineering notation.
  const step = Math.max(intSlots.length, 1);
  let exp = 0;
  let mantissa = { int: "", frac: "0".repeat(fracSlots.length) };
  if (x !== 0) {
    const magnitude = Number(x.toExponential(14).split("e")[1]);
    exp = intSlots.length === 0 ? magnitude + 1 : Math.floor(magnitude / step) * step;
    mantissa = toFixedDigits(x, fracSlots.length, exp);
    if (mantissa.int.length > (intSlots.length === 0 ? 0 : step)) {
      // Rounding carried into a new digit (9.999 → 10.00): move up one exponent step.
      exp += intSlots.length === 0 ? 1 : step;
      mantissa = toFixedDigits(x, fracSlots.length, exp);
    }
  }
  const isZero = mantissa.int === "" && !/[1-9]/.test(mantissa.frac);
  const intParts = layInteger(mantissa.int, intSlots, false);
  const fracParts = layFraction(mantissa.frac, fracSlots);
  const expParts = layInteger(String(Math.abs(exp)).replace(/^0$/, ""), expSlots, false);

  let out = negative && !isZero ? "-" : "";
  let intCursor = 0;
  let fracCursor = 0;
  let expCursor = 0;
  tokens.forEach((token, k) => {
    switch (token.kind) {
      case "digit":
        if (k > expIndex) {
          out += expParts[expCursor++];
        } else if (decimalIndex !== -1 && decimalIndex < expIndex && k > decimalIndex) {
          out += fracParts[fracCursor++];
        } else {
          out += intParts[intCursor++];
        }
        break;
      case "exponent":
        if (k === expIndex) {
          out += (token.upper ? "E" : "e") + (exp < 0 ? "-" : token.sign === "+" ? "+" : "");
        } else {
          out += plainText(token);
        }
        break;
      case "thousands":
        break;
      default:
        out += plainText(token);
    }
  });
  return out;
}

interface FractionLayout {
  /** Token indices of the whole-number placeholders; empty for an improper fraction. */
  readonly integer: readonly number[];
  readonly numerator: readonly number[];
  readonly slash: number;
  /** Token indices of the denominator placeholders, or of the digits of a fixed denominator. */
  readonly denominator: readonly number[];
  readonly fixedDenominator: number | undefined;
}

/** A `?/?`, `# ??/??` or `# ?/8` layout, if the section has one. */
function findFraction(tokens: readonly FormatToken[]): FractionLayout | undefined {
  for (let slash = 1; slash < tokens.length; slash++) {
    if (tokens[slash].kind !== "slash" || tokens[slash - 1].kind !== "digit") {
      continue;
    }
    let numStart = slash - 1;
    while (numStart > 0 && tokens[numStart - 1].kind === "digit") {
      numStart--;
    }
    const numerator: number[] = [];
    for (let k = numStart; k < slash; k++) {
      numerator.push(k);
    }
    const denominator: number[] = [];
    let fixedDenominator: number | undefined;
    if (tokens[slash + 1]?.kind === "numeral") {
      // A fixed denominator is written as bare digits (`?/16`); a quoted or escaped digit is text.
      let text = "";
      for (let k = slash + 1; k < tokens.length; k++) {
        const t = tokens[k];
        if (t.kind === "numeral") {
          text += t.text;
        } else if (t.kind === "digit" && t.char === "0") {
          text += "0";
        } else {
          break;
        }
        denominator.push(k);
      }
      fixedDenominator = Number(text);
    } else {
      for (let k = slash + 1; k < tokens.length && tokens[k].kind === "digit"; k++) {
        denominator.push(k);
      }
    }
    if (denominator.length === 0) {
      continue;
    }
    const integer: number[] = [];
    for (let k = 0; k < numStart; k++) {
      if (tokens[k].kind === "digit") {
        integer.push(k);
      }
    }
    return { integer, numerator, slash, denominator, fixedDenominator };
  }
  return undefined;
}

/**
 * The fraction Excel shows for `x ≥ 0` with `d` no larger than `maxDenominator`: the last continued-fraction
 * convergent within the bound. It is not always the closest fraction — Excel shows 0.14159 as 1/7, not
 * 14/99 — and it matched Excel on all 150 values measured. A few dozen steps for any bound, where trying every
 * denominator cost a billion iterations per cell for `?/?????????`.
 */
function closestFraction(x: number, maxDenominator: number): { n: number; d: number } {
  let p0 = 0;
  let q0 = 1;
  let p1 = 1;
  let q1 = 0;
  let r = x;
  for (let step = 0; step < 64; step++) {
    const a = Math.floor(r);
    const q2 = q0 + a * q1;
    if (q2 > maxDenominator) {
      break;
    }
    const p2 = p0 + a * p1;
    p0 = p1;
    q0 = q1;
    p1 = p2;
    q1 = q2;
    const rest = r - a;
    if (rest === 0 || p1 / q1 === x) {
      break;
    }
    r = 1 / rest;
  }
  return { n: p1, d: q1 };
}

/** A non-negative integer's decimal digits, without exponent notation however large it is. */
function integerDigits(n: number): string {
  return n === 0 ? "0" : toFixedDigits(n, 0).int;
}

function renderFraction(
  tokens: readonly FormatToken[],
  layout: FractionLayout,
  x: number,
  negative: boolean,
  grouping: boolean
): string {
  const mixed = layout.integer.length > 0;
  let whole = mixed ? Math.floor(x) : 0;
  const part = mixed ? x - whole : x;
  let n: number;
  let d: number;
  if (layout.fixedDenominator !== undefined) {
    d = layout.fixedDenominator;
    n = Math.round(part * d);
  } else {
    ({ n, d } = closestFraction(part, Math.pow(10, layout.denominator.length) - 1));
  }
  if (mixed && n === d) {
    whole++;
    n = 0;
  }

  const isZero = whole === 0 && n === 0;
  // A whole number shows no fraction: Excel blanks it to the width it would have taken.
  const blank = n === 0 && mixed;
  const wholeDigits = whole > 0 ? integerDigits(whole) : n === 0 ? "0" : "";
  const intParts = layInteger(
    wholeDigits,
    layout.integer.map(k => slotOf(tokens[k])),
    grouping
  );
  const numParts = layInteger(
    integerDigits(n),
    layout.numerator.map(k => slotOf(tokens[k])),
    false
  );
  // The first denominator placeholder decides how unused places are filled: `0` pads zeros in front
  // (`1/02`), `?` pads spaces behind so fractions line up (`1/2 `), `#` pads nothing. Padding never goes
  // after the digits as zeros, which would change the value: 1/2 is not 1/20.
  const denSlots = layout.denominator.map(k => slotOf(tokens[k]));
  let denText = integerDigits(d);
  if (layout.fixedDenominator === undefined && denText.length < denSlots.length) {
    if (denSlots[0] === "0") {
      denText = denText.padStart(denSlots.length, "0");
    } else if (denSlots[0] === "?") {
      denText = denText.padEnd(denSlots.length, " ");
    }
  }

  const numStart = layout.numerator[0];
  const denStart = layout.denominator[0];
  const denEnd = layout.denominator[layout.denominator.length - 1];

  let out = negative && !isZero ? "-" : "";
  let intCursor = 0;
  let numCursor = 0;
  tokens.forEach((token, k) => {
    if (k >= numStart && k < layout.slash) {
      out += blank ? " ".repeat(Math.max(numParts[numCursor].length, 1)) : numParts[numCursor];
      numCursor++;
      return;
    }
    if (k === layout.slash) {
      out += blank ? " " : "/";
      return;
    }
    if (k >= denStart && k <= denEnd) {
      if (k === denStart) {
        out += blank ? " ".repeat(Math.max(denText.length, layout.denominator.length)) : denText;
      }
      return;
    }
    if (layout.integer.includes(k)) {
      out += intParts[intCursor++];
      return;
    }
    if (token.kind === "thousands") {
      return;
    }
    out += plainText(token);
  });
  return out;
}

// =============================================================================
// Dates and times
// =============================================================================

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

function renderDateSection(
  tokens: readonly FormatToken[],
  value: number,
  date1904: boolean
): string | undefined {
  if (value < 0) {
    // Only a lone elapsed count reaches here negative (see `tryRenderNumberFormat`).
    const magnitude = renderDateSection(tokens, -value, date1904);
    return magnitude === undefined ? undefined : "-" + magnitude;
  }

  let precision = 0;
  for (const t of tokens) {
    if (t.kind === "date" || t.kind === "elapsed") {
      if (t.fraction > MAX_SECOND_DECIMALS) {
        return undefined;
      }
      precision = Math.max(precision, t.fraction);
    }
  }
  // The day and the time of day are separated before scaling, so precision multiplies only a value below
  // 86400: scaling the whole serial overflowed a safe integer, and a valid year-2447 date showed hashes at
  // six decimals.
  const scale = Math.pow(10, precision);
  const unitsPerDay = 86400 * scale;
  let day = Math.floor(value);
  // Rounded to the shown precision of a second even when no time is shown, as Excel does: 23:59:59.95
  // is the next day under `yyyy-mm-dd`.
  let units = Math.round((value - day) * unitsPerDay);
  if (units >= unitsPerDay) {
    day++;
    units -= unitsPerDay;
  }
  const subSecond = String(units % scale).padStart(precision, "0");
  const secondOfDay = Math.floor(units / scale);
  const totalSeconds = day * 86400 + secondOfDay;
  // Checked after rounding, so 9999-12-31 23:59:59.9 is past the end too, and against the workbook's own
  // epoch: the last displayable serial is 1462 lower under 1904. A lone elapsed count (`[ss]`) is a plain
  // number and has no such bound; beside a clock or calendar field it does, as in Excel.
  if (
    !isLoneElapsed(tokens) &&
    day >= FIRST_SERIAL_PAST_9999 - (date1904 ? EPOCH_1904_OFFSET : 0)
  ) {
    return undefined;
  }
  // Serial 0 of the 1900 system is Excel's 1900-01-00, a Saturday.
  const parts =
    day === 0 && !date1904 ? { year: 1900, month: 1, day: 0 } : serialToParts(day, date1904);
  // From the serial, not the calendar date: Excel's 1900 calendar counts a 1900-02-29 that never existed,
  // so serial 1 is a Sunday and every day before March 1900 is a weekday earlier than the real one.
  // Serial 0 is a Saturday under 1900 and a Friday under 1904.
  const weekday = (day + (date1904 ? 5 : 6)) % 7;
  const hour = Math.floor(secondOfDay / 3600);
  const minute = Math.floor((secondOfDay % 3600) / 60);
  const second = secondOfDay % 60;
  const twelveHour = tokens.some(t => t.kind === "ampm");
  const shownHour = twelveHour ? hour % 12 || 12 : hour;
  const fractionOf = (places: number): string =>
    places === 0 ? "" : "." + subSecond.padEnd(places, "0").slice(0, places);

  let out = "";
  // An elapsed count after a year, month or day field counts only the whole hours of that day, in its own
  // unit: `dd [h]` shows 1.5 as `01 12`, `yyyy [mm]` shows 12:15 as 720 — where `[h] dd` shows `36 01`.
  let calendarShown = false;
  for (const token of tokens) {
    switch (token.kind) {
      case "date":
        calendarShown ||=
          token.part !== "hour" && token.part !== "minute" && token.part !== "second";
        out += datePart(
          token.part,
          token.width,
          parts.year,
          parts.month,
          parts.day,
          weekday,
          shownHour,
          minute,
          second
        );
        out += token.part === "second" ? fractionOf(token.fraction) : "";
        break;
      case "elapsed": {
        const seconds = calendarShown ? hour * 3600 : totalSeconds;
        const total =
          token.unit === "h"
            ? Math.floor(seconds / 3600)
            : token.unit === "m"
              ? Math.floor(seconds / 60)
              : seconds;
        out += String(total).padStart(token.width, "0");
        out += token.unit === "s" ? fractionOf(token.fraction) : "";
        break;
      }
      case "ampm": {
        const pm = hour >= 12;
        const text = token.short ? (pm ? "P" : "A") : pm ? "PM" : "AM";
        out += token.lower ? text.toLowerCase() : text;
        break;
      }
      default:
        out += plainText(token);
    }
  }
  return out;
}

function datePart(
  part: FormatDatePart,
  width: number,
  year: number,
  month: number,
  day: number,
  weekday: number,
  hour: number,
  minute: number,
  second: number
): string {
  switch (part) {
    case "year":
    case "buddhistYear": {
      const y = part === "year" ? year : year + 543;
      return width <= 2 ? pad2(((y % 100) + 100) % 100) : String(y).padStart(4, "0");
    }
    case "month":
      if (width === 1) {
        return String(month);
      }
      if (width === 2) {
        return pad2(month);
      }
      if (width === 3) {
        return MONTH_NAMES[month - 1].slice(0, 3);
      }
      return width === 5 ? MONTH_NAMES[month - 1][0] : MONTH_NAMES[month - 1];
    case "day":
      if (width === 1) {
        return String(day);
      }
      if (width === 2) {
        return pad2(day);
      }
      return width === 3 ? DAY_NAMES[weekday].slice(0, 3) : DAY_NAMES[weekday];
    case "hour":
      return width === 1 ? String(hour) : pad2(hour);
    case "minute":
      return width === 1 ? String(minute) : pad2(minute);
    default:
      return width === 1 ? String(second) : pad2(second);
  }
}
