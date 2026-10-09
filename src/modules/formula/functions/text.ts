/**
 * Text Functions — Native RuntimeValue implementations.
 */

import { argToNumber, checkError, excelWildcardToRegex } from "@formula/functions/_shared";
import type { FunctionRuntimeContext } from "@formula/runtime/function-context";
import { DEFAULT_FUNCTION_CONTEXT } from "@formula/runtime/function-context";
import type { RuntimeValue, ScalarValue, ErrorValue } from "@formula/runtime/values";
import {
  RVKind,
  ERRORS,
  isError,
  isArray,
  toNumberRV,
  toStringRV,
  toBooleanRV,
  topLeft,
  rvNumber,
  rvString,
  rvBoolean,
  rvArray
} from "@formula/runtime/values";
import { parseNumberFormat } from "@utils/number-format";
import { tryRenderNumberFormat } from "@utils/number-format-render";

// ============================================================================
// Local utility
// ============================================================================

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Type alias for a native text function. */
type NativeFn = (args: RuntimeValue[]) => RuntimeValue;

// ============================================================================
// CONCATENATE / CONCAT
// ============================================================================

export function fnCONCATENATE(args: RuntimeValue[]): RuntimeValue {
  const parts: string[] = [];
  for (const a of args) {
    if (isArray(a)) {
      for (const row of a.rows) {
        for (const cell of row) {
          const err = checkError(cell);
          if (err) {
            return err;
          }
          parts.push(toStringRV(cell));
        }
      }
    } else {
      const err = checkError(a);
      if (err) {
        return err;
      }
      parts.push(toStringRV(a));
    }
  }
  return rvString(parts.join(""));
}

// CONCAT has the same semantics as CONCATENATE
export const fnCONCAT: NativeFn = fnCONCATENATE;

// ============================================================================
// TEXTJOIN
// ============================================================================

export function fnTEXTJOIN(args: RuntimeValue[]): RuntimeValue {
  if (args.length < 3) {
    return ERRORS.VALUE;
  }
  const e0 = checkError(args[0]);
  if (e0) {
    return e0;
  }
  const delimiter = toStringRV(topLeft(args[0]));
  const ignoreEmptyRV = toBooleanRV(topLeft(args[1]));
  if (isError(ignoreEmptyRV)) {
    return ignoreEmptyRV;
  }
  const ignoreEmpty = ignoreEmptyRV.value;
  const parts: string[] = [];
  for (let i = 2; i < args.length; i++) {
    const a = args[i];
    if (isArray(a)) {
      for (const row of a.rows) {
        for (const cell of row) {
          const err = checkError(cell);
          if (err) {
            return err;
          }
          const s = toStringRV(cell);
          if (ignoreEmpty && s === "") {
            continue;
          }
          parts.push(s);
        }
      }
    } else {
      const err = checkError(a);
      if (err) {
        return err;
      }
      const s = toStringRV(a);
      if (ignoreEmpty && s === "") {
        continue;
      }
      parts.push(s);
    }
  }
  return rvString(parts.join(delimiter));
}

// ============================================================================
// LEFT / RIGHT / MID / LEN
// ============================================================================

export function fnLEFT(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  // Implicit intersection on the text arg (see MID for rationale).
  const text = toStringRV(topLeft(args[0]));
  let n: number;
  if (args.length > 1) {
    // Use `argToNumber` so array arguments get implicit-intersection to
    // their top-left cell before numeric coercion — otherwise
    // `LEFT("abc", A1:A2)` would land in `toNumberRV`'s array path and
    // incorrectly surface #VALUE! instead of using A1.
    const nRV = argToNumber(args[1]);
    if (isError(nRV)) {
      return nRV;
    }
    n = nRV.value;
  } else {
    n = 1;
  }
  // Excel rejects negative lengths outright. Without this guard,
  // `text.slice(0, -1)` would silently trim the last character.
  if (n < 0) {
    return ERRORS.VALUE;
  }
  return rvString(text.slice(0, Math.trunc(n)));
}

export function fnRIGHT(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0]));
  let n: number;
  if (args.length > 1) {
    // Implicit intersection via `argToNumber` — see LEFT for rationale.
    const nRV = argToNumber(args[1]);
    if (isError(nRV)) {
      return nRV;
    }
    n = nRV.value;
  } else {
    n = 1;
  }
  if (n < 0) {
    return ERRORS.VALUE;
  }
  const k = Math.trunc(n);
  if (k === 0) {
    return rvString("");
  }
  return rvString(text.slice(-k));
}

export function fnMID(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  // Implicit intersection on the text arg — without topLeft, passing
  // an array would route through `toStringRV`'s default branch and
  // silently return the empty string, making `MID(A1:A2, 1, 3)` look
  // like an empty cell instead of a 3-char prefix of the first cell.
  const text = toStringRV(topLeft(args[0]));
  const startNumRV = argToNumber(args[1]);
  if (isError(startNumRV)) {
    return startNumRV;
  }
  const startNum = Math.trunc(startNumRV.value);
  const numCharsRV = argToNumber(args[2]);
  if (isError(numCharsRV)) {
    return numCharsRV;
  }
  const numChars = Math.trunc(numCharsRV.value);
  // MID: start_num must be >= 1, num_chars must be >= 0.
  if (startNum < 1 || numChars < 0) {
    return ERRORS.VALUE;
  }
  return rvString(text.slice(startNum - 1, startNum - 1 + numChars));
}

export function fnLEN(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  // `toStringRV` doesn't dereference arrays — passing `A1:A2` would hit
  // its `default: ""` branch and silently return 0. Do an implicit
  // intersection via `topLeft` so `LEN(A1:A2)` behaves like Excel's
  // legacy implicit-intersection semantics (pick the first cell).
  return rvNumber(toStringRV(topLeft(args[0])).length);
}

// ============================================================================
// TRIM / LOWER / UPPER / PROPER
// ============================================================================

export function fnTRIM(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  // Implicit intersection: turn an array argument into its top-left
  // cell before stringifying, to match Excel's legacy behaviour.
  // Excel's TRIM only collapses plain ASCII space (U+0020), NOT tabs,
  // newlines, or non-breaking space (U+00A0).
  return rvString(
    toStringRV(topLeft(args[0]))
      .replace(/^ +| +$/g, "")
      .replace(/ +/g, " ")
  );
}

export function fnLOWER(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  return rvString(toStringRV(topLeft(args[0])).toLowerCase());
}

export function fnUPPER(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  return rvString(toStringRV(topLeft(args[0])).toUpperCase());
}

export function fnPROPER(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  return rvString(
    toStringRV(topLeft(args[0])).replace(
      /\p{L}+/gu,
      word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
    )
  );
}

// ============================================================================
// SUBSTITUTE / REPLACE
// ============================================================================

export function fnSUBSTITUTE(args: RuntimeValue[]): RuntimeValue {
  const err0 = checkError(args[0]);
  if (err0) {
    return err0;
  }
  const err1 = checkError(args[1]);
  if (err1) {
    return err1;
  }
  const err2 = checkError(args[2]);
  if (err2) {
    return err2;
  }
  const text = toStringRV(topLeft(args[0]));
  const oldText = toStringRV(topLeft(args[1]));
  const newText = toStringRV(topLeft(args[2]));
  // An empty old_text is a no-op in Excel. Without this guard we would
  // `"abc".split("").join(newText)` and insert newText between every
  // character, and the regex path would match empty strings infinitely.
  if (oldText === "") {
    return rvString(text);
  }
  if (args.length > 3) {
    const instanceNumRV = toNumberRV(topLeft(args[3]));
    if (isError(instanceNumRV)) {
      return instanceNumRV;
    }
    // Excel requires a positive integer; zero, negative, or non-numeric
    // values are #VALUE!. Previously we let the replace pass silently
    // no-op (since `count === 0` never matched), masking caller bugs.
    if (!Number.isFinite(instanceNumRV.value) || instanceNumRV.value < 1) {
      return ERRORS.VALUE;
    }
    const instanceNum = Math.trunc(instanceNumRV.value);
    let count = 0;
    return rvString(
      text.replace(new RegExp(escapeRegex(oldText), "g"), match => {
        count++;
        return count === instanceNum ? newText : match;
      })
    );
  }
  return rvString(text.split(oldText).join(newText));
}

export function fnREPLACE(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0]));
  // Implicit intersection on the numeric arguments — see LEFT.
  const startNumRV = argToNumber(args[1]);
  if (isError(startNumRV)) {
    return startNumRV;
  }
  const startNum = Math.trunc(startNumRV.value);
  const numCharsRV = argToNumber(args[2]);
  if (isError(numCharsRV)) {
    return numCharsRV;
  }
  const numChars = Math.trunc(numCharsRV.value);
  // REPLACE: start_num >= 1, num_chars >= 0. Without this check, negative
  // start_num becomes a slice with negative index and silently trims from
  // the right, which does not match Excel's #VALUE! result.
  if (startNum < 1 || numChars < 0) {
    return ERRORS.VALUE;
  }
  const e3 = checkError(args[3]);
  if (e3) {
    return e3;
  }
  const newText = toStringRV(topLeft(args[3]));
  return rvString(text.slice(0, startNum - 1) + newText + text.slice(startNum - 1 + numChars));
}

// ============================================================================
// FIND / SEARCH
// ============================================================================

export function fnFIND(args: RuntimeValue[]): RuntimeValue {
  const err0 = checkError(args[0]);
  if (err0) {
    return err0;
  }
  const err1 = checkError(args[1]);
  if (err1) {
    return err1;
  }
  // Implicit intersection on text args so arrays collapse to top-left.
  const findText = toStringRV(topLeft(args[0]));
  const withinText = toStringRV(topLeft(args[1]));
  let startNum: number;
  if (args.length > 2) {
    // Implicit intersection so an array supplied as start_num collapses
    // to its top-left cell — matches Excel and the other text family.
    const startNumRV = argToNumber(args[2]);
    if (isError(startNumRV)) {
      return startNumRV;
    }
    startNum = Math.trunc(startNumRV.value);
  } else {
    startNum = 1;
  }
  // Excel's FIND rejects start_num outside [1, length(withinText)].
  if (startNum < 1 || startNum > withinText.length + 1) {
    return ERRORS.VALUE;
  }
  const idx = withinText.indexOf(findText, startNum - 1);
  return idx === -1 ? ERRORS.VALUE : rvNumber(idx + 1);
}

export function fnSEARCH(args: RuntimeValue[]): RuntimeValue {
  const err0 = checkError(args[0]);
  if (err0) {
    return err0;
  }
  const err1 = checkError(args[1]);
  if (err1) {
    return err1;
  }
  // Implicit intersection on text args (see FIND for rationale).
  let findText = toStringRV(topLeft(args[0]));
  const withinText = toStringRV(topLeft(args[1]));
  let startNum: number;
  if (args.length > 2) {
    const startNumRV = argToNumber(args[2]);
    if (isError(startNumRV)) {
      return startNumRV;
    }
    startNum = Math.trunc(startNumRV.value);
  } else {
    startNum = 1;
  }
  if (startNum < 1 || startNum > withinText.length + 1) {
    return ERRORS.VALUE;
  }
  // Use the shared Excel-wildcard → regex converter so SEARCH, MATCH,
  // XLOOKUP, and SUMIF/COUNTIF agree on escape semantics (`~*`, `~?`, `~~`).
  const pattern = excelWildcardToRegex(findText);
  try {
    const re = new RegExp(pattern, "i");
    const sub = withinText.slice(startNum - 1);
    const match = re.exec(sub);
    return match ? rvNumber(match.index + startNum) : ERRORS.VALUE;
  } catch {
    // If regex is invalid, fall back to simple case-insensitive indexOf.
    findText = findText.toLowerCase();
    const idx = withinText.toLowerCase().indexOf(findText, startNum - 1);
    return idx === -1 ? ERRORS.VALUE : rvNumber(idx + 1);
  }
}

// ============================================================================
// REPT
// ============================================================================

export function fnREPT(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0]));
  const timesRV = toNumberRV(topLeft(args[1]));
  if (isError(timesRV)) {
    return timesRV;
  }
  const times = Math.floor(timesRV.value);
  if (times < 0) {
    return ERRORS.VALUE;
  }
  // Excel caps the result at 32767 characters; we additionally bail out
  // early on huge products so the engine can't be DoS'd into allocating
  // a multi-gigabyte string. (R6-P1-4)
  const total = text.length * times;
  if (total > 32767) {
    return ERRORS.VALUE;
  }
  return rvString(text.repeat(times));
}

// ============================================================================
// TEXT (complex number/date formatting)
// ============================================================================

export function fnTEXT(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const rawVal = topLeft(args[0]);
  if (isError(rawVal)) {
    return rawVal;
  }
  const e1 = checkError(args[1]);
  if (e1) {
    return e1;
  }
  const fmt = toStringRV(topLeft(args[1]));

  // The format is read and rendered by the same code that produces a cell's display text, so
  // `TEXT(x, f)` and a cell holding `x` formatted with `f` cannot disagree. Where the cell would show
  // hashes — a date past 9999, a negative time — `TEXT` is `#VALUE!`, as in Excel.
  const render = (value: number | string): RuntimeValue => {
    const text = tryRenderNumberFormat(fmt, value, { date1904: context.date1904 });
    return text === undefined ? ERRORS.VALUE : rvString(text);
  };

  // A logical is shown as itself under any format: `TEXT(TRUE, "0")` is `TRUE`.
  if (rawVal.kind === RVKind.Boolean) {
    return rvString(toStringRV(rawVal));
  }
  if (rawVal.kind === RVKind.String) {
    // A four-section format formats text with its fourth section. Otherwise text that reads as a number is that
    // number, and any other text goes through the text section if there is one — or comes back unchanged, which
    // is what Excel does with `TEXT("abc", "0.00")` rather than failing.
    if (parseNumberFormat(fmt).sections.length < 4) {
      const num = toNumberRV(rawVal);
      if (!isError(num)) {
        return render(num.value);
      }
    }
    return render(toStringRV(rawVal));
  }

  const valRV = toNumberRV(rawVal);
  if (isError(valRV)) {
    return valRV;
  }
  return render(valRV.value);
}

// ============================================================================
// VALUE / EXACT
// ============================================================================

export function fnVALUE(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  // Delegate to the central numeric-string parser. It rejects empty /
  // whitespace-only / Infinity / NaN / hex forms the way Excel does, which
  // a naive `Number(s)` would accept silently.
  return toNumberRV(topLeft(args[0]));
}

export function fnEXACT(args: RuntimeValue[]): RuntimeValue {
  const err0 = checkError(args[0]);
  if (err0) {
    return err0;
  }
  const err1 = checkError(args[1]);
  if (err1) {
    return err1;
  }
  return rvBoolean(toStringRV(topLeft(args[0])) === toStringRV(topLeft(args[1])));
}

// ============================================================================
// Additional Text Functions
// ============================================================================

export function fnCODE(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0]));
  return text.length > 0 ? rvNumber(text.charCodeAt(0)) : ERRORS.VALUE;
}

export function fnCHAR(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const nRV = toNumberRV(topLeft(args[0]));
  if (isError(nRV)) {
    return nRV;
  }
  // Excel's CHAR accepts integers in [1, 255] only; outside the ANSI range
  // it returns #VALUE!. We also truncate fractional inputs toward zero to
  // match Excel's coercion semantics.
  const code = Math.trunc(nRV.value);
  if (code < 1 || code > 255) {
    return ERRORS.VALUE;
  }
  return rvString(String.fromCharCode(code));
}

export function fnCLEAN(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0]));
  // Remove non-printable ASCII control characters (0x00-0x1F)
  let result = "";
  for (let i = 0; i < text.length; i++) {
    if (text.charCodeAt(i) >= 32) {
      result += text[i];
    }
  }
  return rvString(result);
}

export function fnT(args: RuntimeValue[]): RuntimeValue {
  const v = topLeft(args[0]);
  if (isError(v)) {
    return v;
  }
  return v.kind === RVKind.String ? v : rvString("");
}

// ============================================================================
// More Text Functions
// ============================================================================

export function fnUNICHAR(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const nRV = toNumberRV(topLeft(args[0]));
  if (isError(nRV)) {
    return nRV;
  }
  const code = Math.floor(nRV.value);
  if (code < 1) {
    return ERRORS.VALUE;
  }
  try {
    return rvString(String.fromCodePoint(code));
  } catch {
    return ERRORS.VALUE;
  }
}

export function fnUNICODE(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0]));
  if (text.length === 0) {
    return ERRORS.VALUE;
  }
  const cp = text.codePointAt(0);
  return cp !== undefined ? rvNumber(cp) : ERRORS.VALUE;
}

export function fnBAHTTEXT(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  return rvString(toStringRV(topLeft(args[0])));
}

export function fnDOLLAR(args: RuntimeValue[]): RuntimeValue {
  const numRV = toNumberRV(topLeft(args[0]));
  if (isError(numRV)) {
    return numRV;
  }
  const num = numRV.value;
  let decimals: number;
  if (args.length > 1) {
    const decRV = toNumberRV(topLeft(args[1]));
    if (isError(decRV)) {
      return decRV;
    }
    decimals = decRV.value;
  } else {
    decimals = 2;
  }
  const d = Math.floor(decimals);
  let rounded: number;
  if (d < 0) {
    const factor = Math.pow(10, -d);
    rounded = Math.round(Math.abs(num) / factor) * factor;
  } else {
    rounded = Math.abs(num);
  }
  const formatted = rounded.toFixed(Math.max(0, d));
  const parts = formatted.split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const result = parts.join(".");
  return rvString(num < 0 ? `($${result})` : `$${result}`);
}

export function fnFIXED(args: RuntimeValue[]): RuntimeValue {
  const numRV = toNumberRV(topLeft(args[0]));
  if (isError(numRV)) {
    return numRV;
  }
  const num = numRV.value;
  let decimals: number;
  if (args.length > 1) {
    const decRV = toNumberRV(topLeft(args[1]));
    if (isError(decRV)) {
      return decRV;
    }
    decimals = decRV.value;
  } else {
    decimals = 2;
  }
  let noCommas: boolean;
  if (args.length > 2) {
    const ncRV = toBooleanRV(topLeft(args[2]));
    if (isError(ncRV)) {
      return ncRV;
    }
    noCommas = ncRV.value;
  } else {
    noCommas = false;
  }
  const d = Math.floor(decimals);
  let rounded: number;
  if (d < 0) {
    const factor = Math.pow(10, -d);
    rounded = Math.round(num / factor) * factor;
  } else {
    rounded = num;
  }
  let result = rounded.toFixed(Math.max(0, d));
  if (!noCommas) {
    const parts = result.split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    result = parts.join(".");
  }
  return rvString(result);
}

export function fnASC(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0]));
  return rvString(
    text.replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
  );
}

export function fnDBCS(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0]));
  return rvString(text.replace(/[!-~]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0xfee0)));
}

export function fnJIS(args: RuntimeValue[]): RuntimeValue {
  return fnDBCS(args);
}

export function fnPHONETIC(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  // Implicit intersection: array → top-left cell, matching the rest of
  // the text-function family.
  return rvString(toStringRV(topLeft(args[0])));
}

export function fnNUMBERVALUE(args: RuntimeValue[]): RuntimeValue {
  const e0 = checkError(args[0]);
  if (e0) {
    return e0;
  }
  // Implicit intersection on every text arg.
  let text = toStringRV(topLeft(args[0]));
  let decSep = ".";
  if (args.length > 1) {
    const e1 = checkError(args[1]);
    if (e1) {
      return e1;
    }
    decSep = toStringRV(topLeft(args[1]));
  }
  let grpSep = ",";
  if (args.length > 2) {
    const e2 = checkError(args[2]);
    if (e2) {
      return e2;
    }
    grpSep = toStringRV(topLeft(args[2]));
  }
  text = text.split(grpSep).join("");
  if (decSep !== ".") {
    text = text.replace(decSep, ".");
  }
  // Handle percentage. Excel divides by 100 for EACH trailing `%`, so
  // `NUMBERVALUE("50%%")` = 50 / 10000 = 0.005. Previously we only
  // recognised the first `%` and treated `50%%` as the literal string
  // "50%" → NaN → #VALUE!.
  let pctCount = 0;
  while (text.endsWith("%")) {
    pctCount++;
    text = text.slice(0, -1);
  }
  // `Number("")` is 0, not NaN — reject empty / whitespace-only inputs
  // so `NUMBERVALUE("")` does not silently produce 0 (R6-P1-6).
  if (text.trim() === "") {
    return ERRORS.VALUE;
  }
  const n = Number(text);
  if (isNaN(n)) {
    return ERRORS.VALUE;
  }
  return rvNumber(pctCount > 0 ? n / Math.pow(100, pctCount) : n);
}

// ============================================================================
// Excel 365 Text Functions: TEXTBEFORE, TEXTAFTER, TEXTSPLIT
// ============================================================================

/**
 * Parse the common [instance_num, match_mode, match_end, if_not_found]
 * tail used by TEXTBEFORE / TEXTAFTER. Returns the numeric values (with
 * defaults filled in) or an error if any argument is malformed.
 *
 * - match_mode: 0 = case-sensitive (default), 1 = case-insensitive.
 * - match_end:  0 = don't treat string edge as delimiter (default),
 *               1 = treat string edge as a virtual delimiter so that
 *                   TEXTAFTER with a missing delimiter returns "".
 */
function parseTextBeforeAfterTail(
  args: RuntimeValue[]
):
  | { inst: number; matchMode: 0 | 1; matchEnd: 0 | 1; ifNotFound: RuntimeValue | null }
  | ErrorValue {
  let inst = 1;
  if (args.length > 2) {
    const instRV = toNumberRV(topLeft(args[2]));
    if (isError(instRV)) {
      return instRV;
    }
    inst = Math.trunc(instRV.value);
  }
  let matchMode: 0 | 1 = 0;
  if (args.length > 3) {
    const mmRV = toNumberRV(topLeft(args[3]));
    if (isError(mmRV)) {
      return mmRV;
    }
    const mm = Math.trunc(mmRV.value);
    if (mm !== 0 && mm !== 1) {
      return ERRORS.VALUE;
    }
    matchMode = mm;
  }
  let matchEnd: 0 | 1 = 0;
  if (args.length > 4) {
    const meRV = toNumberRV(topLeft(args[4]));
    if (isError(meRV)) {
      return meRV;
    }
    const me = Math.trunc(meRV.value);
    if (me !== 0 && me !== 1) {
      return ERRORS.VALUE;
    }
    matchEnd = me;
  }
  const ifNotFound = args.length > 5 ? args[5] : null;
  return { inst, matchMode, matchEnd, ifNotFound };
}

export function fnTEXTBEFORE(args: RuntimeValue[]): RuntimeValue {
  const e0 = checkError(args[0]);
  if (e0) {
    return e0;
  }
  const e1 = checkError(args[1]);
  if (e1) {
    return e1;
  }
  const text = toStringRV(topLeft(args[0]));
  const delimiter = toStringRV(topLeft(args[1]));
  const tail = parseTextBeforeAfterTail(args);
  if ("kind" in tail && tail.kind === RVKind.Error) {
    return tail;
  }
  const { inst, matchMode, matchEnd, ifNotFound } = tail as Exclude<typeof tail, ErrorValue>;
  if (inst === 0) {
    return ERRORS.VALUE;
  }
  if (delimiter === "") {
    return rvString(inst > 0 ? "" : text);
  }
  // For case-insensitive matching we search within the lower-cased
  // haystack but slice against the original so the returned prefix/
  // suffix preserves the source text's case.
  const haystack = matchMode === 1 ? text.toLowerCase() : text;
  const needle = matchMode === 1 ? delimiter.toLowerCase() : delimiter;
  const notFound = (): RuntimeValue => {
    if (matchEnd === 1 && inst === 1) {
      // Treat the string end as a virtual delimiter: everything is "before".
      return rvString(text);
    }
    return ifNotFound !== null ? ifNotFound : ERRORS.NA;
  };
  if (inst > 0) {
    let pos = -1;
    for (let i = 0; i < inst; i++) {
      pos = haystack.indexOf(needle, pos + 1);
      if (pos === -1) {
        return notFound();
      }
    }
    return rvString(text.slice(0, pos));
  }
  // Negative: search from end
  let pos = haystack.length;
  for (let i = 0; i < -inst; i++) {
    pos = haystack.lastIndexOf(needle, pos - 1);
    if (pos === -1) {
      return notFound();
    }
  }
  return rvString(text.slice(0, pos));
}

export function fnTEXTAFTER(args: RuntimeValue[]): RuntimeValue {
  const e0 = checkError(args[0]);
  if (e0) {
    return e0;
  }
  const e1 = checkError(args[1]);
  if (e1) {
    return e1;
  }
  const text = toStringRV(topLeft(args[0]));
  const delimiter = toStringRV(topLeft(args[1]));
  const tail = parseTextBeforeAfterTail(args);
  if ("kind" in tail && tail.kind === RVKind.Error) {
    return tail;
  }
  const { inst, matchMode, matchEnd, ifNotFound } = tail as Exclude<typeof tail, ErrorValue>;
  if (inst === 0) {
    return ERRORS.VALUE;
  }
  if (delimiter === "") {
    return rvString(inst > 0 ? text : "");
  }
  const haystack = matchMode === 1 ? text.toLowerCase() : text;
  const needle = matchMode === 1 ? delimiter.toLowerCase() : delimiter;
  const notFound = (): RuntimeValue => {
    if (matchEnd === 1 && inst === 1) {
      // String end is a virtual delimiter → everything after it is "".
      return rvString("");
    }
    return ifNotFound !== null ? ifNotFound : ERRORS.NA;
  };
  if (inst > 0) {
    let pos = -1;
    for (let i = 0; i < inst; i++) {
      pos = haystack.indexOf(needle, pos + 1);
      if (pos === -1) {
        return notFound();
      }
    }
    return rvString(text.slice(pos + delimiter.length));
  }
  let pos = haystack.length;
  for (let i = 0; i < -inst; i++) {
    pos = haystack.lastIndexOf(needle, pos - 1);
    if (pos === -1) {
      return notFound();
    }
  }
  return rvString(text.slice(pos + delimiter.length));
}

export function fnTEXTSPLIT(args: RuntimeValue[]): RuntimeValue {
  const e0 = checkError(args[0]);
  if (e0) {
    return e0;
  }
  const text = toStringRV(topLeft(args[0]));
  let colDelimiter = "";
  if (args.length > 1) {
    const e1 = checkError(args[1]);
    if (e1) {
      return e1;
    }
    colDelimiter = toStringRV(topLeft(args[1]));
  }
  const rowDelimiter =
    args.length > 2 && args[2].kind !== RVKind.Blank ? toStringRV(topLeft(args[2])) : "";

  // `ignore_empty` (4th arg, default FALSE) — when TRUE, suppress empty
  // fragments produced by consecutive delimiters.
  let ignoreEmpty = false;
  if (args.length > 3 && args[3].kind !== RVKind.Blank) {
    const ieRV = toBooleanRV(topLeft(args[3]));
    if (isError(ieRV)) {
      return ieRV;
    }
    ignoreEmpty = ieRV.value;
  }

  // `match_mode` (5th arg, default 0 = case-sensitive). When 1 the
  // delimiter match is case-insensitive; we implement that by lowercasing
  // both the haystack and the delimiter(s) before splitting, which is
  // consistent with Excel's specification for TEXTSPLIT.
  let matchMode = 0;
  if (args.length > 4 && args[4].kind !== RVKind.Blank) {
    const mmRV = toNumberRV(topLeft(args[4]));
    if (isError(mmRV)) {
      return mmRV;
    }
    matchMode = Math.trunc(mmRV.value);
    if (matchMode !== 0 && matchMode !== 1) {
      return ERRORS.VALUE;
    }
  }

  // `pad_with` (6th arg, default #N/A) — value used to fill shorter rows
  // when the split produces a ragged 2D shape. Explicit error arguments
  // (e.g. `TEXTSPLIT(…, #VALUE!)`) propagate into the pad cells verbatim,
  // matching Excel.
  const pad: ScalarValue = args.length > 5 ? topLeft(args[5]) : ERRORS.NA;

  const splitString = (s: string, delim: string): string[] => {
    if (!delim) {
      return [s];
    }
    if (matchMode === 1) {
      // Case-insensitive split — find positions by scanning the lowercased
      // haystack but slice the original so case is preserved in output.
      const haystack = s.toLowerCase();
      const needle = delim.toLowerCase();
      const parts: string[] = [];
      let last = 0;
      let i = 0;
      while (i <= haystack.length - needle.length) {
        if (haystack.slice(i, i + needle.length) === needle) {
          parts.push(s.slice(last, i));
          i += needle.length;
          last = i;
        } else {
          i++;
        }
      }
      parts.push(s.slice(last));
      return parts;
    }
    return s.split(delim);
  };

  let rows: string[];
  if (rowDelimiter) {
    rows = splitString(text, rowDelimiter);
  } else {
    rows = [text];
  }

  // Split each row into columns, applying ignore_empty per row after the
  // split. When ignore_empty is TRUE at the row level we also drop rows
  // that were themselves empty (i.e. empty string from consecutive row
  // delimiters).
  const matrix: ScalarValue[][] = [];
  let maxWidth = 0;
  for (const row of rows) {
    if (ignoreEmpty && row === "") {
      continue;
    }
    let parts: string[];
    if (colDelimiter) {
      parts = splitString(row, colDelimiter);
      if (ignoreEmpty) {
        parts = parts.filter(p => p !== "");
      }
    } else {
      parts = [row];
    }
    if (parts.length === 0) {
      // All fragments were empty and ignore_empty discarded them; keep a
      // pad row so the result is still a well-formed rectangle.
      parts = [""];
    }
    matrix.push(parts.map(p => rvString(p)));
    if (parts.length > maxWidth) {
      maxWidth = parts.length;
    }
  }

  if (matrix.length === 0) {
    // ignore_empty consumed everything → return a single pad cell so the
    // array is still a valid 1×1 spill (matches Excel).
    return rvArray([[pad]]);
  }

  // Pad ragged rows out to the maximum width with `pad_with`.
  const result: ScalarValue[][] = [];
  for (const row of matrix) {
    if (row.length < maxWidth) {
      const padded: ScalarValue[] = row.slice();
      while (padded.length < maxWidth) {
        padded.push(pad);
      }
      result.push(padded);
    } else {
      result.push(row);
    }
  }
  return rvArray(result);
}

// ============================================================================
// REGEX family (Excel 365, 2024)
// ============================================================================

/**
 * Convert an Excel REGEX pattern to a JavaScript RegExp. Excel's regex
 * dialect is close to PCRE; JavaScript's RegExp is close enough for the
 * vast majority of practical patterns, but a few constructs (named
 * captures, look-behind, some Unicode classes) behave slightly
 * differently. We pass patterns through as-is and let JavaScript's
 * parser surface #VALUE! on the rare incompatibility.
 */
function compileExcelRegex(
  pattern: string,
  caseSensitive: boolean,
  global: boolean
): RegExp | null {
  try {
    let flags = "u";
    if (!caseSensitive) {
      flags += "i";
    }
    if (global) {
      flags += "g";
    }
    return new RegExp(pattern, flags);
  } catch {
    return null;
  }
}

/**
 * Resolve the optional `case_sensitivity` argument used by every REGEX
 * function. `0`/FALSE/omitted → case-insensitive (Excel default),
 * any other value → case-sensitive. Errors propagate.
 */
function resolveCaseSensitivity(
  arg: RuntimeValue | undefined
): { caseSensitive: boolean } | ErrorValue {
  if (arg === undefined) {
    return { caseSensitive: false };
  }
  // Accept boolean or number; anything else coerced via toBooleanRV.
  const b = toBooleanRV(arg);
  if (isError(b)) {
    return b;
  }
  return { caseSensitive: b.value };
}

/**
 * REGEXTEST(text, pattern, [case_sensitivity]) — returns TRUE iff the
 * regex matches any substring of `text`.
 */
export function fnREGEXTEST(args: RuntimeValue[]): RuntimeValue {
  const textV = toStringRV(topLeft(args[0]));
  const patternV = toStringRV(topLeft(args[1]));
  const cs = resolveCaseSensitivity(args[2]);
  if ("kind" in cs) {
    return cs; // error
  }
  const errCheck = checkError(args[0]) ?? checkError(args[1]);
  if (errCheck) {
    return errCheck;
  }
  const re = compileExcelRegex(patternV, cs.caseSensitive, false);
  if (!re) {
    return ERRORS.VALUE;
  }
  return rvBoolean(re.test(textV));
}

/**
 * REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity]) —
 *   return_mode = 0 (default) → first match as a string
 *   return_mode = 1 → all matches as a 1-column array
 *   return_mode = 2 → capture groups of the first match as a 1-row array
 */
export function fnREGEXEXTRACT(args: RuntimeValue[]): RuntimeValue {
  const textV = toStringRV(topLeft(args[0]));
  const patternV = toStringRV(topLeft(args[1]));
  const errCheck = checkError(args[0]) ?? checkError(args[1]);
  if (errCheck) {
    return errCheck;
  }
  const modeV = args.length > 2 ? toNumberRV(topLeft(args[2])) : rvNumber(0);
  if (isError(modeV)) {
    return modeV;
  }
  const mode = Math.trunc(modeV.value);
  if (mode !== 0 && mode !== 1 && mode !== 2) {
    return ERRORS.VALUE;
  }
  const cs = resolveCaseSensitivity(args[3]);
  if ("kind" in cs) {
    return cs;
  }
  const needGlobal = mode === 1;
  const re = compileExcelRegex(patternV, cs.caseSensitive, needGlobal);
  if (!re) {
    return ERRORS.VALUE;
  }
  if (mode === 0) {
    const m = re.exec(textV);
    if (!m) {
      return ERRORS.NA;
    }
    return rvString(m[0]);
  }
  if (mode === 1) {
    const matches: string[] = [];
    let m: RegExpExecArray | null;
    // eslint-disable-next-line no-cond-assign
    while ((m = re.exec(textV)) !== null) {
      matches.push(m[0]);
      // Guard against zero-length matches causing an infinite loop.
      if (m.index === re.lastIndex) {
        re.lastIndex++;
      }
    }
    if (matches.length === 0) {
      return ERRORS.NA;
    }
    return rvArray(matches.map(s => [rvString(s)]));
  }
  // mode === 2 — capture groups of first match as a row array.
  const m = re.exec(textV);
  if (!m) {
    return ERRORS.NA;
  }
  // Exclude the full-match element (index 0) — only capture groups.
  if (m.length <= 1) {
    // No capture groups defined in the pattern — return the full match.
    return rvArray([[rvString(m[0])]]);
  }
  const row: ScalarValue[] = [];
  for (let i = 1; i < m.length; i++) {
    row.push(rvString(m[i] ?? ""));
  }
  return rvArray([row]);
}

/**
 * REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity]) —
 *   occurrence = 0 (default) → replace all
 *   occurrence = n (positive) → replace only the n-th match
 *   occurrence = n (negative) → replace only the n-th-last match
 */
export function fnREGEXREPLACE(args: RuntimeValue[]): RuntimeValue {
  const textV = toStringRV(topLeft(args[0]));
  const patternV = toStringRV(topLeft(args[1]));
  const replacementV = toStringRV(topLeft(args[2]));
  const errCheck = checkError(args[0]) ?? checkError(args[1]) ?? checkError(args[2]);
  if (errCheck) {
    return errCheck;
  }
  const occurrenceV = args.length > 3 ? toNumberRV(topLeft(args[3])) : rvNumber(0);
  if (isError(occurrenceV)) {
    return occurrenceV;
  }
  const occurrence = Math.trunc(occurrenceV.value);
  const cs = resolveCaseSensitivity(args[4]);
  if ("kind" in cs) {
    return cs;
  }
  // Always compile with the global flag — we need to enumerate matches
  // to apply the occurrence filter; `String.replace` without `/g` would
  // only see the first match and we wouldn't be able to address later
  // hits for `occurrence > 1`.
  const re = compileExcelRegex(patternV, cs.caseSensitive, true);
  if (!re) {
    return ERRORS.VALUE;
  }

  if (occurrence === 0) {
    // Replace all.
    return rvString(textV.replace(re, replacementV));
  }

  // Collect every match's range so we can address them by index.
  const ranges: Array<{ start: number; end: number }> = [];
  let m: RegExpExecArray | null;
  // eslint-disable-next-line no-cond-assign
  while ((m = re.exec(textV)) !== null) {
    ranges.push({ start: m.index, end: m.index + m[0].length });
    if (m.index === re.lastIndex) {
      re.lastIndex++;
    }
  }
  if (ranges.length === 0) {
    return rvString(textV); // no match → unchanged (Excel behavior)
  }
  // Negative index counts from the end; -1 is the last match.
  const idx = occurrence > 0 ? occurrence - 1 : ranges.length + occurrence;
  if (idx < 0 || idx >= ranges.length) {
    // Out-of-range occurrence → unchanged (Excel behavior).
    return rvString(textV);
  }
  const { start, end } = ranges[idx];
  return rvString(textV.slice(0, start) + replacementV + textV.slice(end));
}

// ============================================================================
// VALUETOTEXT / ARRAYTOTEXT (Excel 365)
// ============================================================================

/**
 * Format a single scalar for VALUETOTEXT / ARRAYTOTEXT.
 *
 * Format 0 (concise, default):
 *   - Number → plain number string
 *   - String → the string itself (no quotes)
 *   - Boolean → "TRUE" / "FALSE"
 *   - Error → error text (e.g. "#N/A")
 *   - Blank → ""
 *
 * Format 1 (strict):
 *   - String → wrapped in double quotes with `""` escapes
 *   - Everything else → same as format 0
 */
function scalarToText(v: ScalarValue, strict: boolean): string {
  switch (v.kind) {
    case RVKind.Number:
      return String(v.value);
    case RVKind.String:
      if (strict) {
        return `"${v.value.replace(/"/g, '""')}"`;
      }
      return v.value;
    case RVKind.Boolean:
      return v.value ? "TRUE" : "FALSE";
    case RVKind.Error:
      return v.code;
    case RVKind.Blank:
      return "";
  }
}

/**
 * VALUETOTEXT(value, [format]) — format a scalar or 1×1 array as text.
 * For multi-cell arrays, this applies implicit intersection at the
 * evaluator layer — so by the time we see args[0] it is already scalar.
 */
export function fnVALUETOTEXT(args: RuntimeValue[]): RuntimeValue {
  const formatV = args.length > 1 ? toNumberRV(topLeft(args[1])) : rvNumber(0);
  if (isError(formatV)) {
    return formatV;
  }
  const fmt = Math.trunc(formatV.value);
  if (fmt !== 0 && fmt !== 1) {
    return ERRORS.VALUE;
  }
  const strict = fmt === 1;
  return rvString(scalarToText(topLeft(args[0]), strict));
}

/**
 * ARRAYTOTEXT(array, [format]) — flatten an array to a delimited text
 * representation.
 *
 * Format 0 (concise, default): row-major join with ", ".
 * Format 1 (strict): wraps output in `{…}`, rows separated by `;`,
 *   cells by `,`; strings inside quoted.
 */
export function fnARRAYTOTEXT(args: RuntimeValue[]): RuntimeValue {
  const formatV = args.length > 1 ? toNumberRV(topLeft(args[1])) : rvNumber(0);
  if (isError(formatV)) {
    return formatV;
  }
  const fmt = Math.trunc(formatV.value);
  if (fmt !== 0 && fmt !== 1) {
    return ERRORS.VALUE;
  }
  const strict = fmt === 1;
  const arg = args[0];
  if (arg.kind !== RVKind.Array) {
    return rvString(scalarToText(topLeft(arg), strict));
  }
  if (!strict) {
    // Concise: flatten row-major, join with ", ".
    const parts: string[] = [];
    for (const row of arg.rows) {
      for (const cell of row) {
        parts.push(scalarToText(cell, false));
      }
    }
    return rvString(parts.join(", "));
  }
  // Strict: `{row1;row2;...}` with rows as `a,b,c` and strings quoted.
  const rowStrs: string[] = [];
  for (const row of arg.rows) {
    const cellStrs: string[] = [];
    for (const cell of row) {
      cellStrs.push(scalarToText(cell, true));
    }
    rowStrs.push(cellStrs.join(","));
  }
  return rvString(`{${rowStrs.join(";")}}`);
}

// ============================================================================
// ENCODEURL
// ============================================================================

/**
 * ENCODEURL(text) — percent-encode a string for URL use.
 *
 * Excel follows RFC 3986's "unreserved" character rule: A-Z, a-z, 0-9,
 * and `- _ . ~` are kept verbatim; everything else is encoded as
 * `%HH` using the UTF-8 byte sequence. This is exactly what JavaScript's
 * `encodeURIComponent` does, so we delegate to it.
 */
export function fnENCODEURL(args: RuntimeValue[]): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const s = toStringRV(topLeft(args[0]));
  return rvString(encodeURIComponent(s));
}
