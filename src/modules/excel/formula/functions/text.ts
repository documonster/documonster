/**
 * Text Functions — Native RuntimeValue implementations.
 */

import { excelToDate } from "@utils/utils.base";

import type { RuntimeValue, ScalarValue } from "../runtime/values";
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
} from "../runtime/values";

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

export const fnCONCATENATE: NativeFn = args => {
  const parts: string[] = [];
  for (const a of args) {
    if (isArray(a)) {
      for (const row of a.rows) {
        for (const cell of row) {
          parts.push(toStringRV(cell));
        }
      }
    } else {
      parts.push(toStringRV(a));
    }
  }
  return rvString(parts.join(""));
};

// CONCAT has the same semantics as CONCATENATE
export const fnCONCAT: NativeFn = fnCONCATENATE;

// ============================================================================
// TEXTJOIN
// ============================================================================

export const fnTEXTJOIN: NativeFn = args => {
  if (args.length < 3) {
    return ERRORS.VALUE;
  }
  const delimiter = toStringRV(args[0]);
  const ignoreEmptyRV = toBooleanRV(args[1]);
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
          const s = toStringRV(cell);
          if (ignoreEmpty && s === "") {
            continue;
          }
          parts.push(s);
        }
      }
    } else {
      const s = toStringRV(a);
      if (ignoreEmpty && s === "") {
        continue;
      }
      parts.push(s);
    }
  }
  return rvString(parts.join(delimiter));
};

// ============================================================================
// LEFT / RIGHT / MID / LEN
// ============================================================================

export const fnLEFT: NativeFn = args => {
  const text = toStringRV(args[0]);
  let n: number;
  if (args.length > 1) {
    const nRV = toNumberRV(args[1]);
    if (isError(nRV)) {
      return nRV;
    }
    n = nRV.value;
  } else {
    n = 1;
  }
  return rvString(text.slice(0, n));
};

export const fnRIGHT: NativeFn = args => {
  const text = toStringRV(args[0]);
  let n: number;
  if (args.length > 1) {
    const nRV = toNumberRV(args[1]);
    if (isError(nRV)) {
      return nRV;
    }
    n = nRV.value;
  } else {
    n = 1;
  }
  if (n <= 0) {
    return rvString("");
  }
  return rvString(text.slice(-n));
};

export const fnMID: NativeFn = args => {
  const text = toStringRV(args[0]);
  const startNumRV = toNumberRV(args[1]);
  if (isError(startNumRV)) {
    return startNumRV;
  }
  const startNum = startNumRV.value;
  const numCharsRV = toNumberRV(args[2]);
  if (isError(numCharsRV)) {
    return numCharsRV;
  }
  const numChars = numCharsRV.value;
  return rvString(text.slice(startNum - 1, startNum - 1 + numChars));
};

export const fnLEN: NativeFn = args => {
  return rvNumber(toStringRV(args[0]).length);
};

// ============================================================================
// TRIM / LOWER / UPPER / PROPER
// ============================================================================

export const fnTRIM: NativeFn = args => {
  return rvString(toStringRV(args[0]).trim().replace(/\s+/g, " "));
};

export const fnLOWER: NativeFn = args => {
  return rvString(toStringRV(args[0]).toLowerCase());
};

export const fnUPPER: NativeFn = args => {
  return rvString(toStringRV(args[0]).toUpperCase());
};

export const fnPROPER: NativeFn = args => {
  return rvString(
    toStringRV(args[0]).replace(/\w\S*/g, t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase())
  );
};

// ============================================================================
// SUBSTITUTE / REPLACE
// ============================================================================

export const fnSUBSTITUTE: NativeFn = args => {
  const text = toStringRV(args[0]);
  const oldText = toStringRV(args[1]);
  const newText = toStringRV(args[2]);
  if (args.length > 3) {
    const instanceNumRV = toNumberRV(args[3]);
    if (isError(instanceNumRV)) {
      return instanceNumRV;
    }
    const instanceNum = instanceNumRV.value;
    let count = 0;
    return rvString(
      text.replace(new RegExp(escapeRegex(oldText), "g"), match => {
        count++;
        return count === instanceNum ? newText : match;
      })
    );
  }
  return rvString(text.split(oldText).join(newText));
};

export const fnREPLACE: NativeFn = args => {
  const text = toStringRV(args[0]);
  const startNumRV = toNumberRV(args[1]);
  if (isError(startNumRV)) {
    return startNumRV;
  }
  const startNum = startNumRV.value;
  const numCharsRV = toNumberRV(args[2]);
  if (isError(numCharsRV)) {
    return numCharsRV;
  }
  const numChars = numCharsRV.value;
  const newText = toStringRV(args[3]);
  return rvString(text.slice(0, startNum - 1) + newText + text.slice(startNum - 1 + numChars));
};

// ============================================================================
// FIND / SEARCH
// ============================================================================

export const fnFIND: NativeFn = args => {
  const findText = toStringRV(args[0]);
  const withinText = toStringRV(args[1]);
  let startNum: number;
  if (args.length > 2) {
    const startNumRV = toNumberRV(args[2]);
    if (isError(startNumRV)) {
      return startNumRV;
    }
    startNum = startNumRV.value;
  } else {
    startNum = 1;
  }
  const idx = withinText.indexOf(findText, startNum - 1);
  return idx === -1 ? ERRORS.VALUE : rvNumber(idx + 1);
};

export const fnSEARCH: NativeFn = args => {
  let findText = toStringRV(args[0]);
  const withinText = toStringRV(args[1]);
  let startNum: number;
  if (args.length > 2) {
    const startNumRV = toNumberRV(args[2]);
    if (isError(startNumRV)) {
      return startNumRV;
    }
    startNum = startNumRV.value;
  } else {
    startNum = 1;
  }
  // Convert Excel wildcards to regex: ? → ., * → .*, ~? → \?, ~* → \*
  // Use Unicode PUA char U+E000 as placeholder to avoid control char warnings
  const PUA = "\uE000";
  const pattern = findText
    .replace(/~/g, PUA)
    .replace(new RegExp(PUA + "\\?", "g"), "\\?")
    .replace(new RegExp(PUA + "\\*", "g"), "\\*")
    .replace(/[.*+^${}()|[\]\\]/g, m => (m === "." || m === "*" ? m : "\\" + m))
    .replace(/\?/g, ".")
    .replace(/\*/g, ".*")
    .replace(new RegExp(PUA, "g"), "");
  try {
    const re = new RegExp(pattern, "i");
    const sub = withinText.slice(startNum - 1);
    const match = re.exec(sub);
    return match ? rvNumber(match.index + startNum) : ERRORS.VALUE;
  } catch {
    // If regex is invalid, fall back to simple indexOf
    findText = findText.toLowerCase();
    const idx = withinText.toLowerCase().indexOf(findText, startNum - 1);
    return idx === -1 ? ERRORS.VALUE : rvNumber(idx + 1);
  }
};

// ============================================================================
// REPT
// ============================================================================

export const fnREPT: NativeFn = args => {
  const text = toStringRV(args[0]);
  const timesRV = toNumberRV(args[1]);
  if (isError(timesRV)) {
    return timesRV;
  }
  const times = timesRV.value;
  return rvString(text.repeat(Math.max(0, Math.floor(times))));
};

// ============================================================================
// TEXT (complex number/date formatting)
// ============================================================================

export const fnTEXT: NativeFn = args => {
  const rawVal = topLeft(args[0]);
  if (isError(rawVal)) {
    return rawVal;
  }
  const fmt = toStringRV(args[1]);

  // "@" format = return text as-is
  if (fmt === "@") {
    return rvString(toStringRV(rawVal));
  }

  // Conditional sections: positive;negative;zero (or positive;negative)
  const sections = splitFormatSections(fmt);
  const valRV = toNumberRV(rawVal);
  if (isError(valRV)) {
    return valRV;
  }
  const val = valRV.value;
  let activeFmt: string;
  if (sections.length >= 3) {
    activeFmt = val > 0 ? sections[0] : val < 0 ? sections[1] : sections[2];
  } else if (sections.length === 2) {
    activeFmt = val >= 0 ? sections[0] : sections[1];
  } else {
    activeFmt = sections[0];
  }
  // For negative section, use absolute value (sign is in the format)
  const useVal = sections.length >= 2 && val < 0 ? Math.abs(val) : val;

  return rvString(formatWithCode(useVal, activeFmt, rawVal));
};

/**
 * Split format string on `;` separators, respecting quoted strings.
 */
function splitFormatSections(fmt: string): string[] {
  const sections: string[] = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < fmt.length; i++) {
    if (fmt[i] === '"') {
      inQuotes = !inQuotes;
      current += fmt[i];
    } else if (fmt[i] === ";" && !inQuotes) {
      sections.push(current);
      current = "";
    } else {
      current += fmt[i];
    }
  }
  sections.push(current);
  return sections;
}

/**
 * Format a number using a single format code section.
 */
function formatWithCode(val: number, fmt: string, rawVal: ScalarValue): string {
  const upper = fmt.toUpperCase();

  // Date/time formats — detect by presence of date/time tokens
  if (
    upper.includes("YYYY") ||
    upper.includes("YY") ||
    upper.includes("MMMM") ||
    upper.includes("MMM") ||
    upper.includes("DDDD") ||
    upper.includes("DDD") ||
    upper.includes("DD") ||
    /(?:^|[^H])MM(?!M)/.test(upper) ||
    /(?:^|[^M])M(?!M)/.test(upper.replace(/MMMM|MMM/g, "")) ||
    upper.includes("HH") ||
    /(?:^|[^H])H(?!H)/.test(upper) ||
    upper.includes("SS") ||
    upper.includes("AM/PM") ||
    upper.includes("A/P")
  ) {
    return formatDate(val, fmt);
  }

  // Percentage format
  if (fmt.includes("%")) {
    const stripped = fmt.replace(/[^0#.%,]/g, "");
    const dotIdx = stripped.indexOf(".");
    const afterDot =
      dotIdx >= 0
        ? stripped
            .slice(dotIdx + 1)
            .replace(/%/g, "")
            .replace(/,/g, "")
        : "";
    const decimals = afterDot.length;
    const pctVal = val * 100;
    let result = pctVal.toFixed(decimals);
    if (fmt.includes(",")) {
      const parts = result.split(".");
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      result = parts.join(".");
    }
    return result + "%";
  }

  // Scientific notation: 0.00E+00
  if (/[0#]\.?[0#]*E[+-][0#]+/i.test(fmt)) {
    return formatScientific(val, fmt);
  }

  // Fraction format: # ?/? or # ??/??
  if (/[?]+\/[?]+/.test(fmt)) {
    return formatFraction(val, fmt);
  }

  // Number format with 0 and #
  if (fmt.includes("0") || fmt.includes("#")) {
    return formatNumber(val, fmt);
  }

  return String(val);
}

/**
 * Format number using #, 0, comma patterns.
 */
function formatNumber(val: number, fmt: string): string {
  // Strip literal strings (quoted text) and currency symbols for analysis, but preserve for output
  const literals: { pos: number; text: string }[] = [];
  let cleanFmt = "";
  let inQuotes = false;
  let quoteText = "";
  for (let i = 0; i < fmt.length; i++) {
    if (fmt[i] === '"') {
      if (inQuotes) {
        inQuotes = false;
        // record literal at this position in the clean format
        literals.push({ pos: cleanFmt.length, text: quoteText });
        quoteText = "";
      } else {
        inQuotes = true;
      }
    } else if (inQuotes) {
      quoteText += fmt[i];
    } else {
      cleanFmt += fmt[i];
    }
  }

  // Check for $ prefix
  const hasDollar = cleanFmt.includes("$");
  const hasComma = cleanFmt.includes(",");
  const analysisFormat = cleanFmt.replace(/\$/g, "").replace(/[^0#.,]/g, "");

  // Count decimal places
  const dotIdx = analysisFormat.indexOf(".");
  let decimals = 0;
  if (dotIdx >= 0) {
    decimals = analysisFormat.slice(dotIdx + 1).replace(/,/g, "").length;
  }

  let result = val.toFixed(decimals);
  if (hasComma) {
    const parts = result.split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    result = parts.join(".");
  }

  // Apply prefix/suffix from literals
  let output = "";
  if (hasDollar) {
    output += "$";
  }
  for (const lit of literals) {
    if (lit.pos === 0) {
      output = lit.text + output;
    }
  }
  output += result;
  for (const lit of literals) {
    if (lit.pos > 0) {
      output += lit.text;
    }
  }

  return output;
}

/**
 * Format a number in scientific notation: 0.00E+00
 */
function formatScientific(val: number, fmt: string): string {
  const upper = fmt.toUpperCase();
  const eIdx = upper.indexOf("E");
  const beforeE = fmt.slice(0, eIdx);
  const afterE = fmt.slice(eIdx + 1);

  // Count decimal places in mantissa
  const dotIdx = beforeE.indexOf(".");
  const mantissaDecimals = dotIdx >= 0 ? beforeE.slice(dotIdx + 1).replace(/[^0#]/g, "").length : 0;

  // Count exponent digits
  const expSign = afterE[0] === "+" || afterE[0] === "-" ? afterE[0] : "+";
  const expDigits = afterE.replace(/[^0#]/g, "").length;

  const exp = val === 0 ? 0 : Math.floor(Math.log10(Math.abs(val)));
  const mantissa = val / Math.pow(10, exp);
  const expStr =
    (exp >= 0 && expSign === "+" ? "+" : exp < 0 ? "-" : "") +
    String(Math.abs(exp)).padStart(expDigits, "0");

  return mantissa.toFixed(mantissaDecimals) + "E" + expStr;
}

/**
 * Format a number as a fraction: "# ?/?" or "# ??/??"
 */
function formatFraction(val: number, fmt: string): string {
  const whole = Math.floor(Math.abs(val));
  const frac = Math.abs(val) - whole;
  const sign = val < 0 ? "-" : "";

  if (frac === 0) {
    if (fmt.includes("#") || fmt.includes("0")) {
      return sign + whole + "      ";
    }
    return sign + whole + " 0/1";
  }

  // Determine max denominator from ? count
  const slashIdx = fmt.indexOf("/");
  const denomPattern = fmt.slice(slashIdx + 1).replace(/[^?0#]/g, "");
  const maxDenom = Math.pow(10, denomPattern.length) - 1;

  // Find best fraction approximation
  let bestNum = 0;
  let bestDen = 1;
  let bestError = Math.abs(frac);
  for (let d = 1; d <= maxDenom; d++) {
    const n = Math.round(frac * d);
    const error = Math.abs(frac - n / d);
    if (error < bestError) {
      bestError = error;
      bestNum = n;
      bestDen = d;
      if (error === 0) {
        break;
      }
    }
  }

  const hasWholePart = fmt.indexOf("?") > 0 && fmt[0] !== "?" && fmt[0] !== "/";
  if (hasWholePart) {
    const numStr = String(bestNum).padStart(denomPattern.length, " ");
    const denStr = String(bestDen).padStart(denomPattern.length, " ");
    return sign + (whole > 0 ? whole + " " : "") + numStr + "/" + denStr;
  }
  return sign + (whole * bestDen + bestNum) + "/" + bestDen;
}

/**
 * Day name arrays for date formatting.
 */
const DAY_NAMES_FULL = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday"
];
const DAY_NAMES_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const MONTH_NAMES_FULL = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
];
const MONTH_NAMES_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec"
];

/**
 * Format a numeric Excel serial date with a date/time format code.
 *
 * In the native value system, dates are always numbers (serial values).
 * We always use excelToDate() to convert.
 */
function formatDate(val: number, fmt: string): string {
  const d = excelToDate(val);

  const year = d.getFullYear();
  const month0 = d.getMonth(); // 0-based
  const day = d.getDate();
  const dow = d.getDay();
  // Time from the fractional part of the serial
  const fracDay = val % 1;
  const totalSeconds = Math.round(Math.abs(fracDay) * 86400);
  const hours = Math.floor(totalSeconds / 3600) % 24;
  const minutes = Math.floor(totalSeconds / 60) % 60;
  const seconds = totalSeconds % 60;

  // Check for AM/PM format
  const hasAmPm = /AM\/PM/i.test(fmt) || /A\/P/i.test(fmt);
  const h12 = hours % 12 === 0 ? 12 : hours % 12;
  const ampm = hours < 12 ? "AM" : "PM";
  const ap = hours < 12 ? "A" : "P";

  // Build result by processing format tokens
  let result = fmt;

  // Replace 4-char year first, then 2-char
  result = result.replace(/YYYY/gi, String(year));
  result = result.replace(/YY/gi, String(year).slice(-2));

  // Day-of-week names (must be before DD/D to avoid conflict)
  result = result.replace(/DDDD/gi, DAY_NAMES_FULL[dow]);
  result = result.replace(/DDD/gi, DAY_NAMES_SHORT[dow]);

  // Day of month
  result = result.replace(/DD/g, String(day).padStart(2, "0"));
  result = result.replace(/(?<![A-Za-z])D(?![A-Za-z])/g, String(day));

  // Month names (longest first)
  result = result.replace(/MMMM/gi, MONTH_NAMES_FULL[month0]);
  result = result.replace(/MMM/gi, MONTH_NAMES_SHORT[month0]);
  // MM and M — but only when not adjacent to time tokens
  result = result.replace(/MM/g, String(month0 + 1).padStart(2, "0"));
  result = result.replace(/(?<![A-Za-z])M(?![A-Za-z])/g, String(month0 + 1));

  // Time tokens
  if (hasAmPm) {
    result = result.replace(/HH/gi, String(h12).padStart(2, "0"));
    result = result.replace(/(?<![A-Za-z])H(?![A-Za-z])/gi, String(h12));
  } else {
    result = result.replace(/HH/gi, String(hours).padStart(2, "0"));
    result = result.replace(/(?<![A-Za-z])H(?![A-Za-z])/gi, String(hours));
  }
  result = result.replace(/SS/gi, String(seconds).padStart(2, "0"));
  result = result.replace(/(?<![A-Za-z])S(?![A-Za-z])/gi, String(seconds));

  // Minutes (mm/m in time context — already handled above as month, but Excel
  // resolves ambiguity based on proximity to H/S. For simplicity, any remaining
  // mm in the result are treated as minutes if hours/seconds were in the format)
  // Since we already replaced MM as month, for formats like "hh:mm:ss" the MM
  // was replaced as month digits which happens to be correct for pure time formats
  // only when the serial is a time fraction. For mixed date+time formats the
  // month replacement is correct.
  // However, for pure time formats like "h:mm", the MM should be minutes.
  // Re-scan: if original format had H or S, replace month-like mm with minutes.
  const upperFmt = fmt.toUpperCase();
  if (upperFmt.includes("H") || upperFmt.includes("S")) {
    // We need to re-process: find mm/m that appear after H or before S
    // Simpler approach: if format has H, replace the result's month values
    // that appear between time delimiters with minute values.
    // Since month and minutes share MM, let's detect if this is a pure time format.
    const hasDatePart =
      upperFmt.includes("Y") ||
      upperFmt.includes("D") ||
      upperFmt.includes("MMMM") ||
      upperFmt.includes("MMM");
    if (!hasDatePart) {
      // Pure time format — re-do: replace month-replaced values with minutes
      result = fmt;
      if (hasAmPm) {
        result = result.replace(/HH/gi, String(h12).padStart(2, "0"));
        result = result.replace(/(?<![A-Za-z])H(?![A-Za-z])/gi, String(h12));
      } else {
        result = result.replace(/HH/gi, String(hours).padStart(2, "0"));
        result = result.replace(/(?<![A-Za-z])H(?![A-Za-z])/gi, String(hours));
      }
      result = result.replace(/MM/gi, String(minutes).padStart(2, "0"));
      result = result.replace(/(?<![A-Za-z])M(?![A-Za-z])/gi, String(minutes));
      result = result.replace(/SS/gi, String(seconds).padStart(2, "0"));
      result = result.replace(/(?<![A-Za-z])S(?![A-Za-z])/gi, String(seconds));
    }
  }

  // AM/PM or A/P
  result = result.replace(/AM\/PM/gi, ampm);
  result = result.replace(/A\/P/gi, ap);

  return result;
}

// ============================================================================
// VALUE / EXACT
// ============================================================================

export const fnVALUE: NativeFn = args => {
  const s = toStringRV(args[0]).trim();
  const n = Number(s);
  return isNaN(n) ? ERRORS.VALUE : rvNumber(n);
};

export const fnEXACT: NativeFn = args => {
  return rvBoolean(toStringRV(args[0]) === toStringRV(args[1]));
};

// ============================================================================
// Additional Text Functions
// ============================================================================

export const fnCODE: NativeFn = args => {
  const text = toStringRV(args[0]);
  return text.length > 0 ? rvNumber(text.charCodeAt(0)) : ERRORS.VALUE;
};

export const fnCHAR: NativeFn = args => {
  const nRV = toNumberRV(args[0]);
  if (isError(nRV)) {
    return nRV;
  }
  return rvString(String.fromCharCode(nRV.value));
};

export const fnCLEAN: NativeFn = args => {
  const text = toStringRV(args[0]);
  // Remove non-printable ASCII control characters (0x00-0x1F)
  let result = "";
  for (let i = 0; i < text.length; i++) {
    if (text.charCodeAt(i) >= 32) {
      result += text[i];
    }
  }
  return rvString(result);
};

export const fnT: NativeFn = args => {
  const v = topLeft(args[0]);
  return v.kind === RVKind.String ? v : rvString("");
};

// ============================================================================
// More Text Functions
// ============================================================================

export const fnUNICHAR: NativeFn = args => {
  const nRV = toNumberRV(args[0]);
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
};

export const fnUNICODE: NativeFn = args => {
  const text = toStringRV(args[0]);
  if (text.length === 0) {
    return ERRORS.VALUE;
  }
  const cp = text.codePointAt(0);
  return cp !== undefined ? rvNumber(cp) : ERRORS.VALUE;
};

export const fnBAHTTEXT: NativeFn = args => rvString(toStringRV(args[0]));

export const fnDOLLAR: NativeFn = args => {
  const numRV = toNumberRV(args[0]);
  if (isError(numRV)) {
    return numRV;
  }
  const num = numRV.value;
  let decimals: number;
  if (args.length > 1) {
    const decRV = toNumberRV(args[1]);
    if (isError(decRV)) {
      return decRV;
    }
    decimals = decRV.value;
  } else {
    decimals = 2;
  }
  const d = Math.max(0, Math.floor(decimals));
  const formatted = Math.abs(num).toFixed(d);
  const parts = formatted.split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const result = parts.join(".");
  return rvString(num < 0 ? `($${result})` : `$${result}`);
};

export const fnFIXED: NativeFn = args => {
  const numRV = toNumberRV(args[0]);
  if (isError(numRV)) {
    return numRV;
  }
  const num = numRV.value;
  let decimals: number;
  if (args.length > 1) {
    const decRV = toNumberRV(args[1]);
    if (isError(decRV)) {
      return decRV;
    }
    decimals = decRV.value;
  } else {
    decimals = 2;
  }
  let noCommas: boolean;
  if (args.length > 2) {
    const ncRV = toBooleanRV(args[2]);
    if (isError(ncRV)) {
      return ncRV;
    }
    noCommas = ncRV.value;
  } else {
    noCommas = false;
  }
  const d = Math.max(0, Math.floor(decimals));
  let result = num.toFixed(d);
  if (!noCommas) {
    const parts = result.split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    result = parts.join(".");
  }
  return rvString(result);
};

export const fnASC: NativeFn = args => {
  const text = toStringRV(args[0]);
  return rvString(
    text.replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
  );
};

export const fnDBCS: NativeFn = args => {
  const text = toStringRV(args[0]);
  return rvString(text.replace(/[!-~]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0xfee0)));
};

export const fnJIS: NativeFn = args => fnDBCS(args);

export const fnPHONETIC: NativeFn = args => rvString(toStringRV(args[0]));

export const fnNUMBERVALUE: NativeFn = args => {
  let text = toStringRV(args[0]);
  const decSep = args.length > 1 ? toStringRV(args[1]) : ".";
  const grpSep = args.length > 2 ? toStringRV(args[2]) : ",";
  text = text.split(grpSep).join("");
  if (decSep !== ".") {
    text = text.replace(decSep, ".");
  }
  // Handle percentage
  const isPct = text.endsWith("%");
  if (isPct) {
    text = text.slice(0, -1);
  }
  const n = Number(text);
  if (isNaN(n)) {
    return ERRORS.VALUE;
  }
  return rvNumber(isPct ? n / 100 : n);
};

// ============================================================================
// Excel 365 Text Functions: TEXTBEFORE, TEXTAFTER, TEXTSPLIT
// ============================================================================

export const fnTEXTBEFORE: NativeFn = args => {
  const text = toStringRV(args[0]);
  const delimiter = toStringRV(args[1]);
  let inst: number;
  if (args.length > 2) {
    const instRV = toNumberRV(args[2]);
    if (isError(instRV)) {
      return instRV;
    }
    inst = instRV.value;
  } else {
    inst = 1;
  }
  if (inst === 0) {
    return ERRORS.VALUE;
  }
  if (delimiter === "") {
    return rvString(inst > 0 ? "" : text);
  }
  // Find the nth occurrence
  if (inst > 0) {
    let pos = -1;
    for (let i = 0; i < inst; i++) {
      pos = text.indexOf(delimiter, pos + 1);
      if (pos === -1) {
        // match_mode/if_not_found handling simplified
        return args.length > 5 ? args[5] : ERRORS.NA;
      }
    }
    return rvString(text.slice(0, pos));
  }
  // Negative: search from end
  let pos = text.length;
  for (let i = 0; i < -inst; i++) {
    pos = text.lastIndexOf(delimiter, pos - 1);
    if (pos === -1) {
      return args.length > 5 ? args[5] : ERRORS.NA;
    }
  }
  return rvString(text.slice(0, pos));
};

export const fnTEXTAFTER: NativeFn = args => {
  const text = toStringRV(args[0]);
  const delimiter = toStringRV(args[1]);
  let inst: number;
  if (args.length > 2) {
    const instRV = toNumberRV(args[2]);
    if (isError(instRV)) {
      return instRV;
    }
    inst = instRV.value;
  } else {
    inst = 1;
  }
  if (inst === 0) {
    return ERRORS.VALUE;
  }
  if (delimiter === "") {
    return rvString(inst > 0 ? text : "");
  }
  if (inst > 0) {
    let pos = -1;
    for (let i = 0; i < inst; i++) {
      pos = text.indexOf(delimiter, pos + 1);
      if (pos === -1) {
        return args.length > 5 ? args[5] : ERRORS.NA;
      }
    }
    return rvString(text.slice(pos + delimiter.length));
  }
  let pos = text.length;
  for (let i = 0; i < -inst; i++) {
    pos = text.lastIndexOf(delimiter, pos - 1);
    if (pos === -1) {
      return args.length > 5 ? args[5] : ERRORS.NA;
    }
  }
  return rvString(text.slice(pos + delimiter.length));
};

export const fnTEXTSPLIT: NativeFn = args => {
  const text = toStringRV(args[0]);
  const colDelimiter = args.length > 1 ? toStringRV(args[1]) : "";
  const rowDelimiter = args.length > 2 && args[2].kind !== RVKind.Blank ? toStringRV(args[2]) : "";

  let rows: string[];
  if (rowDelimiter) {
    rows = text.split(rowDelimiter);
  } else {
    rows = [text];
  }

  const result: ScalarValue[][] = [];
  for (const row of rows) {
    if (colDelimiter) {
      result.push(row.split(colDelimiter).map(s => rvString(s)));
    } else {
      result.push([rvString(row)]);
    }
  }
  return rvArray(result);
};
