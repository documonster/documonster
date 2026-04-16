/**
 * Built-in Excel Functions
 *
 * Implements the most commonly used Excel functions.
 * Each function receives an array of already-evaluated arguments
 * and returns a FormulaResult.
 */

import type { CellErrorValue } from "@excel/types";
import { dateToExcel, excelToDate } from "@utils/utils.base";

// ============================================================================
// Result Types
// ============================================================================

export type CalcValue = number | string | boolean | Date | CellErrorValue | null;
export type CalcArray = CalcValue[][] & {
  /** Origin row of this array in the worksheet (1-based). Set by resolveRange. */
  _originRow?: number;
  /** Origin column of this array in the worksheet (1-based). Set by resolveRange. */
  _originCol?: number;
};

export function isError(v: CalcValue): v is CellErrorValue {
  return v !== null && typeof v === "object" && !(v instanceof Date) && "error" in v;
}

export function toNumber(v: CalcValue): number | CellErrorValue {
  if (v === null || v === "") {
    return 0;
  }
  if (typeof v === "number") {
    return v;
  }
  if (typeof v === "boolean") {
    return v ? 1 : 0;
  }
  if (typeof v === "string") {
    const n = Number(v);
    if (isNaN(n)) {
      return { error: "#VALUE!" };
    }
    return n;
  }
  if (v instanceof Date) {
    return dateToExcel(v);
  }
  if (isError(v)) {
    return v;
  }
  return { error: "#VALUE!" };
}

function toBoolean(v: CalcValue): boolean | CellErrorValue {
  if (v === null) {
    return false;
  }
  if (typeof v === "boolean") {
    return v;
  }
  if (typeof v === "number") {
    return v !== 0;
  }
  if (typeof v === "string") {
    const u = v.toUpperCase();
    if (u === "TRUE") {
      return true;
    }
    if (u === "FALSE") {
      return false;
    }
    return { error: "#VALUE!" };
  }
  if (isError(v)) {
    return v;
  }
  return { error: "#VALUE!" };
}

function toString(v: CalcValue): string {
  if (v === null) {
    return "";
  }
  if (typeof v === "string") {
    return v;
  }
  if (typeof v === "number") {
    return String(v);
  }
  if (typeof v === "boolean") {
    return v ? "TRUE" : "FALSE";
  }
  if (v instanceof Date) {
    return v.toLocaleDateString();
  }
  if (isError(v)) {
    return v.error;
  }
  return "";
}

/** Flatten all arguments (scalars and arrays) into a flat number array.
 * For range (array) arguments: only numbers are included (Excel skips booleans/strings/nulls).
 * For direct scalar arguments: booleans and numeric strings are coerced to numbers. */
function flattenNumbers(args: (CalcValue | CalcArray)[]): (number | CellErrorValue)[] {
  const result: (number | CellErrorValue)[] = [];
  for (const arg of args) {
    if (Array.isArray(arg)) {
      // Range argument — only include numbers, skip booleans/strings/nulls (Excel behavior)
      for (const row of arg) {
        for (const cell of row) {
          if (isError(cell)) {
            result.push(cell);
          } else if (typeof cell === "number") {
            result.push(cell);
          }
        }
      }
    } else {
      // Direct scalar argument — coerce booleans and numeric strings
      if (isError(arg)) {
        result.push(arg);
      } else if (arg !== null) {
        const n = toNumber(arg);
        if (isError(n)) {
          result.push(n);
        } else {
          result.push(n);
        }
      }
    }
  }
  return result;
}

/** Flatten all values (including non-numeric) for COUNTA etc. */
function flattenAll(args: (CalcValue | CalcArray)[]): CalcValue[] {
  const result: CalcValue[] = [];
  for (const arg of args) {
    if (Array.isArray(arg)) {
      for (const row of arg) {
        for (const cell of row) {
          result.push(cell);
        }
      }
    } else {
      result.push(arg);
    }
  }
  return result;
}

/** Check for first error in a list */
function firstError(values: (number | CellErrorValue)[]): CellErrorValue | null {
  for (const v of values) {
    if (isError(v)) {
      return v;
    }
  }
  return null;
}

// ============================================================================
// Function Type
// ============================================================================

export type ExcelFunction = (args: (CalcValue | CalcArray)[]) => CalcValue | CalcArray;

// ============================================================================
// Math / Aggregate Functions
// ============================================================================

const fnSUM: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  let sum = 0;
  for (const n of nums) {
    sum += n as number;
  }
  return sum;
};

const fnAVERAGE: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return { error: "#DIV/0!" };
  }
  let sum = 0;
  for (const n of nums) {
    sum += n as number;
  }
  return sum / nums.length;
};

const fnMIN: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return 0;
  }
  let min = Infinity;
  for (const n of nums) {
    if ((n as number) < min) {
      min = n as number;
    }
  }
  return min;
};

const fnMAX: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return 0;
  }
  let max = -Infinity;
  for (const n of nums) {
    if ((n as number) > max) {
      max = n as number;
    }
  }
  return max;
};

const fnCOUNT: ExcelFunction = args => {
  let count = 0;
  const all = flattenAll(args);
  for (const v of all) {
    if (typeof v === "number") {
      count++;
    }
  }
  return count;
};

const fnCOUNTA: ExcelFunction = args => {
  let count = 0;
  const all = flattenAll(args);
  for (const v of all) {
    if (v !== null && v !== "") {
      count++;
    }
  }
  return count;
};

const fnCOUNTBLANK: ExcelFunction = args => {
  let count = 0;
  const all = flattenAll(args);
  for (const v of all) {
    if (v === null || v === "") {
      count++;
    }
  }
  return count;
};

const fnPRODUCT: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return 0;
  }
  let product = 1;
  for (const n of nums) {
    product *= n as number;
  }
  return product;
};

const fnSUMPRODUCT: ExcelFunction = args => {
  if (args.length === 0) {
    return { error: "#VALUE!" };
  }
  // All args must be arrays of the same dimensions
  const arrays: CalcArray[] = [];
  for (const a of args) {
    if (!Array.isArray(a)) {
      return { error: "#VALUE!" };
    }
    arrays.push(a);
  }
  const rows = arrays[0].length;
  const cols = arrays[0][0]?.length ?? 0;
  // Verify all same dimensions
  for (const arr of arrays) {
    if (arr.length !== rows) {
      return { error: "#VALUE!" };
    }
    for (const row of arr) {
      if (row.length !== cols) {
        return { error: "#VALUE!" };
      }
    }
  }
  let sum = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let product = 1;
      for (const arr of arrays) {
        const val = arr[r][c];
        if (isError(val)) {
          return val;
        }
        const n = typeof val === "number" ? val : typeof val === "boolean" ? (val ? 1 : 0) : 0;
        product *= n;
      }
      sum += product;
    }
  }
  return sum;
};

// ============================================================================
// Math Functions
// ============================================================================

const fnABS: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.abs(n);
};

const fnCEILING: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const sig = args.length > 1 ? toNumber(args[1] as CalcValue) : 1;
  if (isError(sig)) {
    return sig;
  }
  if (sig === 0) {
    return 0;
  }
  return Math.ceil(num / sig) * sig;
};

const fnFLOOR: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const sig = args.length > 1 ? toNumber(args[1] as CalcValue) : 1;
  if (isError(sig)) {
    return sig;
  }
  if (sig === 0) {
    return { error: "#DIV/0!" };
  }
  return Math.floor(num / sig) * sig;
};

const fnINT: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.floor(n);
};

const fnMOD: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const div = toNumber(args[1] as CalcValue);
  if (isError(div)) {
    return div;
  }
  if (div === 0) {
    return { error: "#DIV/0!" };
  }
  return num - div * Math.floor(num / div);
};

const fnPOWER: ExcelFunction = args => {
  const base = toNumber(args[0] as CalcValue);
  if (isError(base)) {
    return base;
  }
  const exp = toNumber(args[1] as CalcValue);
  if (isError(exp)) {
    return exp;
  }
  const result = Math.pow(base, exp);
  return !isFinite(result) ? ({ error: "#NUM!" } as CellErrorValue) : result;
};

const fnROUND: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const digits = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(digits)) {
    return digits;
  }
  const factor = Math.pow(10, digits);
  return Math.round(num * factor) / factor;
};

const fnROUNDDOWN: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const digits = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(digits)) {
    return digits;
  }
  const factor = Math.pow(10, digits);
  return Math.trunc(num * factor) / factor;
};

const fnROUNDUP: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const digits = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(digits)) {
    return digits;
  }
  const factor = Math.pow(10, digits);
  const truncated = Math.trunc(num * factor);
  return (num * factor === truncated ? truncated : truncated + (num >= 0 ? 1 : -1)) / factor;
};

const fnSQRT: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n < 0) {
    return { error: "#NUM!" };
  }
  return Math.sqrt(n);
};

const fnLN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n <= 0) {
    return { error: "#NUM!" };
  }
  return Math.log(n);
};

const fnLOG: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n <= 0) {
    return { error: "#NUM!" };
  }
  const base = args.length > 1 ? toNumber(args[1] as CalcValue) : 10;
  if (isError(base)) {
    return base;
  }
  if (base <= 0 || base === 1) {
    return { error: "#NUM!" };
  }
  return Math.log(n) / Math.log(base);
};

const fnLOG10: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n <= 0) {
    return { error: "#NUM!" };
  }
  return Math.log10(n);
};

const fnEXP: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.exp(n);
};

const fnPI: ExcelFunction = () => Math.PI;

const fnRAND: ExcelFunction = () => Math.random();

const fnRANDBETWEEN: ExcelFunction = args => {
  const bottom = toNumber(args[0] as CalcValue);
  if (isError(bottom)) {
    return bottom;
  }
  const top = toNumber(args[1] as CalcValue);
  if (isError(top)) {
    return top;
  }
  const lo = Math.ceil(bottom);
  const hi = Math.floor(top);
  return Math.floor(Math.random() * (hi - lo + 1)) + lo;
};

const fnSIGN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.sign(n);
};

// ============================================================================
// Logical Functions
// ============================================================================

const fnIF: ExcelFunction = args => {
  const cond = toBoolean(args[0] as CalcValue);
  if (isError(cond)) {
    return cond;
  }
  if (cond) {
    return (args[1] as CalcValue) ?? true;
  }
  return (args[2] as CalcValue) ?? false;
};

const fnIFS: ExcelFunction = args => {
  for (let i = 0; i < args.length - 1; i += 2) {
    const cond = toBoolean(args[i] as CalcValue);
    if (isError(cond)) {
      return cond;
    }
    if (cond) {
      return (args[i + 1] as CalcValue) ?? null;
    }
  }
  return { error: "#N/A" };
};

const fnAND: ExcelFunction = args => {
  const all = flattenAll(args);
  for (const v of all) {
    if (v === null) {
      continue;
    }
    const b = toBoolean(v);
    if (isError(b)) {
      return b;
    }
    if (!b) {
      return false;
    }
  }
  return true;
};

const fnOR: ExcelFunction = args => {
  const all = flattenAll(args);
  for (const v of all) {
    if (v === null) {
      continue;
    }
    const b = toBoolean(v);
    if (isError(b)) {
      return b;
    }
    if (b) {
      return true;
    }
  }
  return false;
};

const fnNOT: ExcelFunction = args => {
  const b = toBoolean(args[0] as CalcValue);
  if (isError(b)) {
    return b;
  }
  return !b;
};

const fnXOR: ExcelFunction = args => {
  let count = 0;
  const all = flattenAll(args);
  for (const v of all) {
    if (v === null) {
      continue;
    }
    const b = toBoolean(v);
    if (isError(b)) {
      return b;
    }
    if (b) {
      count++;
    }
  }
  return count % 2 === 1;
};

const fnSWITCH: ExcelFunction = args => {
  if (args.length < 3) {
    return { error: "#VALUE!" };
  }
  const expr = args[0] as CalcValue;
  for (let i = 1; i < args.length - 1; i += 2) {
    if (expr === (args[i] as CalcValue)) {
      return (args[i + 1] as CalcValue) ?? null;
    }
  }
  // Default value (odd number of args means last is default)
  if (args.length % 2 === 0) {
    return (args[args.length - 1] as CalcValue) ?? null;
  }
  return { error: "#N/A" };
};

const fnCHOOSE: ExcelFunction = args => {
  const idx = toNumber(args[0] as CalcValue);
  if (isError(idx)) {
    return idx;
  }
  const i = Math.floor(idx);
  if (i < 1 || i >= args.length) {
    return { error: "#VALUE!" };
  }
  return (args[i] as CalcValue) ?? null;
};

const fnIFERROR: ExcelFunction = args => {
  const val = args[0] as CalcValue;
  return isError(val) ? ((args[1] as CalcValue) ?? null) : val;
};

const fnIFNA: ExcelFunction = args => {
  const val = args[0] as CalcValue;
  return isError(val) && val.error === "#N/A" ? ((args[1] as CalcValue) ?? null) : val;
};

// ============================================================================
// Text Functions
// ============================================================================

const fnCONCATENATE: ExcelFunction = args => {
  const parts: string[] = [];
  for (const a of args) {
    if (Array.isArray(a)) {
      for (const row of a) {
        for (const cell of row) {
          parts.push(toString(cell));
        }
      }
    } else {
      parts.push(toString(a as CalcValue));
    }
  }
  return parts.join("");
};

// CONCAT has the same semantics as CONCATENATE
const fnCONCAT: ExcelFunction = fnCONCATENATE;

const fnTEXTJOIN: ExcelFunction = args => {
  if (args.length < 3) {
    return { error: "#VALUE!" };
  }
  const delimiter = toString(args[0] as CalcValue);
  const ignoreEmpty = toBoolean(args[1] as CalcValue);
  if (isError(ignoreEmpty)) {
    return ignoreEmpty;
  }
  const parts: string[] = [];
  for (let i = 2; i < args.length; i++) {
    const a = args[i];
    if (Array.isArray(a)) {
      for (const row of a) {
        for (const cell of row) {
          const s = toString(cell);
          if (ignoreEmpty && s === "") {
            continue;
          }
          parts.push(s);
        }
      }
    } else {
      const s = toString(a as CalcValue);
      if (ignoreEmpty && s === "") {
        continue;
      }
      parts.push(s);
    }
  }
  return parts.join(delimiter);
};

const fnLEFT: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const n = args.length > 1 ? toNumber(args[1] as CalcValue) : 1;
  if (isError(n)) {
    return n;
  }
  return text.slice(0, n);
};

const fnRIGHT: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const n = args.length > 1 ? toNumber(args[1] as CalcValue) : 1;
  if (isError(n)) {
    return n;
  }
  if (n <= 0) {
    return "";
  }
  return text.slice(-n);
};

const fnMID: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const startNum = toNumber(args[1] as CalcValue);
  if (isError(startNum)) {
    return startNum;
  }
  const numChars = toNumber(args[2] as CalcValue);
  if (isError(numChars)) {
    return numChars;
  }
  return text.slice(startNum - 1, startNum - 1 + numChars);
};

const fnLEN: ExcelFunction = args => {
  return toString(args[0] as CalcValue).length;
};

const fnTRIM: ExcelFunction = args => {
  return toString(args[0] as CalcValue)
    .trim()
    .replace(/\s+/g, " ");
};

const fnLOWER: ExcelFunction = args => {
  return toString(args[0] as CalcValue).toLowerCase();
};

const fnUPPER: ExcelFunction = args => {
  return toString(args[0] as CalcValue).toUpperCase();
};

const fnPROPER: ExcelFunction = args => {
  return toString(args[0] as CalcValue).replace(
    /\w\S*/g,
    t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase()
  );
};

const fnSUBSTITUTE: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const oldText = toString(args[1] as CalcValue);
  const newText = toString(args[2] as CalcValue);
  if (args.length > 3) {
    const instanceNum = toNumber(args[3] as CalcValue);
    if (isError(instanceNum)) {
      return instanceNum;
    }
    let count = 0;
    return text.replace(new RegExp(escapeRegex(oldText), "g"), match => {
      count++;
      return count === instanceNum ? newText : match;
    });
  }
  return text.split(oldText).join(newText);
};

const fnREPLACE: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const startNum = toNumber(args[1] as CalcValue);
  if (isError(startNum)) {
    return startNum;
  }
  const numChars = toNumber(args[2] as CalcValue);
  if (isError(numChars)) {
    return numChars;
  }
  const newText = toString(args[3] as CalcValue);
  return text.slice(0, startNum - 1) + newText + text.slice(startNum - 1 + numChars);
};

const fnFIND: ExcelFunction = args => {
  const findText = toString(args[0] as CalcValue);
  const withinText = toString(args[1] as CalcValue);
  const startNum = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(startNum)) {
    return startNum;
  }
  const idx = withinText.indexOf(findText, startNum - 1);
  return idx === -1 ? ({ error: "#VALUE!" } as CellErrorValue) : idx + 1;
};

const fnSEARCH: ExcelFunction = args => {
  let findText = toString(args[0] as CalcValue);
  const withinText = toString(args[1] as CalcValue);
  const startNum = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(startNum)) {
    return startNum;
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
    return match ? match.index + startNum : ({ error: "#VALUE!" } as CellErrorValue);
  } catch {
    // If regex is invalid, fall back to simple indexOf
    findText = findText.toLowerCase();
    const idx = withinText.toLowerCase().indexOf(findText, startNum - 1);
    return idx === -1 ? ({ error: "#VALUE!" } as CellErrorValue) : idx + 1;
  }
};

const fnREPT: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const times = toNumber(args[1] as CalcValue);
  if (isError(times)) {
    return times;
  }
  return text.repeat(Math.max(0, Math.floor(times)));
};

const fnTEXT: ExcelFunction = args => {
  const rawVal = args[0] as CalcValue;
  if (isError(rawVal)) {
    return rawVal;
  }
  const fmt = toString(args[1] as CalcValue);

  // "@" format = return text as-is
  if (fmt === "@") {
    return toString(rawVal);
  }

  // Conditional sections: positive;negative;zero (or positive;negative)
  const sections = splitFormatSections(fmt);
  const val = toNumber(rawVal);
  if (isError(val)) {
    return val;
  }
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

  return formatWithCode(useVal, activeFmt, rawVal);
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
function formatWithCode(val: number, fmt: string, rawVal: CalcValue): string {
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
    return formatDate(val, fmt, rawVal);
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
 */
function formatDate(val: number, fmt: string, rawVal: CalcValue): string {
  const d = rawVal instanceof Date ? (rawVal as Date) : excelToDate(val);

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

const fnVALUE: ExcelFunction = args => {
  const s = toString(args[0] as CalcValue).trim();
  const n = Number(s);
  return isNaN(n) ? ({ error: "#VALUE!" } as CellErrorValue) : n;
};

const fnEXACT: ExcelFunction = args => {
  return toString(args[0] as CalcValue) === toString(args[1] as CalcValue);
};

// ============================================================================
// Information Functions
// ============================================================================

const fnISNUMBER: ExcelFunction = args => typeof (args[0] as CalcValue) === "number";
const fnISTEXT: ExcelFunction = args => typeof (args[0] as CalcValue) === "string";
const fnISBLANK: ExcelFunction = args => (args[0] as CalcValue) === null;
const fnISLOGICAL: ExcelFunction = args => typeof (args[0] as CalcValue) === "boolean";
const fnISERROR: ExcelFunction = args => isError(args[0] as CalcValue);
const fnISERR: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  return isError(v) && v.error !== "#N/A";
};
const fnISNA: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  return isError(v) && v.error === "#N/A";
};
const fnISNONTEXT: ExcelFunction = args => typeof (args[0] as CalcValue) !== "string";

const fnN: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  if (typeof v === "number") {
    return v;
  }
  if (typeof v === "boolean") {
    return v ? 1 : 0;
  }
  if (isError(v)) {
    return v;
  }
  return 0;
};

const fnTYPE: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  if (typeof v === "number") {
    return 1;
  }
  if (typeof v === "string") {
    return 2;
  }
  if (typeof v === "boolean") {
    return 4;
  }
  if (isError(v)) {
    return 16;
  }
  return 1; // null → number
};

// ============================================================================
// Lookup / Reference Functions
// ============================================================================

const fnROW: ExcelFunction = args => {
  if (args.length === 0) {
    return { error: "#VALUE!" };
  } // needs context
  // If arg is a cell ref, evaluator passes row number
  const v = args[0] as CalcValue;
  return typeof v === "number" ? v : { error: "#VALUE!" };
};

const fnCOLUMN: ExcelFunction = args => {
  if (args.length === 0) {
    return { error: "#VALUE!" };
  }
  const v = args[0] as CalcValue;
  return typeof v === "number" ? v : { error: "#VALUE!" };
};

const fnROWS: ExcelFunction = args => {
  if (Array.isArray(args[0])) {
    return (args[0] as CalcArray).length;
  }
  return 1;
};

const fnCOLUMNS: ExcelFunction = args => {
  if (Array.isArray(args[0])) {
    const arr = args[0] as CalcArray;
    return arr.length > 0 ? arr[0].length : 0;
  }
  return 1;
};

const fnINDEX: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return args[0] as CalcValue;
  }
  const arr = args[0] as CalcArray;
  const rowNum = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(rowNum)) {
    return rowNum;
  }
  const colNum = args.length > 2 ? toNumber(args[2] as CalcValue) : 0;
  if (isError(colNum)) {
    return colNum;
  }

  if (rowNum === 0 && colNum === 0) {
    // Return entire array
    return arr;
  }

  // rowNum=0: return entire column as array
  if (rowNum === 0) {
    const c = colNum - 1;
    if (c < 0 || c >= (arr[0]?.length ?? 0)) {
      return { error: "#REF!" };
    }
    return arr.map(row => [row[c]]);
  }

  // colNum=0: return entire row as array
  if (colNum === 0) {
    const r = rowNum - 1;
    if (r < 0 || r >= arr.length) {
      return { error: "#REF!" };
    }
    return [arr[r]];
  }

  // Single cell
  const r = rowNum - 1;
  const c = colNum - 1;
  if (r < 0 || r >= arr.length || c < 0 || c >= (arr[0]?.length ?? 0)) {
    return { error: "#REF!" };
  }
  return arr[r][c];
};

const fnMATCH: ExcelFunction = args => {
  const lookupValue = args[0] as CalcValue;
  if (!Array.isArray(args[1])) {
    return { error: "#N/A" };
  }
  const lookupArr = args[1] as CalcArray;
  const matchType = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(matchType)) {
    return matchType;
  }

  // Flatten to 1D
  const flat: CalcValue[] = [];
  for (const row of lookupArr) {
    for (const cell of row) {
      flat.push(cell);
    }
  }

  if (matchType === 0) {
    // Exact match (with wildcard support for string lookups)
    const lookupStr = typeof lookupValue === "string" ? lookupValue : null;
    const hasWildcard = lookupStr !== null && (lookupStr.includes("*") || lookupStr.includes("?"));
    let wildcardRe: RegExp | null = null;
    if (hasWildcard) {
      const pattern = lookupStr
        .replace(/[.*+^${}()|[\]\\]/g, m => (m === "*" || m === "?" ? m : "\\" + m))
        .replace(/\*/g, ".*")
        .replace(/\?/g, ".");
      try {
        wildcardRe = new RegExp("^" + pattern + "$", "i");
      } catch {
        wildcardRe = null;
      }
    }
    for (let i = 0; i < flat.length; i++) {
      if (flat[i] === lookupValue) {
        return i + 1;
      }
      if (typeof flat[i] === "string" && typeof lookupValue === "string") {
        if (wildcardRe) {
          if (wildcardRe.test(flat[i] as string)) {
            return i + 1;
          }
        } else if ((flat[i] as string).toLowerCase() === lookupValue.toLowerCase()) {
          return i + 1;
        }
      }
    }
    return { error: "#N/A" };
  }

  if (matchType === 1 || matchType > 0) {
    // Data is sorted ascending. Find largest value <= lookupValue.
    let bestIdx = -1;
    for (let i = 0; i < flat.length; i++) {
      const v = flat[i];
      if (typeof v === typeof lookupValue) {
        if (typeof v === "number" && typeof lookupValue === "number") {
          if (v <= lookupValue) {
            bestIdx = i;
          } else {
            break;
          }
        } else if (typeof v === "string" && typeof lookupValue === "string") {
          if (v.toLowerCase() <= lookupValue.toLowerCase()) {
            bestIdx = i;
          } else {
            break;
          }
        }
      }
    }
    return bestIdx >= 0 ? bestIdx + 1 : ({ error: "#N/A" } as CellErrorValue);
  }

  // matchType === -1: Data is sorted descending. Find smallest value >= lookupValue.
  let bestIdx = -1;
  for (let i = 0; i < flat.length; i++) {
    const v = flat[i];
    if (typeof v === typeof lookupValue) {
      if (typeof v === "number" && typeof lookupValue === "number") {
        if (v >= lookupValue) {
          bestIdx = i;
        } else {
          break;
        }
      } else if (typeof v === "string" && typeof lookupValue === "string") {
        if (v.toLowerCase() >= lookupValue.toLowerCase()) {
          bestIdx = i;
        } else {
          break;
        }
      }
    }
  }
  return bestIdx >= 0 ? bestIdx + 1 : ({ error: "#N/A" } as CellErrorValue);
};

const fnVLOOKUP: ExcelFunction = args => {
  const lookupValue = args[0] as CalcValue;
  if (!Array.isArray(args[1])) {
    return { error: "#N/A" };
  }
  const table = args[1] as CalcArray;
  const colIndex = toNumber(args[2] as CalcValue);
  if (isError(colIndex)) {
    return colIndex;
  }
  const rangeLookup = args.length > 3 ? toBoolean(args[3] as CalcValue) : true;
  if (isError(rangeLookup)) {
    return rangeLookup;
  }

  if (colIndex < 1 || colIndex > (table[0]?.length ?? 0)) {
    return { error: "#REF!" };
  }

  if (!rangeLookup) {
    // Exact match
    for (const row of table) {
      if (row[0] === lookupValue) {
        return row[colIndex - 1];
      }
      if (
        typeof row[0] === "string" &&
        typeof lookupValue === "string" &&
        (row[0] as string).toLowerCase() === (lookupValue as string).toLowerCase()
      ) {
        return row[colIndex - 1];
      }
    }
    return { error: "#N/A" };
  }
  // Approximate match: data sorted ascending by first column.
  // Find the largest value <= lookupValue.
  let bestRow: CalcValue[] | null = null;
  for (const row of table) {
    const v = row[0];
    if (typeof v === typeof lookupValue) {
      if (typeof v === "number" && typeof lookupValue === "number") {
        if (v <= lookupValue) {
          bestRow = row;
        } else {
          break;
        }
      } else if (typeof v === "string" && typeof lookupValue === "string") {
        if (v.toLowerCase() <= lookupValue.toLowerCase()) {
          bestRow = row;
        } else {
          break;
        }
      }
    }
  }
  return bestRow ? bestRow[colIndex - 1] : ({ error: "#N/A" } as CellErrorValue);
};

const fnHLOOKUP: ExcelFunction = args => {
  const lookupValue = args[0] as CalcValue;
  if (!Array.isArray(args[1])) {
    return { error: "#N/A" };
  }
  const table = args[1] as CalcArray;
  const rowIndex = toNumber(args[2] as CalcValue);
  if (isError(rowIndex)) {
    return rowIndex;
  }
  const rangeLookup = args.length > 3 ? toBoolean(args[3] as CalcValue) : true;
  if (isError(rangeLookup)) {
    return rangeLookup;
  }

  if (rowIndex < 1 || rowIndex > table.length) {
    return { error: "#REF!" };
  }

  const headerRow = table[0] || [];
  if (!rangeLookup) {
    for (let c = 0; c < headerRow.length; c++) {
      if (headerRow[c] === lookupValue) {
        return table[rowIndex - 1][c];
      }
    }
    return { error: "#N/A" };
  }
  let bestCol = -1;
  for (let c = 0; c < headerRow.length; c++) {
    if (typeof headerRow[c] === typeof lookupValue && headerRow[c]! <= lookupValue!) {
      if (bestCol === -1 || headerRow[c]! >= headerRow[bestCol]!) {
        bestCol = c;
      }
    }
  }
  return bestCol >= 0 ? table[rowIndex - 1][bestCol] : ({ error: "#N/A" } as CellErrorValue);
};

// ============================================================================
// Date Functions (simplified)
// ============================================================================

const fnTODAY: ExcelFunction = () => {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
};

const fnNOW: ExcelFunction = () => new Date();

const fnYEAR: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  if (v instanceof Date) {
    return v.getFullYear();
  }
  const n = toNumber(v);
  if (isError(n)) {
    return n;
  }
  return excelToDate(n).getFullYear();
};

const fnMONTH: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  if (v instanceof Date) {
    return v.getMonth() + 1;
  }
  const n = toNumber(v);
  if (isError(n)) {
    return n;
  }
  return excelToDate(n).getMonth() + 1;
};

const fnDAY: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  if (v instanceof Date) {
    return v.getDate();
  }
  const n = toNumber(v);
  if (isError(n)) {
    return n;
  }
  return excelToDate(n).getDate();
};

// ============================================================================
// Utility
// ============================================================================

export function makeError(code: string): CellErrorValue {
  return { error: code } as CellErrorValue;
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Convert a criteria string (like ">5", "<=10", "apple", "*test?") to a predicate */
function buildCriteriaPredicate(criteria: CalcValue): (v: CalcValue) => boolean {
  if (typeof criteria === "number") {
    return v => typeof v === "number" && v === criteria;
  }
  if (typeof criteria === "boolean") {
    return v => v === criteria;
  }
  const s = toString(criteria);
  // Operator-prefixed criteria
  const opMatch = /^([<>]=?|[<>]|=|<>)(.*)$/.exec(s);
  if (opMatch) {
    const [, op, valStr] = opMatch;
    const numVal = Number(valStr);
    const isNum = !isNaN(numVal) && valStr.trim() !== "";
    return (v: CalcValue) => {
      const vn = typeof v === "number" ? v : NaN;
      const vs = toString(v).toLowerCase();
      const cs = valStr.toLowerCase();
      switch (op) {
        case "=":
          return isNum ? vn === numVal : vs === cs;
        case "<>":
          return isNum ? vn !== numVal : vs !== cs;
        case ">":
          return isNum ? vn > numVal : vs > cs;
        case "<":
          return isNum ? vn < numVal : vs < cs;
        case ">=":
          return isNum ? vn >= numVal : vs >= cs;
        case "<=":
          return isNum ? vn <= numVal : vs <= cs;
        default:
          return false;
      }
    };
  }
  // Wildcard match (case-insensitive)
  if (s.includes("*") || s.includes("?")) {
    const pattern = s
      .replace(/[.*+^${}()|[\]\\]/g, m => (m === "*" || m === "?" ? m : "\\" + m))
      .replace(/\*/g, ".*")
      .replace(/\?/g, ".");
    try {
      const re = new RegExp("^" + pattern + "$", "i");
      return v => re.test(toString(v));
    } catch {
      return v => toString(v).toLowerCase() === s.toLowerCase();
    }
  }
  // Exact match (case-insensitive for strings, numeric for numbers)
  const numCriteria = Number(s);
  if (!isNaN(numCriteria) && s.trim() !== "") {
    return v => typeof v === "number" && v === numCriteria;
  }
  return v => toString(v).toLowerCase() === s.toLowerCase();
}

// ============================================================================
// Conditional Aggregate Functions
// ============================================================================

const fnSUMIF: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const range = args[0] as CalcArray;
  const pred = buildCriteriaPredicate(args[1] as CalcValue);
  const sumRange = args.length > 2 && Array.isArray(args[2]) ? (args[2] as CalcArray) : range;
  let sum = 0;
  for (let r = 0; r < range.length; r++) {
    for (let c = 0; c < range[r].length; c++) {
      if (pred(range[r][c])) {
        const sv = sumRange[r]?.[c];
        if (typeof sv === "number") {
          sum += sv;
        }
      }
    }
  }
  return sum;
};

const fnSUMIFS: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || args.length < 3) {
    return { error: "#VALUE!" };
  }
  const sumRange = args[0] as CalcArray;
  const pairs: { range: CalcArray; pred: (v: CalcValue) => boolean }[] = [];
  for (let i = 1; i < args.length - 1; i += 2) {
    if (!Array.isArray(args[i])) {
      return { error: "#VALUE!" };
    }
    pairs.push({
      range: args[i] as CalcArray,
      pred: buildCriteriaPredicate(args[i + 1] as CalcValue)
    });
  }
  let sum = 0;
  for (let r = 0; r < sumRange.length; r++) {
    for (let c = 0; c < (sumRange[r]?.length ?? 0); c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(p.range[r]?.[c] ?? null)) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        const sv = sumRange[r][c];
        if (typeof sv === "number") {
          sum += sv;
        }
      }
    }
  }
  return sum;
};

const fnCOUNTIF: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const range = args[0] as CalcArray;
  const pred = buildCriteriaPredicate(args[1] as CalcValue);
  let count = 0;
  for (const row of range) {
    for (const cell of row) {
      if (pred(cell)) {
        count++;
      }
    }
  }
  return count;
};

const fnCOUNTIFS: ExcelFunction = args => {
  if (args.length < 2 || !Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const pairs: { range: CalcArray; pred: (v: CalcValue) => boolean }[] = [];
  for (let i = 0; i < args.length - 1; i += 2) {
    if (!Array.isArray(args[i])) {
      return { error: "#VALUE!" };
    }
    pairs.push({
      range: args[i] as CalcArray,
      pred: buildCriteriaPredicate(args[i + 1] as CalcValue)
    });
  }
  const rows = pairs[0].range.length;
  const cols = pairs[0].range[0]?.length ?? 0;
  let count = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(p.range[r]?.[c] ?? null)) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        count++;
      }
    }
  }
  return count;
};

const fnAVERAGEIF: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const range = args[0] as CalcArray;
  const pred = buildCriteriaPredicate(args[1] as CalcValue);
  const avgRange = args.length > 2 && Array.isArray(args[2]) ? (args[2] as CalcArray) : range;
  let sum = 0;
  let count = 0;
  for (let r = 0; r < range.length; r++) {
    for (let c = 0; c < range[r].length; c++) {
      if (pred(range[r][c])) {
        const sv = avgRange[r]?.[c];
        if (typeof sv === "number") {
          sum += sv;
          count++;
        }
      }
    }
  }
  return count === 0 ? ({ error: "#DIV/0!" } as CellErrorValue) : sum / count;
};

const fnAVERAGEIFS: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || args.length < 3) {
    return { error: "#VALUE!" };
  }
  const avgRange = args[0] as CalcArray;
  const pairs: { range: CalcArray; pred: (v: CalcValue) => boolean }[] = [];
  for (let i = 1; i < args.length - 1; i += 2) {
    if (!Array.isArray(args[i])) {
      return { error: "#VALUE!" };
    }
    pairs.push({
      range: args[i] as CalcArray,
      pred: buildCriteriaPredicate(args[i + 1] as CalcValue)
    });
  }
  let sum = 0;
  let count = 0;
  for (let r = 0; r < avgRange.length; r++) {
    for (let c = 0; c < (avgRange[r]?.length ?? 0); c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(p.range[r]?.[c] ?? null)) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        const sv = avgRange[r][c];
        if (typeof sv === "number") {
          sum += sv;
          count++;
        }
      }
    }
  }
  return count === 0 ? ({ error: "#DIV/0!" } as CellErrorValue) : sum / count;
};

const fnMAXIFS: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || args.length < 3) {
    return { error: "#VALUE!" };
  }
  const maxRange = args[0] as CalcArray;
  const pairs: { range: CalcArray; pred: (v: CalcValue) => boolean }[] = [];
  for (let i = 1; i < args.length - 1; i += 2) {
    if (!Array.isArray(args[i])) {
      return { error: "#VALUE!" };
    }
    pairs.push({
      range: args[i] as CalcArray,
      pred: buildCriteriaPredicate(args[i + 1] as CalcValue)
    });
  }
  let result = -Infinity;
  let found = false;
  for (let r = 0; r < maxRange.length; r++) {
    for (let c = 0; c < (maxRange[r]?.length ?? 0); c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(p.range[r]?.[c] ?? null)) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        const sv = maxRange[r][c];
        if (typeof sv === "number") {
          if (sv > result) {
            result = sv;
          }
          found = true;
        }
      }
    }
  }
  return found ? result : 0;
};

const fnMINIFS: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || args.length < 3) {
    return { error: "#VALUE!" };
  }
  const minRange = args[0] as CalcArray;
  const pairs: { range: CalcArray; pred: (v: CalcValue) => boolean }[] = [];
  for (let i = 1; i < args.length - 1; i += 2) {
    if (!Array.isArray(args[i])) {
      return { error: "#VALUE!" };
    }
    pairs.push({
      range: args[i] as CalcArray,
      pred: buildCriteriaPredicate(args[i + 1] as CalcValue)
    });
  }
  let result = Infinity;
  let found = false;
  for (let r = 0; r < minRange.length; r++) {
    for (let c = 0; c < (minRange[r]?.length ?? 0); c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(p.range[r]?.[c] ?? null)) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        const sv = minRange[r][c];
        if (typeof sv === "number") {
          if (sv < result) {
            result = sv;
          }
          found = true;
        }
      }
    }
  }
  return found ? result : 0;
};

// ============================================================================
// Statistical Functions
// ============================================================================

const fnMEDIAN: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return { error: "#NUM!" };
  }
  const sorted = (nums as number[]).slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
};

const fnLARGE: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const nums = flattenNumbers([args[0]]);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  const k = toNumber(args[1] as CalcValue);
  if (isError(k)) {
    return k;
  }
  const sorted = (nums as number[]).slice().sort((a, b) => b - a);
  const idx = Math.floor(k) - 1;
  return idx >= 0 && idx < sorted.length ? sorted[idx] : ({ error: "#NUM!" } as CellErrorValue);
};

const fnSMALL: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const nums = flattenNumbers([args[0]]);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  const k = toNumber(args[1] as CalcValue);
  if (isError(k)) {
    return k;
  }
  const sorted = (nums as number[]).slice().sort((a, b) => a - b);
  const idx = Math.floor(k) - 1;
  return idx >= 0 && idx < sorted.length ? sorted[idx] : ({ error: "#NUM!" } as CellErrorValue);
};

const fnRANK: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  if (!Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const nums = flattenNumbers([args[1]]);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  const order = args.length > 2 ? toNumber(args[2] as CalcValue) : 0;
  if (isError(order)) {
    return order;
  }
  const sorted =
    order === 0
      ? (nums as number[]).slice().sort((a, b) => b - a)
      : (nums as number[]).slice().sort((a, b) => a - b);
  const idx = sorted.indexOf(num);
  return idx === -1 ? ({ error: "#N/A" } as CellErrorValue) : idx + 1;
};

const fnSTDEV: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length < 2) {
    return { error: "#DIV/0!" };
  }
  const n = nums.length;
  let sum = 0;
  for (const v of nums) {
    sum += v as number;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of nums) {
    sumSq += ((v as number) - mean) ** 2;
  }
  return Math.sqrt(sumSq / (n - 1));
};

const fnSTDEVP: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return { error: "#DIV/0!" };
  }
  const n = nums.length;
  let sum = 0;
  for (const v of nums) {
    sum += v as number;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of nums) {
    sumSq += ((v as number) - mean) ** 2;
  }
  return Math.sqrt(sumSq / n);
};

const fnVAR: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length < 2) {
    return { error: "#DIV/0!" };
  }
  const n = nums.length;
  let sum = 0;
  for (const v of nums) {
    sum += v as number;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of nums) {
    sumSq += ((v as number) - mean) ** 2;
  }
  return sumSq / (n - 1);
};

const fnVARP: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return { error: "#DIV/0!" };
  }
  const n = nums.length;
  let sum = 0;
  for (const v of nums) {
    sum += v as number;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of nums) {
    sumSq += ((v as number) - mean) ** 2;
  }
  return sumSq / n;
};

// ============================================================================
// Additional Date/Time Functions
// ============================================================================

const fnDATE: ExcelFunction = args => {
  const year = toNumber(args[0] as CalcValue);
  if (isError(year)) {
    return year;
  }
  const month = toNumber(args[1] as CalcValue);
  if (isError(month)) {
    return month;
  }
  const day = toNumber(args[2] as CalcValue);
  if (isError(day)) {
    return day;
  }

  // Lotus 1-2-3 bug: DATE(1900,2,29) should return serial 60 even though
  // 1900 is not a leap year. JavaScript's Date constructor rolls Feb 29, 1900
  // forward to March 1, 1900, so we handle this specially.
  if (year === 1900 && month === 2 && day === 29) {
    return 60; // The fictitious Feb 29, 1900
  }

  const d = new Date(Date.UTC(year, month - 1, day));
  if (year >= 0 && year < 100) {
    d.setUTCFullYear(year);
  }
  return dateToExcel(d);
};

const fnTIME: ExcelFunction = args => {
  const hour = toNumber(args[0] as CalcValue);
  if (isError(hour)) {
    return hour;
  }
  const minute = toNumber(args[1] as CalcValue);
  if (isError(minute)) {
    return minute;
  }
  const second = toNumber(args[2] as CalcValue);
  if (isError(second)) {
    return second;
  }
  return (hour * 3600 + minute * 60 + second) / 86400;
};

const fnHOUR: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  if (v instanceof Date) {
    return v.getHours();
  }
  const n = toNumber(v);
  if (isError(n)) {
    return n;
  }
  const totalSeconds = Math.round((n % 1) * 86400);
  return Math.floor(totalSeconds / 3600) % 24;
};

const fnMINUTE: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  if (v instanceof Date) {
    return v.getMinutes();
  }
  const n = toNumber(v);
  if (isError(n)) {
    return n;
  }
  const totalSeconds = Math.round((n % 1) * 86400);
  return Math.floor(totalSeconds / 60) % 60;
};

const fnSECOND: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  if (v instanceof Date) {
    return v.getSeconds();
  }
  const n = toNumber(v);
  if (isError(n)) {
    return n;
  }
  const totalSeconds = Math.round((n % 1) * 86400);
  return totalSeconds % 60;
};

const fnWEEKDAY: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  let d: Date;
  if (v instanceof Date) {
    d = v;
  } else {
    const n = toNumber(v);
    if (isError(n)) {
      return n;
    }
    d = excelToDate(n);
  }
  const returnType = args.length > 1 ? toNumber(args[1] as CalcValue) : 1;
  if (isError(returnType)) {
    return returnType;
  }
  const day = d.getDay(); // 0=Sun, 6=Sat
  switch (returnType) {
    case 1:
      return day + 1; // 1=Sun, 7=Sat
    case 2:
      return day === 0 ? 7 : day; // 1=Mon, 7=Sun
    case 3:
      return day === 0 ? 6 : day - 1; // 0=Mon, 6=Sun
    default:
      return day + 1;
  }
};

const fnEOMONTH: ExcelFunction = args => {
  const startDate = toNumber(args[0] as CalcValue);
  if (isError(startDate)) {
    return startDate;
  }
  const months = toNumber(args[1] as CalcValue);
  if (isError(months)) {
    return months;
  }
  const d = args[0] instanceof Date ? (args[0] as Date) : excelToDate(startDate);
  const result = new Date(Date.UTC(d.getFullYear(), d.getMonth() + months + 1, 0));
  return dateToExcel(result);
};

const fnEDATE: ExcelFunction = args => {
  const startDate = toNumber(args[0] as CalcValue);
  if (isError(startDate)) {
    return startDate;
  }
  const months = toNumber(args[1] as CalcValue);
  if (isError(months)) {
    return months;
  }
  const d = args[0] instanceof Date ? (args[0] as Date) : excelToDate(startDate);
  const result = new Date(Date.UTC(d.getFullYear(), d.getMonth() + months, d.getDate()));
  return dateToExcel(result);
};

const fnDATEDIF: ExcelFunction = args => {
  const startN = toNumber(args[0] as CalcValue);
  if (isError(startN)) {
    return startN;
  }
  const endN = toNumber(args[1] as CalcValue);
  if (isError(endN)) {
    return endN;
  }
  const unit = toString(args[2] as CalcValue).toUpperCase();
  const startD = args[0] instanceof Date ? (args[0] as Date) : excelToDate(startN);
  const endD = args[1] instanceof Date ? (args[1] as Date) : excelToDate(endN);
  const sy = startD.getFullYear();
  const sm = startD.getMonth();
  const sd = startD.getDate();
  const ey = endD.getFullYear();
  const em = endD.getMonth();
  const ed = endD.getDate();
  switch (unit) {
    case "Y":
      return ey - sy - (em < sm || (em === sm && ed < sd) ? 1 : 0);
    case "M":
      return (ey - sy) * 12 + em - sm - (ed < sd ? 1 : 0);
    case "D":
      return Math.floor((endD.getTime() - startD.getTime()) / 86400000);
    default:
      return { error: "#NUM!" };
  }
};

// ============================================================================
// Dynamic Array Functions
//
// These functions return CalcArray (2D arrays). The evaluator's top-level
// `evaluateFormula` unwraps arrays to a single scalar (top-left cell value)
// for storage in `cell.result`. Full spill semantics (writing results into
// adjacent cells) are NOT implemented — only the first value is persisted.
// ============================================================================

const fnFILTER: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const data = args[0] as CalcArray;
  const include = args[1] as CalcArray;
  const ifEmpty = args.length > 2 ? (args[2] as CalcValue) : null;

  const result: CalcValue[][] = [];
  for (let r = 0; r < data.length; r++) {
    const inc = include[r]?.[0] ?? null;
    if (inc === true || (typeof inc === "number" && inc !== 0)) {
      result.push(data[r]);
    }
  }
  if (result.length === 0) {
    if (ifEmpty !== null) {
      return [[ifEmpty]];
    }
    return { error: "#VALUE!" };
  }
  return result;
};

const fnSORT: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const data = (args[0] as CalcArray).map(row => [...row]);
  const sortIndex = args.length > 1 ? toNumber(args[1] as CalcValue) : 1;
  if (isError(sortIndex)) {
    return sortIndex;
  }
  const sortOrder = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(sortOrder)) {
    return sortOrder;
  }
  const byCol = args.length > 3 ? toBoolean(args[3] as CalcValue) : false;
  if (isError(byCol)) {
    return byCol;
  }

  if (byCol) {
    // Sort columns instead of rows
    const numCols = data[0]?.length ?? 0;
    const rowIdx = (sortIndex as number) - 1;
    // Build column indices and sort them
    const colIndices = Array.from({ length: numCols }, (_, i) => i);
    colIndices.sort((a, b) => {
      const va = data[rowIdx]?.[a];
      const vb = data[rowIdx]?.[b];
      if (typeof va === "number" && typeof vb === "number") {
        return (va - vb) * (sortOrder as number);
      }
      return toString(va).localeCompare(toString(vb)) * (sortOrder as number);
    });
    return data.map(row => colIndices.map(c => row[c]));
  }

  const col = (sortIndex as number) - 1;
  data.sort((a, b) => {
    const va = a[col];
    const vb = b[col];
    if (typeof va === "number" && typeof vb === "number") {
      return (va - vb) * (sortOrder as number);
    }
    return toString(va).localeCompare(toString(vb)) * (sortOrder as number);
  });
  return data;
};

const fnUNIQUE: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const data = args[0] as CalcArray;
  const byCol = args.length > 1 ? toBoolean(args[1] as CalcValue) : false;
  if (isError(byCol)) {
    return byCol;
  }
  const exactlyOnce = args.length > 2 ? toBoolean(args[2] as CalcValue) : false;
  if (isError(exactlyOnce)) {
    return exactlyOnce;
  }

  if (byCol) {
    // Transpose, apply, transpose back
    const cols = data[0]?.length ?? 0;
    const transposed: CalcValue[][] = [];
    for (let c = 0; c < cols; c++) {
      transposed.push(data.map(row => row[c]));
    }
    const unique = applyUnique(transposed, exactlyOnce as boolean);
    // Transpose back
    if (unique.length === 0) {
      return { error: "#VALUE!" } as CellErrorValue;
    }
    const rows = unique[0].length;
    const result: CalcValue[][] = [];
    for (let r = 0; r < rows; r++) {
      result.push(unique.map(col => col[r]));
    }
    return result;
  }
  const result = applyUnique(data, exactlyOnce as boolean);
  return result.length > 0 ? result : ({ error: "#VALUE!" } as CellErrorValue);
};

function applyUnique(rows: CalcValue[][], exactlyOnce: boolean): CalcValue[][] {
  const keyCount = new Map<string, number>();
  const keyToRows = new Map<string, CalcValue[]>();
  const order: string[] = [];
  for (const row of rows) {
    const key = row.map(c => toString(c)).join("\0");
    if (!keyCount.has(key)) {
      order.push(key);
      keyToRows.set(key, row);
    }
    keyCount.set(key, (keyCount.get(key) ?? 0) + 1);
  }
  const result: CalcValue[][] = [];
  for (const key of order) {
    if (exactlyOnce && (keyCount.get(key) ?? 0) > 1) {
      continue;
    }
    result.push(keyToRows.get(key)!);
  }
  return result;
}

const fnSORTBY: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || args.length < 2) {
    return { error: "#VALUE!" };
  }
  const data = (args[0] as CalcArray).map((row, i) => ({ row, idx: i }));
  // Collect sort keys and orders
  const sortKeys: { arr: CalcArray; order: number }[] = [];
  for (let i = 1; i < args.length; i += 2) {
    if (!Array.isArray(args[i])) {
      return { error: "#VALUE!" };
    }
    const order = i + 1 < args.length ? toNumber(args[i + 1] as CalcValue) : 1;
    if (isError(order)) {
      return order;
    }
    sortKeys.push({ arr: args[i] as CalcArray, order });
  }
  data.sort((a, b) => {
    for (const sk of sortKeys) {
      const va = sk.arr[a.idx]?.[0] ?? null;
      const vb = sk.arr[b.idx]?.[0] ?? null;
      let cmp: number;
      if (typeof va === "number" && typeof vb === "number") {
        cmp = va - vb;
      } else {
        cmp = toString(va).localeCompare(toString(vb));
      }
      if (cmp !== 0) {
        return cmp * sk.order;
      }
    }
    return 0;
  });
  return data.map(d => d.row);
};

// ============================================================================
// Additional Math Functions
// ============================================================================

const fnTRUNC: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const digits = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(digits)) {
    return digits;
  }
  const factor = Math.pow(10, digits);
  return Math.trunc(num * factor) / factor;
};

const fnSUMSQ: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  let sum = 0;
  for (const n of nums) {
    sum += (n as number) ** 2;
  }
  return sum;
};

const fnGCD: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return 0;
  }
  let result = Math.abs(Math.floor(nums[0] as number));
  for (let i = 1; i < nums.length; i++) {
    let b = Math.abs(Math.floor(nums[i] as number));
    while (b) {
      const t = b;
      b = result % b;
      result = t;
    }
  }
  return result;
};

const fnLCM: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return 0;
  }
  let result = Math.abs(Math.floor(nums[0] as number));
  for (let i = 1; i < nums.length; i++) {
    const b = Math.abs(Math.floor(nums[i] as number));
    if (result === 0 && b === 0) {
      result = 0;
    } else {
      let g = result;
      let t = b;
      while (t) {
        const tmp = t;
        t = g % t;
        g = tmp;
      }
      result = (result * b) / g;
    }
  }
  return result;
};

// ============================================================================
// Additional Text Functions
// ============================================================================

const fnCODE: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  return text.length > 0 ? text.charCodeAt(0) : ({ error: "#VALUE!" } as CellErrorValue);
};

const fnCHAR: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  return String.fromCharCode(n);
};

const fnCLEAN: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  // Remove non-printable ASCII control characters (0x00-0x1F)
  let result = "";
  for (let i = 0; i < text.length; i++) {
    if (text.charCodeAt(i) >= 32) {
      result += text[i];
    }
  }
  return result;
};

const fnT: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  return typeof v === "string" ? v : "";
};

// ============================================================================
// Additional Lookup Functions (XLOOKUP, XMATCH, ADDRESS, OFFSET)
// ============================================================================

const fnXLOOKUP: ExcelFunction = args => {
  const lookupValue = args[0] as CalcValue;
  if (!Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const lookupArr = args[1] as CalcArray;
  if (!Array.isArray(args[2])) {
    return { error: "#VALUE!" };
  }
  const returnArr = args[2] as CalcArray;
  const ifNotFound = args.length > 3 ? (args[3] as CalcValue) : null;
  const matchMode = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(matchMode)) {
    return matchMode;
  }
  // searchMode: 1=first-to-last, -1=last-to-first, 2=binary asc, -2=binary desc
  // For simplicity we implement linear search modes
  const searchMode = args.length > 5 ? toNumber(args[5] as CalcValue) : 1;
  if (isError(searchMode)) {
    return searchMode;
  }

  // Flatten lookup array to 1D
  const flat: CalcValue[] = [];
  const isRow = lookupArr.length === 1;
  if (isRow) {
    for (const cell of lookupArr[0]) {
      flat.push(cell);
    }
  } else {
    for (const row of lookupArr) {
      flat.push(row[0]);
    }
  }

  let foundIdx = -1;

  const doCompare = (a: CalcValue, b: CalcValue): number => {
    if (typeof a === "number" && typeof b === "number") {
      return a - b;
    }
    if (typeof a === "string" && typeof b === "string") {
      return a.toLowerCase().localeCompare(b.toLowerCase());
    }
    return 0;
  };

  if (matchMode === 0) {
    // Exact match
    const start = searchMode === -1 ? flat.length - 1 : 0;
    const end = searchMode === -1 ? -1 : flat.length;
    const step = searchMode === -1 ? -1 : 1;
    for (let i = start; i !== end; i += step) {
      if (flat[i] === lookupValue) {
        foundIdx = i;
        break;
      }
      if (
        typeof flat[i] === "string" &&
        typeof lookupValue === "string" &&
        (flat[i] as string).toLowerCase() === lookupValue.toLowerCase()
      ) {
        foundIdx = i;
        break;
      }
    }
  } else if (matchMode === -1) {
    // Exact match or next smaller
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      if (typeof flat[i] === typeof lookupValue) {
        const cmp = doCompare(flat[i], lookupValue);
        if (cmp === 0) {
          best = i;
          break;
        }
        if (cmp < 0 && (best === -1 || doCompare(flat[i], flat[best]) > 0)) {
          best = i;
        }
      }
    }
    foundIdx = best;
  } else if (matchMode === 1) {
    // Exact match or next larger
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      if (typeof flat[i] === typeof lookupValue) {
        const cmp = doCompare(flat[i], lookupValue);
        if (cmp === 0) {
          best = i;
          break;
        }
        if (cmp > 0 && (best === -1 || doCompare(flat[i], flat[best]) < 0)) {
          best = i;
        }
      }
    }
    foundIdx = best;
  } else if (matchMode === 2) {
    // Wildcard match
    const pattern = toString(lookupValue)
      .replace(/[.*+^${}()|[\]\\]/g, m => (m === "*" || m === "?" ? m : "\\" + m))
      .replace(/\*/g, ".*")
      .replace(/\?/g, ".");
    try {
      const re = new RegExp("^" + pattern + "$", "i");
      for (let i = 0; i < flat.length; i++) {
        if (re.test(toString(flat[i]))) {
          foundIdx = i;
          break;
        }
      }
    } catch {
      // Fallback: exact
      for (let i = 0; i < flat.length; i++) {
        if (toString(flat[i]).toLowerCase() === toString(lookupValue).toLowerCase()) {
          foundIdx = i;
          break;
        }
      }
    }
  }

  if (foundIdx === -1) {
    return ifNotFound !== null ? ifNotFound : ({ error: "#N/A" } as CellErrorValue);
  }

  // Return from return array
  if (isRow) {
    // Return array is also row-oriented
    if (returnArr.length === 1) {
      return returnArr[0][foundIdx] ?? null;
    }
    // Multiple rows in return — return column
    return returnArr[foundIdx]?.[0] ?? null;
  }
  // Column lookup — return from same row index
  const retRow = returnArr[foundIdx];
  return retRow ? (retRow.length === 1 ? retRow[0] : [retRow]) : null;
};

const fnXMATCH: ExcelFunction = args => {
  const lookupValue = args[0] as CalcValue;
  if (!Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const lookupArr = args[1] as CalcArray;
  const matchMode = args.length > 2 ? toNumber(args[2] as CalcValue) : 0;
  if (isError(matchMode)) {
    return matchMode;
  }
  const searchMode = args.length > 3 ? toNumber(args[3] as CalcValue) : 1;
  if (isError(searchMode)) {
    return searchMode;
  }

  const flat: CalcValue[] = [];
  if (lookupArr.length === 1) {
    for (const cell of lookupArr[0]) {
      flat.push(cell);
    }
  } else {
    for (const row of lookupArr) {
      flat.push(row[0]);
    }
  }

  if (matchMode === 0) {
    const start = searchMode === -1 ? flat.length - 1 : 0;
    const end = searchMode === -1 ? -1 : flat.length;
    const step = searchMode === -1 ? -1 : 1;
    for (let i = start; i !== end; i += step) {
      if (flat[i] === lookupValue) {
        return i + 1;
      }
      if (
        typeof flat[i] === "string" &&
        typeof lookupValue === "string" &&
        (flat[i] as string).toLowerCase() === lookupValue.toLowerCase()
      ) {
        return i + 1;
      }
    }
    return { error: "#N/A" } as CellErrorValue;
  }
  // For other match modes, delegate to simpler logic
  if (matchMode === -1) {
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      if (typeof flat[i] === "number" && typeof lookupValue === "number") {
        if ((flat[i] as number) <= (lookupValue as number)) {
          if (best === -1 || (flat[i] as number) > (flat[best] as number)) {
            best = i;
          }
        }
      }
    }
    return best >= 0 ? best + 1 : ({ error: "#N/A" } as CellErrorValue);
  }
  if (matchMode === 1) {
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      if (typeof flat[i] === "number" && typeof lookupValue === "number") {
        if ((flat[i] as number) >= (lookupValue as number)) {
          if (best === -1 || (flat[i] as number) < (flat[best] as number)) {
            best = i;
          }
        }
      }
    }
    return best >= 0 ? best + 1 : ({ error: "#N/A" } as CellErrorValue);
  }
  return { error: "#N/A" } as CellErrorValue;
};

const fnADDRESS: ExcelFunction = args => {
  const rowNum = toNumber(args[0] as CalcValue);
  if (isError(rowNum)) {
    return rowNum;
  }
  const colNum = toNumber(args[1] as CalcValue);
  if (isError(colNum)) {
    return colNum;
  }
  const absNum = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(absNum)) {
    return absNum;
  }
  // a1 style (true/default) vs r1c1 (false)
  const a1 = args.length > 3 ? (args[3] as CalcValue) : true;
  const sheetText = args.length > 4 ? toString(args[4] as CalcValue) : "";

  if (a1 === false) {
    // R1C1 style
    const rPart = absNum === 1 || absNum === 2 ? `R${rowNum}` : `R[${rowNum}]`;
    const cPart = absNum === 1 || absNum === 3 ? `C${colNum}` : `C[${colNum}]`;
    const prefix = sheetText ? `${sheetText}!` : "";
    return prefix + rPart + cPart;
  }

  // Convert column number to letters
  let col = "";
  let c = colNum;
  while (c > 0) {
    c--;
    col = String.fromCharCode(65 + (c % 26)) + col;
    c = Math.floor(c / 26);
  }

  let result = "";
  switch (absNum) {
    case 1:
      result = "$" + col + "$" + rowNum;
      break;
    case 2:
      result = col + "$" + rowNum;
      break;
    case 3:
      result = "$" + col + rowNum;
      break;
    case 4:
      result = col + rowNum;
      break;
    default:
      result = "$" + col + "$" + rowNum;
  }

  if (sheetText) {
    const needsQuote = /\s/.test(sheetText);
    result = (needsQuote ? `'${sheetText}'` : sheetText) + "!" + result;
  }
  return result;
};

const fnNUMBERVALUE: ExcelFunction = args => {
  let text = toString(args[0] as CalcValue);
  const decSep = args.length > 1 ? toString(args[1] as CalcValue) : ".";
  const grpSep = args.length > 2 ? toString(args[2] as CalcValue) : ",";
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
    return { error: "#VALUE!" } as CellErrorValue;
  }
  return isPct ? n / 100 : n;
};

const fnDATEVALUE: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);

  // Lotus 1-2-3 bug: "2/29/1900" or "February 29, 1900" etc. should return 60
  const lotus29 =
    /^(2[/-]29[/-]1900|1900[/-]2[/-]29|1900[/-]02[/-]29|02[/-]29[/-]1900|Feb(ruary)?\s+29[,]?\s+1900)$/i;
  if (lotus29.test(text.trim())) {
    return 60;
  }

  const d = new Date(text);
  if (isNaN(d.getTime())) {
    return { error: "#VALUE!" } as CellErrorValue;
  }
  return dateToExcel(new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())));
};

const fnTIMEVALUE: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const m = /(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(text);
  if (!m) {
    return { error: "#VALUE!" } as CellErrorValue;
  }
  const h = parseInt(m[1], 10);
  const min = parseInt(m[2], 10);
  const sec = m[3] ? parseInt(m[3], 10) : 0;
  return (h * 3600 + min * 60 + sec) / 86400;
};

const fnDAYS: ExcelFunction = args => {
  const end = toNumber(args[0] as CalcValue);
  if (isError(end)) {
    return end;
  }
  const start = toNumber(args[1] as CalcValue);
  if (isError(start)) {
    return start;
  }
  return Math.floor(end) - Math.floor(start);
};

const fnISOWEEKNUM: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  let d: Date;
  if (v instanceof Date) {
    d = v;
  } else {
    const n = toNumber(v);
    if (isError(n)) {
      return n;
    }
    d = excelToDate(n);
  }
  const temp = new Date(d.getTime());
  temp.setDate(temp.getDate() + 3 - ((temp.getDay() + 6) % 7));
  const week1 = new Date(temp.getFullYear(), 0, 4);
  return (
    1 +
    Math.round(((temp.getTime() - week1.getTime()) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7)
  );
};

const fnWEEKNUM: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  let d: Date;
  if (v instanceof Date) {
    d = v;
  } else {
    const n = toNumber(v);
    if (isError(n)) {
      return n;
    }
    d = excelToDate(n);
  }
  const returnType = args.length > 1 ? toNumber(args[1] as CalcValue) : 1;
  if (isError(returnType)) {
    return returnType;
  }
  // For returnType 21, use ISO week
  if (returnType === 21) {
    return fnISOWEEKNUM(args);
  }
  // Simple: week starts on Sunday (type 1) or Monday (type 2)
  const startDay = returnType === 2 ? 1 : 0; // 0=Sun, 1=Mon
  const jan1 = new Date(d.getFullYear(), 0, 1);
  const jan1Day = jan1.getDay();
  const dayOfYear = Math.floor((d.getTime() - jan1.getTime()) / 86400000);
  return Math.floor((dayOfYear + jan1Day - startDay + 7) / 7);
};

function networkdaysHelper(startN: number, endN: number, holidays: Set<number>): number {
  const s = Math.floor(Math.min(startN, endN));
  const e = Math.floor(Math.max(startN, endN));
  const sign = startN <= endN ? 1 : -1;
  let count = 0;
  for (let d = s; d <= e; d++) {
    const dt = excelToDate(d);
    const dow = dt.getDay();
    if (dow !== 0 && dow !== 6 && !holidays.has(d)) {
      count++;
    }
  }
  return count * sign;
}

function collectHolidays(arg: CalcValue | CalcArray): Set<number> {
  const set = new Set<number>();
  if (Array.isArray(arg)) {
    for (const row of arg) {
      for (const cell of row) {
        if (typeof cell === "number") {
          set.add(Math.floor(cell));
        } else if (cell instanceof Date) {
          set.add(Math.floor(dateToExcel(cell)));
        }
      }
    }
  } else if (typeof arg === "number") {
    set.add(Math.floor(arg));
  } else if (arg instanceof Date) {
    set.add(Math.floor(dateToExcel(arg)));
  }
  return set;
}

const fnNETWORKDAYS: ExcelFunction = args => {
  const startN = toNumber(args[0] as CalcValue);
  if (isError(startN)) {
    return startN;
  }
  const endN = toNumber(args[1] as CalcValue);
  if (isError(endN)) {
    return endN;
  }
  const holidays = args.length > 2 ? collectHolidays(args[2]) : new Set<number>();
  return networkdaysHelper(startN, endN, holidays);
};

const fnWORKDAY: ExcelFunction = args => {
  const startN = toNumber(args[0] as CalcValue);
  if (isError(startN)) {
    return startN;
  }
  const days = toNumber(args[1] as CalcValue);
  if (isError(days)) {
    return days;
  }
  const holidays = args.length > 2 ? collectHolidays(args[2]) : new Set<number>();
  let current = Math.floor(startN);
  const step = days >= 0 ? 1 : -1;
  let remaining = Math.abs(days);
  while (remaining > 0) {
    current += step;
    const dt = excelToDate(current);
    const dow = dt.getDay();
    if (dow !== 0 && dow !== 6 && !holidays.has(current)) {
      remaining--;
    }
  }
  return current;
};

const fnYEARFRAC: ExcelFunction = args => {
  const startN = toNumber(args[0] as CalcValue);
  if (isError(startN)) {
    return startN;
  }
  const endN = toNumber(args[1] as CalcValue);
  if (isError(endN)) {
    return endN;
  }
  const basis = args.length > 2 ? toNumber(args[2] as CalcValue) : 0;
  if (isError(basis)) {
    return basis;
  }
  const sd = excelToDate(Math.min(startN, endN));
  const ed = excelToDate(Math.max(startN, endN));
  const diffDays = Math.abs(Math.floor(endN) - Math.floor(startN));

  switch (basis) {
    case 0: {
      // US (NASD) 30/360
      let d1 = sd.getDate();
      const m1 = sd.getMonth() + 1;
      const y1 = sd.getFullYear();
      let d2 = ed.getDate();
      const m2 = ed.getMonth() + 1;
      const y2 = ed.getFullYear();
      // NASD adjustment rules
      if (d1 === 31) {
        d1 = 30;
      }
      if (d2 === 31 && d1 >= 30) {
        d2 = 30;
      }
      // Handle end-of-Feb for start date
      const feb1 = new Date(y1, 1, 29).getMonth() === 1 ? 29 : 28;
      if (m1 === 2 && d1 === feb1) {
        d1 = 30;
        if (m2 === 2) {
          const feb2 = new Date(y2, 1, 29).getMonth() === 1 ? 29 : 28;
          if (d2 === feb2) {
            d2 = 30;
          }
        }
      }
      const days30_360 = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
      return days30_360 / 360;
    }
    case 1: {
      // Actual/actual
      const y1 = sd.getFullYear();
      const y2 = ed.getFullYear();
      if (y1 === y2) {
        const yearDays =
          (new Date(y1 + 1, 0, 1).getTime() - new Date(y1, 0, 1).getTime()) / 86400000;
        return diffDays / yearDays;
      }
      // Span across years: average the year lengths
      const totalYearDays =
        (new Date(y2 + 1, 0, 1).getTime() - new Date(y1, 0, 1).getTime()) / 86400000;
      const avgYear = totalYearDays / (y2 - y1 + 1);
      return diffDays / avgYear;
    }
    case 2: // Actual/360
      return diffDays / 360;
    case 3: // Actual/365
      return diffDays / 365;
    case 4: {
      // European 30/360
      const d1 = Math.min(sd.getDate(), 30);
      const d2 = Math.min(ed.getDate(), 30);
      const m1 = sd.getMonth() + 1;
      const m2 = ed.getMonth() + 1;
      const y1 = sd.getFullYear();
      const y2 = ed.getFullYear();
      const days30_360 = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
      return days30_360 / 360;
    }
    default:
      return { error: "#NUM!" } as CellErrorValue;
  }
};

const fnSUBTOTAL: ExcelFunction = args => {
  const funcNum = toNumber(args[0] as CalcValue);
  if (isError(funcNum)) {
    return funcNum;
  }
  const dataArgs = args.slice(1);
  // funcNum 1-11 or 101-111 (101-111 ignores hidden rows — we treat the same since we don't track visibility)
  const fn = funcNum > 100 ? funcNum - 100 : funcNum;
  switch (fn) {
    case 1:
      return fnAVERAGE(dataArgs);
    case 2:
      return fnCOUNT(dataArgs);
    case 3:
      return fnCOUNTA(dataArgs);
    case 4:
      return fnMAX(dataArgs);
    case 5:
      return fnMIN(dataArgs);
    case 6:
      return fnPRODUCT(dataArgs);
    case 7:
      return fnSTDEV(dataArgs);
    case 8:
      return fnSTDEVP(dataArgs);
    case 9:
      return fnSUM(dataArgs);
    case 10:
      return fnVAR(dataArgs);
    case 11:
      return fnVARP(dataArgs);
    default:
      return { error: "#VALUE!" } as CellErrorValue;
  }
};

const fnAGGREGATE: ExcelFunction = args => {
  const funcNum = toNumber(args[0] as CalcValue);
  if (isError(funcNum)) {
    return funcNum;
  }
  // options arg ignored (error/hidden row handling) for simplicity
  const dataArgs = args.slice(2);
  switch (funcNum) {
    case 1:
      return fnAVERAGE(dataArgs);
    case 2:
      return fnCOUNT(dataArgs);
    case 3:
      return fnCOUNTA(dataArgs);
    case 4:
      return fnMAX(dataArgs);
    case 5:
      return fnMIN(dataArgs);
    case 6:
      return fnPRODUCT(dataArgs);
    case 7:
      return fnSTDEV(dataArgs);
    case 8:
      return fnSTDEVP(dataArgs);
    case 9:
      return fnSUM(dataArgs);
    case 10:
      return fnVAR(dataArgs);
    case 11:
      return fnVARP(dataArgs);
    case 12:
      return fnMEDIAN(dataArgs);
    case 14:
      return fnLARGE(dataArgs);
    case 15:
      return fnSMALL(dataArgs);
    default:
      return { error: "#VALUE!" } as CellErrorValue;
  }
};

// ============================================================================
// Financial Functions
// ============================================================================

const fnPMT: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumber(args[1] as CalcValue);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumber(args[2] as CalcValue);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 3 ? toNumber(args[3] as CalcValue) : 0;
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(type)) {
    return type;
  }
  if (rate === 0) {
    return -(pv + fv) / nper;
  }
  const pvif = Math.pow(1 + rate, nper);
  return -(rate * (pv * pvif + fv)) / (pvif - 1) / (1 + rate * type);
};

const fnFV: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumber(args[1] as CalcValue);
  if (isError(nper)) {
    return nper;
  }
  const pmt = toNumber(args[2] as CalcValue);
  if (isError(pmt)) {
    return pmt;
  }
  const pv = args.length > 3 ? toNumber(args[3] as CalcValue) : 0;
  if (isError(pv)) {
    return pv;
  }
  const type = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(type)) {
    return type;
  }
  if (rate === 0) {
    return -(pv + pmt * nper);
  }
  const pvif = Math.pow(1 + rate, nper);
  return -(pv * pvif + pmt * (1 + rate * type) * ((pvif - 1) / rate));
};

const fnPV: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumber(args[1] as CalcValue);
  if (isError(nper)) {
    return nper;
  }
  const pmt = toNumber(args[2] as CalcValue);
  if (isError(pmt)) {
    return pmt;
  }
  const fv = args.length > 3 ? toNumber(args[3] as CalcValue) : 0;
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(type)) {
    return type;
  }
  if (rate === 0) {
    return -pmt * nper - fv;
  }
  const pvif = Math.pow(1 + rate, nper);
  return -(fv + pmt * (1 + rate * type) * ((pvif - 1) / rate)) / pvif;
};

const fnNPV: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const values: number[] = [];
  for (let i = 1; i < args.length; i++) {
    const a = args[i];
    if (Array.isArray(a)) {
      for (const row of a) {
        for (const cell of row) {
          if (typeof cell === "number") {
            values.push(cell);
          }
        }
      }
    } else {
      const n = toNumber(a as CalcValue);
      if (isError(n)) {
        return n;
      }
      values.push(n);
    }
  }
  let npv = 0;
  for (let i = 0; i < values.length; i++) {
    npv += values[i] / Math.pow(1 + rate, i + 1);
  }
  return npv;
};

const fnIRR: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const values: number[] = [];
  for (const row of args[0] as CalcArray) {
    for (const cell of row) {
      if (typeof cell === "number") {
        values.push(cell);
      }
    }
  }
  if (values.length < 2) {
    return { error: "#NUM!" };
  }
  const guess = args.length > 1 ? toNumber(args[1] as CalcValue) : 0.1;
  if (isError(guess)) {
    return guess;
  }
  let g = guess as number;
  // Newton-Raphson
  for (let iter = 0; iter < 100; iter++) {
    let npv = 0;
    let dnpv = 0;
    for (let i = 0; i < values.length; i++) {
      npv += values[i] / Math.pow(1 + g, i);
      dnpv -= (i * values[i]) / Math.pow(1 + g, i + 1);
    }
    if (Math.abs(dnpv) < 1e-15) {
      break;
    }
    const newGuess = g - npv / dnpv;
    if (Math.abs(newGuess - g) < 1e-10) {
      return newGuess;
    }
    g = newGuess;
  }
  // Did not converge after 100 iterations
  return { error: "#NUM!" } as CellErrorValue;
};
const fnNPER: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const pmt = toNumber(args[1] as CalcValue);
  if (isError(pmt)) {
    return pmt;
  }
  const pv = toNumber(args[2] as CalcValue);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 3 ? toNumber(args[3] as CalcValue) : 0;
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(type)) {
    return type;
  }
  if (rate === 0) {
    return -(pv + fv) / pmt;
  }
  const num = pmt * (1 + rate * type) - fv * rate;
  const den = pv * rate + pmt * (1 + rate * type);
  if (num / den <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.log(num / den) / Math.log(1 + rate);
};

const fnRATE: ExcelFunction = args => {
  const nper = toNumber(args[0] as CalcValue);
  if (isError(nper)) {
    return nper;
  }
  const pmt = toNumber(args[1] as CalcValue);
  if (isError(pmt)) {
    return pmt;
  }
  const pv = toNumber(args[2] as CalcValue);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 3 ? toNumber(args[3] as CalcValue) : 0;
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(type)) {
    return type;
  }
  const guess = args.length > 5 ? toNumber(args[5] as CalcValue) : 0.1;
  if (isError(guess)) {
    return guess;
  }
  let g = guess as number;
  // Newton-Raphson
  for (let iter = 0; iter < 100; iter++) {
    if (g <= -1) {
      g = -0.99;
    }
    const pvif = Math.pow(1 + g, nper as number);
    const fvifa = (pvif - 1) / g;
    const f =
      (pv as number) * pvif + (pmt as number) * (1 + g * (type as number)) * fvifa + (fv as number);
    const df =
      (nper as number) * (pv as number) * Math.pow(1 + g, (nper as number) - 1) +
      ((pmt as number) *
        (1 + g * (type as number)) *
        ((nper as number) * Math.pow(1 + g, (nper as number) - 1) * g - pvif + 1)) /
        (g * g) +
      ((type as number) ? (pmt as number) * fvifa : 0);
    if (Math.abs(df) < 1e-15) {
      break;
    }
    const newGuess = g - f / df;
    if (Math.abs(newGuess - g) < 1e-10) {
      return newGuess;
    }
    g = newGuess;
  }
  // Did not converge after 100 iterations
  return { error: "#NUM!" } as CellErrorValue;
};

const fnSLN: ExcelFunction = args => {
  const cost = toNumber(args[0] as CalcValue);
  if (isError(cost)) {
    return cost;
  }
  const salvage = toNumber(args[1] as CalcValue);
  if (isError(salvage)) {
    return salvage;
  }
  const life = toNumber(args[2] as CalcValue);
  if (isError(life)) {
    return life;
  }
  if (life === 0) {
    return { error: "#DIV/0!" };
  }
  return (cost - salvage) / life;
};

const fnDB: ExcelFunction = args => {
  const cost = toNumber(args[0] as CalcValue);
  if (isError(cost)) {
    return cost;
  }
  const salvage = toNumber(args[1] as CalcValue);
  if (isError(salvage)) {
    return salvage;
  }
  const life = toNumber(args[2] as CalcValue);
  if (isError(life)) {
    return life;
  }
  const period = toNumber(args[3] as CalcValue);
  if (isError(period)) {
    return period;
  }
  const month = args.length > 4 ? toNumber(args[4] as CalcValue) : 12;
  if (isError(month)) {
    return month;
  }
  if (life === 0 || cost === 0) {
    return 0;
  }
  const rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000;
  let totalDepreciation = 0;
  let depn: number = 0;
  for (let p = 1; p <= period; p++) {
    if (p === 1) {
      depn = (cost * rate * month) / 12;
    } else if (p === Math.floor(life) + 1) {
      depn = ((cost - totalDepreciation) * rate * (12 - month)) / 12;
    } else {
      depn = (cost - totalDepreciation) * rate;
    }
    totalDepreciation += depn;
  }
  return depn;
};

const fnDDB: ExcelFunction = args => {
  const cost = toNumber(args[0] as CalcValue);
  if (isError(cost)) {
    return cost;
  }
  const salvage = toNumber(args[1] as CalcValue);
  if (isError(salvage)) {
    return salvage;
  }
  const life = toNumber(args[2] as CalcValue);
  if (isError(life)) {
    return life;
  }
  const period = toNumber(args[3] as CalcValue);
  if (isError(period)) {
    return period;
  }
  const factor = args.length > 4 ? toNumber(args[4] as CalcValue) : 2;
  if (isError(factor)) {
    return factor;
  }
  let bookValue = cost;
  let depn = 0;
  for (let p = 1; p <= period; p++) {
    depn = Math.min(bookValue * (factor / life), bookValue - salvage);
    if (depn < 0) {
      depn = 0;
    }
    bookValue -= depn;
  }
  return depn;
};

const fnIPMT: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const per = toNumber(args[1] as CalcValue);
  if (isError(per)) {
    return per;
  }
  const nper = toNumber(args[2] as CalcValue);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumber(args[3] as CalcValue);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 5 ? toNumber(args[5] as CalcValue) : 0;
  if (isError(type)) {
    return type;
  }
  const pmt = fnPMT([rate, nper, pv, fv, type]) as number;
  if (rate === 0) {
    return 0;
  }
  // Compute FV of original loan at period (per-1)
  let bal = pv;
  for (let i = 1; i < per; i++) {
    bal = bal * (1 + rate) + pmt * (1 + rate * type);
  }
  const ipmt = type === 1 && per === 1 ? 0 : bal * rate;
  return type === 1 ? ipmt / (1 + rate) : ipmt;
};

const fnPPMT: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const per = toNumber(args[1] as CalcValue);
  if (isError(per)) {
    return per;
  }
  const nper = toNumber(args[2] as CalcValue);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumber(args[3] as CalcValue);
  if (isError(pv)) {
    return pv;
  }
  const fv = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(fv)) {
    return fv;
  }
  const type = args.length > 5 ? toNumber(args[5] as CalcValue) : 0;
  if (isError(type)) {
    return type;
  }
  const pmt = fnPMT([rate, nper, pv, fv, type]) as number;
  const ipmt = fnIPMT([rate, per, nper, pv, fv, type]) as number;
  return pmt - ipmt;
};

// ============================================================================
// Engineering Functions
// ============================================================================

const fnBIN2DEC: ExcelFunction = args => {
  const s = toString(args[0] as CalcValue);
  if (!/^[01]{1,10}$/.test(s)) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  // 10-bit two's complement
  if (s.length === 10 && s[0] === "1") {
    return parseInt(s.slice(1), 2) - 512;
  }
  return parseInt(s, 2);
};

const fnDEC2BIN: ExcelFunction = args => {
  let n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  n = Math.floor(n);
  if (n < -512 || n > 511) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const places = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(places)) {
    return places;
  }
  if (n < 0) {
    return (n + 1024).toString(2);
  }
  const result = n.toString(2);
  return places > 0 ? result.padStart(places, "0") : result;
};

const fnHEX2DEC: ExcelFunction = args => {
  const s = toString(args[0] as CalcValue);
  if (!/^[0-9A-Fa-f]{1,10}$/.test(s)) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const num = parseInt(s, 16);
  // 10-digit hex: 40-bit two's complement
  if (s.length === 10 && parseInt(s[0], 16) >= 8) {
    return num - Math.pow(16, 10);
  }
  return num;
};

const fnDEC2HEX: ExcelFunction = args => {
  let n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  n = Math.floor(n);
  const places = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(places)) {
    return places;
  }
  if (n < 0) {
    return (n + Math.pow(16, 10)).toString(16).toUpperCase();
  }
  const result = n.toString(16).toUpperCase();
  return places > 0 ? result.padStart(places, "0") : result;
};

const fnOCT2DEC: ExcelFunction = args => {
  const s = toString(args[0] as CalcValue);
  if (!/^[0-7]{1,10}$/.test(s)) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const num = parseInt(s, 8);
  if (s.length === 10 && parseInt(s[0]) >= 4) {
    return num - Math.pow(8, 10);
  }
  return num;
};

const fnDEC2OCT: ExcelFunction = args => {
  let n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  n = Math.floor(n);
  const places = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(places)) {
    return places;
  }
  if (n < 0) {
    return (n + Math.pow(8, 10)).toString(8);
  }
  const result = n.toString(8);
  return places > 0 ? result.padStart(places, "0") : result;
};

const fnDELTA: ExcelFunction = args => {
  const n1 = toNumber(args[0] as CalcValue);
  if (isError(n1)) {
    return n1;
  }
  const n2 = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(n2)) {
    return n2;
  }
  return n1 === n2 ? 1 : 0;
};

const fnGESTEP: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  const step = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(step)) {
    return step;
  }
  return n >= step ? 1 : 0;
};

// ============================================================================
// Advanced Statistical Functions
// ============================================================================

// Peter Acklam's rational approximation for the standard normal inverse CDF.
// Accuracy: |error| < 1.15e-9 across the full range (0, 1).
function normSInv(p: number): number {
  if (p <= 0 || p >= 1) {
    return NaN;
  }
  if (p < 0.5) {
    return -normSInv(1 - p);
  }
  // Coefficients for rational approximation
  const a1 = -3.969683028665376e1;
  const a2 = 2.209460984245205e2;
  const a3 = -2.759285104469687e2;
  const a4 = 1.38357751867269e2;
  const a5 = -3.066479806614716e1;
  const a6 = 2.506628277459239;

  const b1 = -5.447609879822406e1;
  const b2 = 1.615858368580409e2;
  const b3 = -1.556989798598866e2;
  const b4 = 6.680131188771972e1;
  const b5 = -1.328068155288572e1;

  const c1 = -7.784894002430293e-3;
  const c2 = -3.223964580411365e-1;
  const c3 = -2.400758277161838;
  const c4 = -2.549732539343734;
  const c5 = 4.374664141464968;
  const c6 = 2.938163982698783;

  const d1 = 7.784695709041462e-3;
  const d2 = 3.224671290700398e-1;
  const d3 = 2.445134137142996;
  const d4 = 3.754408661907416;

  const pLow = 0.02425;
  const pHigh = 1 - pLow;

  if (p < pLow) {
    // Rational approximation for lower region
    const q = Math.sqrt(-2 * Math.log(p));
    return (
      (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
      ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
    );
  }
  if (p <= pHigh) {
    // Rational approximation for central region
    const q = p - 0.5;
    const r = q * q;
    return (
      ((((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q) /
      (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)
    );
  }
  // Upper region — use symmetry
  const q = Math.sqrt(-2 * Math.log(1 - p));
  return -(
    (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
    ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
  );
}

// Standard normal CDF approximation (Abramowitz & Stegun)
function normSDist(x: number): number {
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  const sign = x < 0 ? -1 : 1;
  x = Math.abs(x) / Math.SQRT2;
  const t = 1.0 / (1.0 + p * x);
  const y = 1.0 - ((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
  return 0.5 * (1.0 + sign * y);
}

// Standard normal PDF
function normSPdf(x: number): number {
  return Math.exp(-0.5 * x * x) / Math.sqrt(2 * Math.PI);
}

const fnNORMSDIST: ExcelFunction = args => {
  const z = toNumber(args[0] as CalcValue);
  if (isError(z)) {
    return z;
  }
  // Legacy NORM.S.DIST compatibility: single arg = CDF
  const cumulative = args.length > 1 ? toBoolean(args[1] as CalcValue) : true;
  if (isError(cumulative)) {
    return cumulative;
  }
  return cumulative ? normSDist(z) : normSPdf(z);
};

const fnNORMDIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const mean = toNumber(args[1] as CalcValue);
  if (isError(mean)) {
    return mean;
  }
  const stddev = toNumber(args[2] as CalcValue);
  if (isError(stddev)) {
    return stddev;
  }
  if (stddev <= 0) {
    return { error: "#NUM!" };
  }
  const cumulative = toBoolean(args[3] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  const z = (x - mean) / stddev;
  if (cumulative) {
    return normSDist(z);
  }
  return normSPdf(z) / stddev;
};

const fnNORMSINV: ExcelFunction = args => {
  const p = toNumber(args[0] as CalcValue);
  if (isError(p)) {
    return p;
  }
  if (p <= 0 || p >= 1) {
    return { error: "#NUM!" };
  }
  return normSInv(p);
};

const fnNORMINV: ExcelFunction = args => {
  const p = toNumber(args[0] as CalcValue);
  if (isError(p)) {
    return p;
  }
  const mean = toNumber(args[1] as CalcValue);
  if (isError(mean)) {
    return mean;
  }
  const stddev = toNumber(args[2] as CalcValue);
  if (isError(stddev)) {
    return stddev;
  }
  if (p <= 0 || p >= 1 || stddev <= 0) {
    return { error: "#NUM!" };
  }
  return mean + stddev * normSInv(p);
};

const fnPERCENTILE: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const nums = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const k = toNumber(args[1] as CalcValue);
  if (isError(k)) {
    return k;
  }
  if (k < 0 || k > 1 || nums.length === 0) {
    return { error: "#NUM!" };
  }
  nums.sort((a, b) => a - b);
  const n = nums.length;
  if (n === 1) {
    return nums[0];
  }
  const rank = k * (n - 1);
  const lower = Math.floor(rank);
  const upper = Math.ceil(rank);
  const frac = rank - lower;
  return nums[lower] + frac * (nums[upper] - nums[lower]);
};

const fnPERCENTILEEXC: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const nums = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const k = toNumber(args[1] as CalcValue);
  if (isError(k)) {
    return k;
  }
  const n = nums.length;
  if (k <= 0 || k >= 1 || n === 0) {
    return { error: "#NUM!" };
  }
  if (k < 1 / (n + 1) || k > n / (n + 1)) {
    return { error: "#NUM!" };
  }
  nums.sort((a, b) => a - b);
  const rank = k * (n + 1) - 1;
  const lower = Math.floor(rank);
  const upper = Math.ceil(rank);
  const frac = rank - lower;
  return (
    nums[Math.max(0, lower)] + frac * (nums[Math.min(n - 1, upper)] - nums[Math.max(0, lower)])
  );
};

const fnQUARTILE: ExcelFunction = args => {
  const quart = toNumber(args[1] as CalcValue);
  if (isError(quart)) {
    return quart;
  }
  if (quart < 0 || quart > 4) {
    return { error: "#NUM!" };
  }
  return fnPERCENTILE([args[0], quart / 4]);
};

const fnQUARTILEEXC: ExcelFunction = args => {
  const quart = toNumber(args[1] as CalcValue);
  if (isError(quart)) {
    return quart;
  }
  if (quart < 1 || quart > 3) {
    return { error: "#NUM!" };
  }
  return fnPERCENTILEEXC([args[0], quart / 4]);
};

const fnMODE: ExcelFunction = args => {
  const nums = flattenNumbers(args).filter((v): v is number => !isError(v));
  if (nums.length === 0) {
    return { error: "#N/A" };
  }
  const counts = new Map<number, number>();
  let maxCount = 0;
  let mode = nums[0];
  for (const n of nums) {
    const c = (counts.get(n) ?? 0) + 1;
    counts.set(n, c);
    if (c > maxCount) {
      maxCount = c;
      mode = n;
    }
  }
  return maxCount > 1 ? mode : ({ error: "#N/A" } as CellErrorValue);
};

const fnCORREL: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const xs = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const ys = flattenNumbers([args[1]]).filter((v): v is number => !isError(v));
  const n = Math.min(xs.length, ys.length);
  if (n < 2) {
    return { error: "#DIV/0!" };
  }
  let sumX = 0;
  let sumY = 0;
  for (let i = 0; i < n; i++) {
    sumX += xs[i];
    sumY += ys[i];
  }
  const meanX = sumX / n;
  const meanY = sumY / n;
  let num = 0;
  let denomX = 0;
  let denomY = 0;
  for (let i = 0; i < n; i++) {
    const dx = xs[i] - meanX;
    const dy = ys[i] - meanY;
    num += dx * dy;
    denomX += dx * dx;
    denomY += dy * dy;
  }
  const denom = Math.sqrt(denomX * denomY);
  return denom === 0 ? ({ error: "#DIV/0!" } as CellErrorValue) : num / denom;
};

const fnSLOPE: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const ys = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const xs = flattenNumbers([args[1]]).filter((v): v is number => !isError(v));
  const n = Math.min(xs.length, ys.length);
  if (n < 2) {
    return { error: "#DIV/0!" };
  }
  let sumX = 0;
  let sumY = 0;
  for (let i = 0; i < n; i++) {
    sumX += xs[i];
    sumY += ys[i];
  }
  const meanX = sumX / n;
  const meanY = sumY / n;
  let num = 0;
  let denom = 0;
  for (let i = 0; i < n; i++) {
    num += (xs[i] - meanX) * (ys[i] - meanY);
    denom += (xs[i] - meanX) ** 2;
  }
  return denom === 0 ? ({ error: "#DIV/0!" } as CellErrorValue) : num / denom;
};

const fnINTERCEPT: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const ys = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const xs = flattenNumbers([args[1]]).filter((v): v is number => !isError(v));
  const n = Math.min(xs.length, ys.length);
  if (n < 2) {
    return { error: "#DIV/0!" };
  }
  let sumX = 0;
  let sumY = 0;
  for (let i = 0; i < n; i++) {
    sumX += xs[i];
    sumY += ys[i];
  }
  const meanX = sumX / n;
  const meanY = sumY / n;
  const slope = fnSLOPE(args) as number;
  if (isError(slope)) {
    return slope;
  }
  return meanY - slope * meanX;
};

const fnRSQ: ExcelFunction = args => {
  const r = fnCORREL(args);
  if (isError(r as CalcValue)) {
    return r;
  }
  return (r as number) ** 2;
};

const fnFORECAST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const slope = fnSLOPE([args[1], args[2]]);
  if (isError(slope as CalcValue)) {
    return slope;
  }
  const intercept = fnINTERCEPT([args[1], args[2]]);
  if (isError(intercept as CalcValue)) {
    return intercept;
  }
  return (intercept as number) + (slope as number) * x;
};

const fnFACT: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  const num = Math.floor(n);
  if (num < 0) {
    return { error: "#NUM!" };
  }
  if (num > 170) {
    return { error: "#NUM!" };
  }
  let result = 1;
  for (let i = 2; i <= num; i++) {
    result *= i;
  }
  return result;
};

const fnFACTDOUBLE: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  const num = Math.floor(n);
  if (num < -1) {
    return { error: "#NUM!" };
  }
  if (num <= 0) {
    return 1;
  }
  let result = 1;
  for (let i = num; i > 0; i -= 2) {
    result *= i;
  }
  return result;
};

const fnCOMBIN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  const k = toNumber(args[1] as CalcValue);
  if (isError(k)) {
    return k;
  }
  const ni = Math.floor(n);
  const ki = Math.floor(k);
  if (ni < 0 || ki < 0 || ki > ni) {
    return { error: "#NUM!" };
  }
  let result = 1;
  for (let i = 0; i < ki; i++) {
    result = (result * (ni - i)) / (i + 1);
  }
  return Math.round(result);
};

const fnCOMBINA: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  const k = toNumber(args[1] as CalcValue);
  if (isError(k)) {
    return k;
  }
  return fnCOMBIN([n + k - 1, k]);
};

const fnPERMUT: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  const k = toNumber(args[1] as CalcValue);
  if (isError(k)) {
    return k;
  }
  const ni = Math.floor(n);
  const ki = Math.floor(k);
  if (ni < 0 || ki < 0 || ki > ni) {
    return { error: "#NUM!" };
  }
  let result = 1;
  for (let i = 0; i < ki; i++) {
    result *= ni - i;
  }
  return result;
};

const fnGEOMEAN: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return { error: "#NUM!" };
  }
  let logSum = 0;
  for (const n of nums) {
    if ((n as number) <= 0) {
      return { error: "#NUM!" };
    }
    logSum += Math.log(n as number);
  }
  return Math.exp(logSum / nums.length);
};

const fnHARMEAN: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return { error: "#NUM!" };
  }
  let recipSum = 0;
  for (const n of nums) {
    if ((n as number) <= 0) {
      return { error: "#NUM!" };
    }
    recipSum += 1 / (n as number);
  }
  return nums.length / recipSum;
};

const fnTRIMMEAN: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const nums = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const pct = toNumber(args[1] as CalcValue);
  if (isError(pct)) {
    return pct;
  }
  if (pct < 0 || pct >= 1) {
    return { error: "#NUM!" };
  }
  nums.sort((a, b) => a - b);
  const trimCount = Math.floor((nums.length * pct) / 2);
  const trimmed = nums.slice(trimCount, nums.length - trimCount);
  if (trimmed.length === 0) {
    return { error: "#DIV/0!" };
  }
  return trimmed.reduce((a, b) => a + b, 0) / trimmed.length;
};

const fnDEVSQ: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return 0;
  }
  let sum = 0;
  for (const n of nums) {
    sum += n as number;
  }
  const mean = sum / nums.length;
  let result = 0;
  for (const n of nums) {
    result += ((n as number) - mean) ** 2;
  }
  return result;
};

const fnAVEDEV: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return { error: "#NUM!" };
  }
  let sum = 0;
  for (const n of nums) {
    sum += n as number;
  }
  const mean = sum / nums.length;
  let result = 0;
  for (const n of nums) {
    result += Math.abs((n as number) - mean);
  }
  return result / nums.length;
};

const fnCONFIDENCENORM: ExcelFunction = args => {
  const alpha = toNumber(args[0] as CalcValue);
  if (isError(alpha)) {
    return alpha;
  }
  const stddev = toNumber(args[1] as CalcValue);
  if (isError(stddev)) {
    return stddev;
  }
  const size = toNumber(args[2] as CalcValue);
  if (isError(size)) {
    return size;
  }
  if (alpha <= 0 || alpha >= 1 || stddev <= 0 || size < 1) {
    return { error: "#NUM!" };
  }
  return (normSInv(1 - alpha / 2) * stddev) / Math.sqrt(size);
};

const fnFISHER: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  if (x <= -1 || x >= 1) {
    return { error: "#NUM!" };
  }
  return 0.5 * Math.log((1 + x) / (1 - x));
};

const fnFISHERINV: ExcelFunction = args => {
  const y = toNumber(args[0] as CalcValue);
  if (isError(y)) {
    return y;
  }
  const e2y = Math.exp(2 * y);
  return (e2y - 1) / (e2y + 1);
};

const fnAVERAGEA: ExcelFunction = args => {
  const all = flattenAll(args);
  if (all.length === 0) {
    return { error: "#DIV/0!" };
  }
  let sum = 0;
  let count = 0;
  for (const v of all) {
    if (v === null) {
      continue;
    }
    if (isError(v)) {
      return v;
    }
    if (typeof v === "number") {
      sum += v;
    } else if (typeof v === "boolean") {
      sum += v ? 1 : 0;
    } else if (typeof v === "string") {
      // Text = 0 for AVERAGEA
    }
    count++;
  }
  return count === 0 ? ({ error: "#DIV/0!" } as CellErrorValue) : sum / count;
};

const fnMAXA: ExcelFunction = args => {
  const all = flattenAll(args);
  let max = -Infinity;
  let found = false;
  for (const v of all) {
    if (v === null) {
      continue;
    }
    if (isError(v)) {
      return v;
    }
    let n: number;
    if (typeof v === "number") {
      n = v;
    } else if (typeof v === "boolean") {
      n = v ? 1 : 0;
    } else {
      n = 0;
    }
    if (n > max) {
      max = n;
    }
    found = true;
  }
  return found ? max : 0;
};

const fnMINA: ExcelFunction = args => {
  const all = flattenAll(args);
  let min = Infinity;
  let found = false;
  for (const v of all) {
    if (v === null) {
      continue;
    }
    if (isError(v)) {
      return v;
    }
    let n: number;
    if (typeof v === "number") {
      n = v;
    } else if (typeof v === "boolean") {
      n = v ? 1 : 0;
    } else {
      n = 0;
    }
    if (n < min) {
      min = n;
    }
    found = true;
  }
  return found ? min : 0;
};

// ============================================================================
// Database Functions
// ============================================================================

function databaseHelper(
  args: (CalcValue | CalcArray)[],
  aggregator: (values: number[]) => CalcValue
): CalcValue {
  if (!Array.isArray(args[0]) || !Array.isArray(args[2])) {
    return { error: "#VALUE!" };
  }
  const database = args[0] as CalcArray;
  const fieldArg = args[1] as CalcValue;
  const criteria = args[2] as CalcArray;

  if (database.length < 2 || criteria.length < 2) {
    return { error: "#VALUE!" };
  }

  // Determine field column index
  const headers = database[0];
  let fieldIdx = -1;
  if (typeof fieldArg === "number") {
    fieldIdx = fieldArg - 1;
  } else if (typeof fieldArg === "string") {
    for (let c = 0; c < headers.length; c++) {
      if (toString(headers[c]).toLowerCase() === fieldArg.toLowerCase()) {
        fieldIdx = c;
        break;
      }
    }
  }
  if (fieldIdx < 0 || fieldIdx >= headers.length) {
    return { error: "#VALUE!" };
  }

  // Parse criteria: each criteria row is an OR condition, columns within a row are AND
  const critHeaders = criteria[0];
  const critColIndices: number[] = [];
  for (const ch of critHeaders) {
    const name = toString(ch).toLowerCase();
    const idx = headers.findIndex(h => toString(h).toLowerCase() === name);
    critColIndices.push(idx);
  }

  // Collect matching rows
  const values: number[] = [];
  for (let r = 1; r < database.length; r++) {
    let matchesAnyCritRow = false;
    for (let cr = 1; cr < criteria.length; cr++) {
      let allMatch = true;
      for (let cc = 0; cc < critHeaders.length; cc++) {
        const critVal = criteria[cr][cc];
        if (critVal === null || critVal === "") {
          continue;
        }
        const dbCol = critColIndices[cc];
        if (dbCol < 0) {
          allMatch = false;
          break;
        }
        const pred = buildCriteriaPredicate(critVal);
        if (!pred(database[r][dbCol])) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        matchesAnyCritRow = true;
        break;
      }
    }
    if (matchesAnyCritRow) {
      const v = database[r][fieldIdx];
      if (typeof v === "number") {
        values.push(v);
      }
    }
  }

  return aggregator(values);
}

const fnDSUM: ExcelFunction = args => databaseHelper(args, vals => vals.reduce((a, b) => a + b, 0));
const fnDAVERAGE: ExcelFunction = args =>
  databaseHelper(args, vals =>
    vals.length === 0
      ? ({ error: "#DIV/0!" } as CellErrorValue as unknown as number)
      : vals.reduce((a, b) => a + b, 0) / vals.length
  );
const fnDCOUNT: ExcelFunction = args => databaseHelper(args, vals => vals.length);
const fnDMAX: ExcelFunction = args =>
  databaseHelper(args, vals => (vals.length === 0 ? 0 : Math.max(...vals)));
const fnDMIN: ExcelFunction = args =>
  databaseHelper(args, vals => (vals.length === 0 ? 0 : Math.min(...vals)));
const fnDPRODUCT: ExcelFunction = args =>
  databaseHelper(args, vals => (vals.length === 0 ? 0 : vals.reduce((a, b) => a * b, 1)));

const fnDGET: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[2])) {
    return { error: "#VALUE!" };
  }
  const database = args[0] as CalcArray;
  const fieldArg = args[1] as CalcValue;
  const criteria = args[2] as CalcArray;

  if (database.length < 2 || criteria.length < 2) {
    return { error: "#VALUE!" };
  }

  const headers = database[0];
  let fieldIdx = -1;
  if (typeof fieldArg === "number") {
    fieldIdx = fieldArg - 1;
  } else if (typeof fieldArg === "string") {
    for (let c = 0; c < headers.length; c++) {
      if (toString(headers[c]).toLowerCase() === fieldArg.toLowerCase()) {
        fieldIdx = c;
        break;
      }
    }
  }
  if (fieldIdx < 0 || fieldIdx >= headers.length) {
    return { error: "#VALUE!" };
  }

  const critHeaders = criteria[0];
  const critColIndices: number[] = [];
  for (const ch of critHeaders) {
    const name = toString(ch).toLowerCase();
    const idx = headers.findIndex(h => toString(h).toLowerCase() === name);
    critColIndices.push(idx);
  }

  let found: CalcValue = null;
  let count = 0;
  for (let r = 1; r < database.length; r++) {
    let matchesAnyCritRow = false;
    for (let cr = 1; cr < criteria.length; cr++) {
      let allMatch = true;
      for (let cc = 0; cc < critHeaders.length; cc++) {
        const critVal = criteria[cr][cc];
        if (critVal === null || critVal === "") {
          continue;
        }
        const dbCol = critColIndices[cc];
        if (dbCol < 0) {
          allMatch = false;
          break;
        }
        const pred = buildCriteriaPredicate(critVal);
        if (!pred(database[r][dbCol])) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        matchesAnyCritRow = true;
        break;
      }
    }
    if (matchesAnyCritRow) {
      found = database[r][fieldIdx];
      count++;
      if (count > 1) {
        return { error: "#NUM!" };
      }
    }
  }
  return count === 0 ? ({ error: "#VALUE!" } as CellErrorValue) : found;
};

// ============================================================================
// Dynamic Array Helper Functions (SEQUENCE, RANDARRAY, etc.)
// ============================================================================

const fnSEQUENCE: ExcelFunction = args => {
  const rows = toNumber(args[0] as CalcValue);
  if (isError(rows)) {
    return rows;
  }
  const cols = args.length > 1 ? toNumber(args[1] as CalcValue) : 1;
  if (isError(cols)) {
    return cols;
  }
  const start = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(start)) {
    return start;
  }
  const step = args.length > 3 ? toNumber(args[3] as CalcValue) : 1;
  if (isError(step)) {
    return step;
  }
  const result: CalcArray = [];
  let val = start;
  for (let r = 0; r < rows; r++) {
    const row: CalcValue[] = [];
    for (let c = 0; c < cols; c++) {
      row.push(val);
      val += step;
    }
    result.push(row);
  }
  return result;
};

const fnRANDARRAY: ExcelFunction = args => {
  const rows = args.length > 0 ? toNumber(args[0] as CalcValue) : 1;
  if (isError(rows)) {
    return rows;
  }
  const cols = args.length > 1 ? toNumber(args[1] as CalcValue) : 1;
  if (isError(cols)) {
    return cols;
  }
  const min = args.length > 2 ? toNumber(args[2] as CalcValue) : 0;
  if (isError(min)) {
    return min;
  }
  const max = args.length > 3 ? toNumber(args[3] as CalcValue) : 1;
  if (isError(max)) {
    return max;
  }
  const wholeNumber = args.length > 4 ? toBoolean(args[4] as CalcValue) : false;
  if (isError(wholeNumber)) {
    return wholeNumber;
  }
  const result: CalcArray = [];
  for (let r = 0; r < rows; r++) {
    const row: CalcValue[] = [];
    for (let c = 0; c < cols; c++) {
      const v = min + Math.random() * (max - min);
      row.push(wholeNumber ? Math.floor(v) : v);
    }
    result.push(row);
  }
  return result;
};

const fnTOCOL: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return [[args[0] as CalcValue]];
  }
  const arr = args[0] as CalcArray;
  const ignore = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(ignore)) {
    return ignore;
  }
  const scanByCol = args.length > 2 ? toBoolean(args[2] as CalcValue) : false;
  if (isError(scanByCol)) {
    return scanByCol;
  }
  const result: CalcValue[][] = [];
  if (scanByCol) {
    const cols = arr[0]?.length ?? 0;
    for (let c = 0; c < cols; c++) {
      for (const row of arr) {
        const v = row[c] ?? null;
        if (ignore === 1 && (v === null || v === "")) {
          continue;
        }
        if (ignore === 2 && isError(v)) {
          continue;
        }
        if (ignore === 3 && (v === null || v === "" || isError(v))) {
          continue;
        }
        result.push([v]);
      }
    }
  } else {
    for (const row of arr) {
      for (const v of row) {
        if (ignore === 1 && (v === null || v === "")) {
          continue;
        }
        if (ignore === 2 && isError(v)) {
          continue;
        }
        if (ignore === 3 && (v === null || v === "" || isError(v))) {
          continue;
        }
        result.push([v]);
      }
    }
  }
  return result.length > 0 ? result : ({ error: "#CALC!" } as CellErrorValue);
};

const fnTOROW: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return [[args[0] as CalcValue]];
  }
  const arr = args[0] as CalcArray;
  const ignore = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(ignore)) {
    return ignore;
  }
  const scanByCol = args.length > 2 ? toBoolean(args[2] as CalcValue) : false;
  if (isError(scanByCol)) {
    return scanByCol;
  }
  const result: CalcValue[] = [];
  if (scanByCol) {
    const cols = arr[0]?.length ?? 0;
    for (let c = 0; c < cols; c++) {
      for (const row of arr) {
        const v = row[c] ?? null;
        if (ignore === 1 && (v === null || v === "")) {
          continue;
        }
        if (ignore === 2 && isError(v)) {
          continue;
        }
        if (ignore === 3 && (v === null || v === "" || isError(v))) {
          continue;
        }
        result.push(v);
      }
    }
  } else {
    for (const row of arr) {
      for (const v of row) {
        if (ignore === 1 && (v === null || v === "")) {
          continue;
        }
        if (ignore === 2 && isError(v)) {
          continue;
        }
        if (ignore === 3 && (v === null || v === "" || isError(v))) {
          continue;
        }
        result.push(v);
      }
    }
  }
  return result.length > 0 ? [result] : ({ error: "#CALC!" } as CellErrorValue);
};

const fnCHOOSEROWS: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const arr = args[0] as CalcArray;
  const result: CalcValue[][] = [];
  for (let i = 1; i < args.length; i++) {
    const n = toNumber(args[i] as CalcValue);
    if (isError(n)) {
      return n;
    }
    const idx = n > 0 ? n - 1 : arr.length + n;
    if (idx < 0 || idx >= arr.length) {
      return { error: "#VALUE!" };
    }
    result.push(arr[idx]);
  }
  return result;
};

const fnCHOOSECOLS: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const arr = args[0] as CalcArray;
  const cols = arr[0]?.length ?? 0;
  const colIndices: number[] = [];
  for (let i = 1; i < args.length; i++) {
    const n = toNumber(args[i] as CalcValue);
    if (isError(n)) {
      return n;
    }
    const idx = n > 0 ? n - 1 : cols + n;
    if (idx < 0 || idx >= cols) {
      return { error: "#VALUE!" };
    }
    colIndices.push(idx);
  }
  return arr.map(row => colIndices.map(c => row[c]));
};

const fnVSTACK: ExcelFunction = args => {
  const result: CalcValue[][] = [];
  for (const a of args) {
    if (Array.isArray(a)) {
      for (const row of a) {
        result.push([...row]);
      }
    } else {
      result.push([a as CalcValue]);
    }
  }
  return result.length > 0 ? result : ({ error: "#VALUE!" } as CellErrorValue);
};

const fnHSTACK: ExcelFunction = args => {
  // Determine max rows
  let maxRows = 0;
  const arrays: CalcArray[] = [];
  for (const a of args) {
    if (Array.isArray(a)) {
      arrays.push(a);
      if (a.length > maxRows) {
        maxRows = a.length;
      }
    } else {
      arrays.push([[a as CalcValue]]);
      if (maxRows < 1) {
        maxRows = 1;
      }
    }
  }
  const result: CalcValue[][] = [];
  for (let r = 0; r < maxRows; r++) {
    const row: CalcValue[] = [];
    for (const arr of arrays) {
      const srcRow = arr[r] ?? [];
      for (const v of srcRow) {
        row.push(v);
      }
      // Pad if this array has fewer columns
      if (!arr[r] && arr[0]) {
        for (let c = 0; c < arr[0].length; c++) {
          row.push({ error: "#N/A" } as CellErrorValue);
        }
      }
    }
    result.push(row);
  }
  return result;
};

const fnTAKE: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const arr = args[0] as CalcArray;
  const rows = args.length > 1 ? toNumber(args[1] as CalcValue) : arr.length;
  if (isError(rows)) {
    return rows;
  }
  const cols = args.length > 2 ? toNumber(args[2] as CalcValue) : (arr[0]?.length ?? 0);
  if (isError(cols)) {
    return cols;
  }
  const rStart = rows >= 0 ? 0 : Math.max(0, arr.length + rows);
  const rEnd = rows >= 0 ? Math.min(rows, arr.length) : arr.length;
  const cStart = cols >= 0 ? 0 : Math.max(0, (arr[0]?.length ?? 0) + cols);
  const cEnd = cols >= 0 ? Math.min(cols, arr[0]?.length ?? 0) : (arr[0]?.length ?? 0);
  const result: CalcValue[][] = [];
  for (let r = rStart; r < rEnd; r++) {
    result.push(arr[r].slice(cStart, cEnd));
  }
  return result.length > 0 ? result : ({ error: "#CALC!" } as CellErrorValue);
};

const fnDROP: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const arr = args[0] as CalcArray;
  const rows = args.length > 1 ? toNumber(args[1] as CalcValue) : 0;
  if (isError(rows)) {
    return rows;
  }
  const cols = args.length > 2 ? toNumber(args[2] as CalcValue) : 0;
  if (isError(cols)) {
    return cols;
  }
  const rStart = rows >= 0 ? rows : 0;
  const rEnd = rows >= 0 ? arr.length : arr.length + rows;
  const totalCols = arr[0]?.length ?? 0;
  const cStart = cols >= 0 ? cols : 0;
  const cEnd = cols >= 0 ? totalCols : totalCols + cols;
  const result: CalcValue[][] = [];
  for (let r = rStart; r < rEnd; r++) {
    if (arr[r]) {
      result.push(arr[r].slice(cStart, cEnd));
    }
  }
  return result.length > 0 ? result : ({ error: "#CALC!" } as CellErrorValue);
};

const fnWRAPROWS: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const flat: CalcValue[] = [];
  for (const row of args[0] as CalcArray) {
    for (const v of row) {
      flat.push(v);
    }
  }
  const wrapCount = toNumber(args[1] as CalcValue);
  if (isError(wrapCount)) {
    return wrapCount;
  }
  if (wrapCount < 1) {
    return { error: "#VALUE!" };
  }
  const padWith = args.length > 2 ? (args[2] as CalcValue) : ({ error: "#N/A" } as CellErrorValue);
  const result: CalcValue[][] = [];
  for (let i = 0; i < flat.length; i += wrapCount) {
    const row = flat.slice(i, i + wrapCount);
    while (row.length < wrapCount) {
      row.push(padWith);
    }
    result.push(row);
  }
  return result;
};

const fnWRAPCOLS: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const flat: CalcValue[] = [];
  for (const row of args[0] as CalcArray) {
    for (const v of row) {
      flat.push(v);
    }
  }
  const wrapCount = toNumber(args[1] as CalcValue);
  if (isError(wrapCount)) {
    return wrapCount;
  }
  if (wrapCount < 1) {
    return { error: "#VALUE!" };
  }
  const padWith = args.length > 2 ? (args[2] as CalcValue) : ({ error: "#N/A" } as CellErrorValue);
  const numCols = Math.ceil(flat.length / wrapCount);
  const result: CalcValue[][] = [];
  for (let r = 0; r < wrapCount; r++) {
    const row: CalcValue[] = [];
    for (let c = 0; c < numCols; c++) {
      const idx = c * wrapCount + r;
      row.push(idx < flat.length ? flat[idx] : padWith);
    }
    result.push(row);
  }
  return result;
};

const fnEXPAND: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const arr = args[0] as CalcArray;
  const rows = args.length > 1 ? toNumber(args[1] as CalcValue) : arr.length;
  if (isError(rows)) {
    return rows;
  }
  const cols = args.length > 2 ? toNumber(args[2] as CalcValue) : (arr[0]?.length ?? 0);
  if (isError(cols)) {
    return cols;
  }
  const padWith = args.length > 3 ? (args[3] as CalcValue) : ({ error: "#N/A" } as CellErrorValue);
  const result: CalcValue[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: CalcValue[] = [];
    for (let c = 0; c < cols; c++) {
      row.push(r < arr.length && c < (arr[r]?.length ?? 0) ? arr[r][c] : padWith);
    }
    result.push(row);
  }
  return result;
};

// ============================================================================
// Trigonometric Functions
// ============================================================================

const fnSIN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.sin(n);
};

const fnCOS: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.cos(n);
};

const fnTAN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.tan(n);
};

const fnASIN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n < -1 || n > 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.asin(n);
};

const fnACOS: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n < -1 || n > 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.acos(n);
};

const fnATAN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.atan(n);
};

const fnATAN2: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const y = toNumber(args[1] as CalcValue);
  if (isError(y)) {
    return y;
  }
  if (x === 0 && y === 0) {
    return { error: "#DIV/0!" } as CellErrorValue;
  }
  return Math.atan2(y, x);
};

const fnSINH: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.sinh(n);
};

const fnCOSH: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.cosh(n);
};

const fnTANH: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.tanh(n);
};

const fnASINH: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : Math.asinh(n);
};

const fnACOSH: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.acosh(n);
};

const fnATANH: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n <= -1 || n >= 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.atanh(n);
};

// ============================================================================
// More Math Functions
// ============================================================================

const fnEVEN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  const sign = n >= 0 ? 1 : -1;
  const abs = Math.abs(n);
  const ceil = Math.ceil(abs);
  return sign * (ceil % 2 === 0 ? ceil : ceil + 1);
};

const fnODD: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n === 0) {
    return 1;
  }
  const sign = n >= 0 ? 1 : -1;
  const abs = Math.abs(n);
  const ceil = Math.ceil(abs);
  return sign * (ceil % 2 === 1 ? ceil : ceil + 1);
};

const fnMROUND: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const multiple = toNumber(args[1] as CalcValue);
  if (isError(multiple)) {
    return multiple;
  }
  if (multiple === 0) {
    return 0;
  }
  if ((num > 0 && multiple < 0) || (num < 0 && multiple > 0)) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.round(num / multiple) * multiple;
};

const fnQUOTIENT: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const den = toNumber(args[1] as CalcValue);
  if (isError(den)) {
    return den;
  }
  if (den === 0) {
    return { error: "#DIV/0!" } as CellErrorValue;
  }
  return Math.trunc(num / den);
};

const fnBASE: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const radix = toNumber(args[1] as CalcValue);
  if (isError(radix)) {
    return radix;
  }
  if (radix < 2 || radix > 36) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const minLen = args.length > 2 ? toNumber(args[2] as CalcValue) : 0;
  if (isError(minLen)) {
    return minLen;
  }
  const result = Math.floor(num).toString(Math.floor(radix)).toUpperCase();
  return minLen > 0 ? result.padStart(minLen, "0") : result;
};

const fnDECIMAL: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const radix = toNumber(args[1] as CalcValue);
  if (isError(radix)) {
    return radix;
  }
  if (radix < 2 || radix > 36) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const result = parseInt(text, Math.floor(radix));
  if (isNaN(result)) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return result;
};

const fnROMAN: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  let n = Math.floor(num);
  if (n < 0 || n > 3999) {
    return { error: "#VALUE!" } as CellErrorValue;
  }
  if (n === 0) {
    return "";
  }
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
  let result = "";
  for (let i = 0; i < vals.length; i++) {
    while (n >= vals[i]) {
      result += syms[i];
      n -= vals[i];
    }
  }
  return result;
};

const fnARABIC: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue)
    .toUpperCase()
    .trim();
  if (text === "") {
    return 0;
  }
  const romanMap: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let result = 0;
  for (let i = 0; i < text.length; i++) {
    const current = romanMap[text[i]];
    const next = romanMap[text[i + 1]];
    if (current === undefined) {
      return { error: "#VALUE!" } as CellErrorValue;
    }
    if (next && current < next) {
      result -= current;
    } else {
      result += current;
    }
  }
  return result;
};

const fnDEGREES: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : (n * 180) / Math.PI;
};

const fnRADIANS: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  return isError(n) ? n : (n * Math.PI) / 180;
};

const fnSUMX2MY2: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const xs = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const ys = flattenNumbers([args[1]]).filter((v): v is number => !isError(v));
  const n = Math.min(xs.length, ys.length);
  let sum = 0;
  for (let i = 0; i < n; i++) {
    sum += xs[i] * xs[i] - ys[i] * ys[i];
  }
  return sum;
};

const fnSUMX2PY2: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const xs = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const ys = flattenNumbers([args[1]]).filter((v): v is number => !isError(v));
  const n = Math.min(xs.length, ys.length);
  let sum = 0;
  for (let i = 0; i < n; i++) {
    sum += xs[i] * xs[i] + ys[i] * ys[i];
  }
  return sum;
};

const fnSUMXMY2: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const xs = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const ys = flattenNumbers([args[1]]).filter((v): v is number => !isError(v));
  const n = Math.min(xs.length, ys.length);
  let sum = 0;
  for (let i = 0; i < n; i++) {
    sum += (xs[i] - ys[i]) ** 2;
  }
  return sum;
};

const fnMULTINOMIAL: ExcelFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  let sum = 0;
  let denom = 1;
  for (const n of nums) {
    const ni = Math.floor(n as number);
    if (ni < 0) {
      return { error: "#NUM!" } as CellErrorValue;
    }
    sum += ni;
    for (let i = 2; i <= ni; i++) {
      denom *= i;
    }
  }
  let numer = 1;
  for (let i = 2; i <= sum; i++) {
    numer *= i;
  }
  return numer / denom;
};

// ============================================================================
// More Text Functions
// ============================================================================

const fnUNICHAR: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  const code = Math.floor(n);
  if (code < 1) {
    return { error: "#VALUE!" } as CellErrorValue;
  }
  try {
    return String.fromCodePoint(code);
  } catch {
    return { error: "#VALUE!" } as CellErrorValue;
  }
};

const fnUNICODE: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  if (text.length === 0) {
    return { error: "#VALUE!" } as CellErrorValue;
  }
  return text.codePointAt(0) ?? ({ error: "#VALUE!" } as CellErrorValue);
};

const fnBAHTTEXT: ExcelFunction = args => toString(args[0] as CalcValue);

const fnDOLLAR: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const decimals = args.length > 1 ? toNumber(args[1] as CalcValue) : 2;
  if (isError(decimals)) {
    return decimals;
  }
  const d = Math.max(0, Math.floor(decimals));
  const formatted = Math.abs(num).toFixed(d);
  const parts = formatted.split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const result = parts.join(".");
  return num < 0 ? `($${result})` : `$${result}`;
};

const fnFIXED: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const decimals = args.length > 1 ? toNumber(args[1] as CalcValue) : 2;
  if (isError(decimals)) {
    return decimals;
  }
  const noCommas = args.length > 2 ? toBoolean(args[2] as CalcValue) : false;
  if (isError(noCommas)) {
    return noCommas;
  }
  const d = Math.max(0, Math.floor(decimals));
  let result = num.toFixed(d);
  if (!noCommas) {
    const parts = result.split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    result = parts.join(".");
  }
  return result;
};

const fnASC: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  return text.replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xfee0));
};

const fnDBCS: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  return text.replace(/[!-~]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0xfee0));
};

const fnJIS: ExcelFunction = args => fnDBCS(args);

const fnPHONETIC: ExcelFunction = args => toString(args[0] as CalcValue);

// ============================================================================
// More Lookup Functions
// ============================================================================

const fnLOOKUP: ExcelFunction = args => {
  const lookupValue = args[0] as CalcValue;
  if (!Array.isArray(args[1])) {
    return { error: "#N/A" } as CellErrorValue;
  }
  const lookupArr = args[1] as CalcArray;

  if (args.length > 2 && Array.isArray(args[2])) {
    const resultArr = args[2] as CalcArray;
    const flat: CalcValue[] = [];
    const isRow = lookupArr.length === 1;
    if (isRow) {
      for (const cell of lookupArr[0]) {
        flat.push(cell);
      }
    } else {
      for (const row of lookupArr) {
        flat.push(row[0]);
      }
    }
    let bestIdx = -1;
    for (let i = 0; i < flat.length; i++) {
      const v = flat[i];
      if (typeof v === typeof lookupValue) {
        if (typeof v === "number" && typeof lookupValue === "number" && v <= lookupValue) {
          bestIdx = i;
        } else if (
          typeof v === "string" &&
          typeof lookupValue === "string" &&
          v.toLowerCase() <= lookupValue.toLowerCase()
        ) {
          bestIdx = i;
        }
      }
    }
    if (bestIdx === -1) {
      return { error: "#N/A" } as CellErrorValue;
    }
    if (isRow) {
      return resultArr.length === 1
        ? (resultArr[0][bestIdx] ?? null)
        : (resultArr[bestIdx]?.[0] ?? null);
    }
    return resultArr[bestIdx]?.[0] ?? null;
  }

  const rows = lookupArr.length;
  const cols = lookupArr[0]?.length ?? 0;
  if (cols === 0) {
    return { error: "#N/A" } as CellErrorValue;
  }
  if (cols >= rows) {
    let bestIdx = -1;
    for (let c = 0; c < cols; c++) {
      const v = lookupArr[0][c];
      if (typeof v === typeof lookupValue) {
        if (typeof v === "number" && typeof lookupValue === "number" && v <= lookupValue) {
          bestIdx = c;
        } else if (
          typeof v === "string" &&
          typeof lookupValue === "string" &&
          v.toLowerCase() <= lookupValue.toLowerCase()
        ) {
          bestIdx = c;
        }
      }
    }
    return bestIdx >= 0 ? lookupArr[rows - 1][bestIdx] : ({ error: "#N/A" } as CellErrorValue);
  }
  let bestIdx = -1;
  for (let r = 0; r < rows; r++) {
    const v = lookupArr[r][0];
    if (typeof v === typeof lookupValue) {
      if (typeof v === "number" && typeof lookupValue === "number" && v <= lookupValue) {
        bestIdx = r;
      } else if (
        typeof v === "string" &&
        typeof lookupValue === "string" &&
        v.toLowerCase() <= lookupValue.toLowerCase()
      ) {
        bestIdx = r;
      }
    }
  }
  return bestIdx >= 0 ? lookupArr[bestIdx][cols - 1] : ({ error: "#N/A" } as CellErrorValue);
};

const fnTRANSPOSE: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return [[args[0] as CalcValue]];
  }
  const arr = args[0] as CalcArray;
  const rows = arr.length;
  const cols = arr[0]?.length ?? 0;
  const result: CalcValue[][] = [];
  for (let c = 0; c < cols; c++) {
    const row: CalcValue[] = [];
    for (let r = 0; r < rows; r++) {
      row.push(arr[r]?.[c] ?? null);
    }
    result.push(row);
  }
  return result;
};

const fnAREAS: ExcelFunction = args =>
  args.length > 0 ? 1 : ({ error: "#VALUE!" } as CellErrorValue);

// ============================================================================
// More Date/Time Functions
// ============================================================================

const fnDAYS360: ExcelFunction = args => {
  const startN = toNumber(args[0] as CalcValue);
  if (isError(startN)) {
    return startN;
  }
  const endN = toNumber(args[1] as CalcValue);
  if (isError(endN)) {
    return endN;
  }
  const method = args.length > 2 ? toBoolean(args[2] as CalcValue) : false;
  if (isError(method)) {
    return method;
  }
  const sd = args[0] instanceof Date ? (args[0] as Date) : excelToDate(startN);
  const ed = args[1] instanceof Date ? (args[1] as Date) : excelToDate(endN);
  let d1 = sd.getDate();
  let d2 = ed.getDate();
  const m1 = sd.getMonth() + 1;
  const m2 = ed.getMonth() + 1;
  const y1 = sd.getFullYear();
  const y2 = ed.getFullYear();
  if (method) {
    if (d1 === 31) {
      d1 = 30;
    }
    if (d2 === 31) {
      d2 = 30;
    }
  } else {
    if (d1 === 31) {
      d1 = 30;
    }
    if (d2 === 31 && d1 >= 30) {
      d2 = 30;
    }
  }
  return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
};

const fnNETWORKDAYS_INTL: ExcelFunction = args => {
  const startN = toNumber(args[0] as CalcValue);
  if (isError(startN)) {
    return startN;
  }
  const endN = toNumber(args[1] as CalcValue);
  if (isError(endN)) {
    return endN;
  }
  const weekendArg = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(weekendArg)) {
    return weekendArg;
  }
  const holidays = args.length > 3 ? collectHolidays(args[3]) : new Set<number>();
  const weekendDays = new Set<number>();
  switch (weekendArg) {
    case 1:
      weekendDays.add(0).add(6);
      break;
    case 2:
      weekendDays.add(0).add(1);
      break;
    case 3:
      weekendDays.add(1).add(2);
      break;
    case 7:
      weekendDays.add(5).add(6);
      break;
    case 11:
      weekendDays.add(0);
      break;
    case 12:
      weekendDays.add(1);
      break;
    case 13:
      weekendDays.add(2);
      break;
    case 14:
      weekendDays.add(3);
      break;
    case 15:
      weekendDays.add(4);
      break;
    case 16:
      weekendDays.add(5);
      break;
    case 17:
      weekendDays.add(6);
      break;
    default:
      weekendDays.add(0).add(6);
      break;
  }
  const s = Math.floor(Math.min(startN, endN));
  const e = Math.floor(Math.max(startN, endN));
  const sign = startN <= endN ? 1 : -1;
  let count = 0;
  for (let d = s; d <= e; d++) {
    const dt = excelToDate(d);
    if (!weekendDays.has(dt.getDay()) && !holidays.has(d)) {
      count++;
    }
  }
  return count * sign;
};

const fnWORKDAY_INTL: ExcelFunction = args => {
  const startN = toNumber(args[0] as CalcValue);
  if (isError(startN)) {
    return startN;
  }
  const days = toNumber(args[1] as CalcValue);
  if (isError(days)) {
    return days;
  }
  const weekendArg = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(weekendArg)) {
    return weekendArg;
  }
  const holidays = args.length > 3 ? collectHolidays(args[3]) : new Set<number>();
  const weekendDays = new Set<number>();
  switch (weekendArg) {
    case 1:
      weekendDays.add(0).add(6);
      break;
    case 2:
      weekendDays.add(0).add(1);
      break;
    case 7:
      weekendDays.add(5).add(6);
      break;
    case 11:
      weekendDays.add(0);
      break;
    default:
      weekendDays.add(0).add(6);
      break;
  }
  let current = Math.floor(startN);
  const step = days >= 0 ? 1 : -1;
  let remaining = Math.abs(days);
  while (remaining > 0) {
    current += step;
    const dt = excelToDate(current);
    if (!weekendDays.has(dt.getDay()) && !holidays.has(current)) {
      remaining--;
    }
  }
  return current;
};

// ============================================================================
// More Financial Functions
// ============================================================================

const fnEFFECT: ExcelFunction = args => {
  const nomRate = toNumber(args[0] as CalcValue);
  if (isError(nomRate)) {
    return nomRate;
  }
  const npery = toNumber(args[1] as CalcValue);
  if (isError(npery)) {
    return npery;
  }
  if (nomRate <= 0 || npery < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.pow(1 + nomRate / Math.floor(npery), Math.floor(npery)) - 1;
};

const fnNOMINAL: ExcelFunction = args => {
  const effRate = toNumber(args[0] as CalcValue);
  if (isError(effRate)) {
    return effRate;
  }
  const npery = toNumber(args[1] as CalcValue);
  if (isError(npery)) {
    return npery;
  }
  if (effRate <= 0 || npery < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const np = Math.floor(npery);
  return np * (Math.pow(effRate + 1, 1 / np) - 1);
};

const fnXNPV: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  if (!Array.isArray(args[1]) || !Array.isArray(args[2])) {
    return { error: "#VALUE!" };
  }
  const values = flattenNumbers([args[1]]).filter((v): v is number => !isError(v));
  const dates = flattenNumbers([args[2]]).filter((v): v is number => !isError(v));
  if (values.length === 0 || values.length !== dates.length) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const d0 = dates[0];
  let npv = 0;
  for (let i = 0; i < values.length; i++) {
    npv += values[i] / Math.pow(1 + rate, (dates[i] - d0) / 365);
  }
  return npv;
};

const fnXIRR: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const values = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const dates = flattenNumbers([args[1]]).filter((v): v is number => !isError(v));
  if (values.length < 2 || values.length !== dates.length) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const guess = args.length > 2 ? toNumber(args[2] as CalcValue) : 0.1;
  if (isError(guess)) {
    return guess;
  }
  const d0 = dates[0];
  let g = guess as number;
  for (let iter = 0; iter < 100; iter++) {
    let npv = 0;
    let dnpv = 0;
    for (let i = 0; i < values.length; i++) {
      const t = (dates[i] - d0) / 365;
      npv += values[i] / Math.pow(1 + g, t);
      dnpv -= (t * values[i]) / Math.pow(1 + g, t + 1);
    }
    if (Math.abs(dnpv) < 1e-15) {
      break;
    }
    const newG = g - npv / dnpv;
    if (Math.abs(newG - g) < 1e-10) {
      return newG;
    }
    g = newG;
  }
  return { error: "#NUM!" } as CellErrorValue;
};

const fnMIRR: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const values = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const financeRate = toNumber(args[1] as CalcValue);
  if (isError(financeRate)) {
    return financeRate;
  }
  const reinvestRate = toNumber(args[2] as CalcValue);
  if (isError(reinvestRate)) {
    return reinvestRate;
  }
  const n = values.length;
  if (n < 2) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  let npvPos = 0;
  let npvNeg = 0;
  for (let i = 0; i < n; i++) {
    if (values[i] >= 0) {
      npvPos += values[i] * Math.pow(1 + reinvestRate, n - 1 - i);
    } else {
      npvNeg += values[i] / Math.pow(1 + financeRate, i);
    }
  }
  if (npvNeg === 0) {
    return { error: "#DIV/0!" } as CellErrorValue;
  }
  return Math.pow(-npvPos / npvNeg, 1 / (n - 1)) - 1;
};

const fnISPMT: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const per = toNumber(args[1] as CalcValue);
  if (isError(per)) {
    return per;
  }
  const nper = toNumber(args[2] as CalcValue);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumber(args[3] as CalcValue);
  if (isError(pv)) {
    return pv;
  }
  return pv * rate * (per / nper - 1);
};

const fnCUMPRINC: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumber(args[1] as CalcValue);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumber(args[2] as CalcValue);
  if (isError(pv)) {
    return pv;
  }
  const startPeriod = toNumber(args[3] as CalcValue);
  if (isError(startPeriod)) {
    return startPeriod;
  }
  const endPeriod = toNumber(args[4] as CalcValue);
  if (isError(endPeriod)) {
    return endPeriod;
  }
  const type = toNumber(args[5] as CalcValue);
  if (isError(type)) {
    return type;
  }
  if (rate <= 0 || nper <= 0 || pv <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  let cumPrinc = 0;
  for (let p = Math.floor(startPeriod); p <= Math.floor(endPeriod); p++) {
    cumPrinc += fnPPMT([rate, p, nper, pv, 0, type]) as number;
  }
  return cumPrinc;
};

const fnCUMIPMT: ExcelFunction = args => {
  const rate = toNumber(args[0] as CalcValue);
  if (isError(rate)) {
    return rate;
  }
  const nper = toNumber(args[1] as CalcValue);
  if (isError(nper)) {
    return nper;
  }
  const pv = toNumber(args[2] as CalcValue);
  if (isError(pv)) {
    return pv;
  }
  const startPeriod = toNumber(args[3] as CalcValue);
  if (isError(startPeriod)) {
    return startPeriod;
  }
  const endPeriod = toNumber(args[4] as CalcValue);
  if (isError(endPeriod)) {
    return endPeriod;
  }
  const type = toNumber(args[5] as CalcValue);
  if (isError(type)) {
    return type;
  }
  if (rate <= 0 || nper <= 0 || pv <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  let cumIpmt = 0;
  for (let p = Math.floor(startPeriod); p <= Math.floor(endPeriod); p++) {
    cumIpmt += fnIPMT([rate, p, nper, pv, 0, type]) as number;
  }
  return cumIpmt;
};

const fnDOLLARDE: ExcelFunction = args => {
  const fractionalDollar = toNumber(args[0] as CalcValue);
  if (isError(fractionalDollar)) {
    return fractionalDollar;
  }
  const fraction = toNumber(args[1] as CalcValue);
  if (isError(fraction)) {
    return fraction;
  }
  if (fraction < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const f = Math.floor(fraction);
  const intPart = Math.trunc(fractionalDollar);
  const fracPart = fractionalDollar - intPart;
  return intPart + (fracPart / f) * Math.pow(10, Math.ceil(Math.log10(f)));
};

const fnDOLLARFR: ExcelFunction = args => {
  const decimalDollar = toNumber(args[0] as CalcValue);
  if (isError(decimalDollar)) {
    return decimalDollar;
  }
  const fraction = toNumber(args[1] as CalcValue);
  if (isError(fraction)) {
    return fraction;
  }
  if (fraction < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const f = Math.floor(fraction);
  const intPart = Math.trunc(decimalDollar);
  const fracPart = decimalDollar - intPart;
  return intPart + (fracPart * f) / Math.pow(10, Math.ceil(Math.log10(f)));
};

const fnDISC: ExcelFunction = args => {
  const settlement = toNumber(args[0] as CalcValue);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumber(args[1] as CalcValue);
  if (isError(maturity)) {
    return maturity;
  }
  const pr = toNumber(args[2] as CalcValue);
  if (isError(pr)) {
    return pr;
  }
  const redemption = toNumber(args[3] as CalcValue);
  if (isError(redemption)) {
    return redemption;
  }
  const basis = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity) - Math.floor(settlement);
  if (days <= 0 || redemption <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const yearDays = basis === 1 ? 365.25 : basis === 3 ? 365 : 360;
  return ((redemption - pr) / redemption) * (yearDays / days);
};

const fnPRICEDISC: ExcelFunction = args => {
  const settlement = toNumber(args[0] as CalcValue);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumber(args[1] as CalcValue);
  if (isError(maturity)) {
    return maturity;
  }
  const disc = toNumber(args[2] as CalcValue);
  if (isError(disc)) {
    return disc;
  }
  const redemption = toNumber(args[3] as CalcValue);
  if (isError(redemption)) {
    return redemption;
  }
  const basis = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity) - Math.floor(settlement);
  const yearDays = basis === 1 ? 365.25 : basis === 3 ? 365 : 360;
  return redemption - disc * redemption * (days / yearDays);
};

const fnYIELDDISC: ExcelFunction = args => {
  const settlement = toNumber(args[0] as CalcValue);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumber(args[1] as CalcValue);
  if (isError(maturity)) {
    return maturity;
  }
  const pr = toNumber(args[2] as CalcValue);
  if (isError(pr)) {
    return pr;
  }
  const redemption = toNumber(args[3] as CalcValue);
  if (isError(redemption)) {
    return redemption;
  }
  const basis = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity) - Math.floor(settlement);
  if (days <= 0 || pr <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const yearDays = basis === 1 ? 365.25 : basis === 3 ? 365 : 360;
  return ((redemption - pr) / pr) * (yearDays / days);
};

const fnRECEIVED: ExcelFunction = args => {
  const settlement = toNumber(args[0] as CalcValue);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumber(args[1] as CalcValue);
  if (isError(maturity)) {
    return maturity;
  }
  const investment = toNumber(args[2] as CalcValue);
  if (isError(investment)) {
    return investment;
  }
  const disc = toNumber(args[3] as CalcValue);
  if (isError(disc)) {
    return disc;
  }
  const basis = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity) - Math.floor(settlement);
  const yearDays = basis === 1 ? 365.25 : basis === 3 ? 365 : 360;
  const denom = 1 - disc * (days / yearDays);
  if (denom === 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return investment / denom;
};

const fnINTRATE: ExcelFunction = args => {
  const settlement = toNumber(args[0] as CalcValue);
  if (isError(settlement)) {
    return settlement;
  }
  const maturity = toNumber(args[1] as CalcValue);
  if (isError(maturity)) {
    return maturity;
  }
  const investment = toNumber(args[2] as CalcValue);
  if (isError(investment)) {
    return investment;
  }
  const redemption = toNumber(args[3] as CalcValue);
  if (isError(redemption)) {
    return redemption;
  }
  const basis = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(basis)) {
    return basis;
  }
  const days = Math.floor(maturity) - Math.floor(settlement);
  if (days <= 0 || investment <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const yearDays = basis === 1 ? 365.25 : basis === 3 ? 365 : 360;
  return ((redemption - investment) / investment) * (yearDays / days);
};

// ============================================================================
// Statistical Distribution Helpers
// ============================================================================

function gammaFn(z: number): number {
  if (z < 0.5) {
    return Math.PI / (Math.sin(Math.PI * z) * gammaFn(1 - z));
  }
  z -= 1;
  const g = 7;
  const c = [
    0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313,
    -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6,
    1.5056327351493116e-7
  ];
  let x = c[0];
  for (let i = 1; i < g + 2; i++) {
    x += c[i] / (z + i);
  }
  const t = z + g + 0.5;
  return Math.sqrt(2 * Math.PI) * Math.pow(t, z + 0.5) * Math.exp(-t) * x;
}

function lnGamma(x: number): number {
  return Math.log(gammaFn(x));
}

function betaIncomplete(x: number, a: number, b: number): number {
  if (x <= 0) {
    return 0;
  }
  if (x >= 1) {
    return 1;
  }
  if (x > (a + 1) / (a + b + 2)) {
    return 1 - betaIncomplete(1 - x, b, a);
  }
  const lbeta = lnGamma(a) + lnGamma(b) - lnGamma(a + b);
  const front = Math.exp(Math.log(x) * a + Math.log(1 - x) * b - lbeta) / a;
  let f = 1;
  let c = 1;
  let d = 1 - ((a + b) * x) / (a + 1);
  if (Math.abs(d) < 1e-30) {
    d = 1e-30;
  }
  d = 1 / d;
  f = d;
  for (let m = 1; m <= 200; m++) {
    let num = (m * (b - m) * x) / ((a + 2 * m - 1) * (a + 2 * m));
    d = 1 + num * d;
    if (Math.abs(d) < 1e-30) {
      d = 1e-30;
    }
    c = 1 + num / c;
    if (Math.abs(c) < 1e-30) {
      c = 1e-30;
    }
    d = 1 / d;
    f *= c * d;
    num = -((a + m) * (a + b + m) * x) / ((a + 2 * m) * (a + 2 * m + 1));
    d = 1 + num * d;
    if (Math.abs(d) < 1e-30) {
      d = 1e-30;
    }
    c = 1 + num / c;
    if (Math.abs(c) < 1e-30) {
      c = 1e-30;
    }
    d = 1 / d;
    const delta = c * d;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-10) {
      break;
    }
  }
  return front * f;
}

function gammaIncomplete(a: number, x: number): number {
  if (x < 0) {
    return 0;
  }
  if (x === 0) {
    return 0;
  }
  if (x < a + 1) {
    let sum = 1 / a;
    let term = 1 / a;
    for (let n = 1; n <= 200; n++) {
      term *= x / (a + n);
      sum += term;
      if (Math.abs(term) < Math.abs(sum) * 1e-14) {
        break;
      }
    }
    return sum * Math.exp(-x + a * Math.log(x) - lnGamma(a));
  }
  let f = 1;
  const b0 = x + 1 - a;
  let ci = 1e30;
  let d = 1 / b0;
  f = d;
  for (let i = 1; i <= 200; i++) {
    const an = -i * (i - a);
    const bn = x + 2 * i + 1 - a;
    d = bn + an * d;
    if (Math.abs(d) < 1e-30) {
      d = 1e-30;
    }
    ci = bn + an / ci;
    if (Math.abs(ci) < 1e-30) {
      ci = 1e-30;
    }
    d = 1 / d;
    const delta = d * ci;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-10) {
      break;
    }
  }
  return 1 - f * Math.exp(-x + a * Math.log(x) - lnGamma(a));
}

// ============================================================================
// More Statistical Distribution Functions
// ============================================================================

const fnPOISSON_DIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const mean = toNumber(args[1] as CalcValue);
  if (isError(mean)) {
    return mean;
  }
  const cumulative = toBoolean(args[2] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  const k = Math.floor(x);
  if (k < 0 || mean < 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  if (!cumulative) {
    return Math.exp(-mean + k * Math.log(mean) - lnGamma(k + 1));
  }
  return 1 - gammaIncomplete(k + 1, mean);
};

const fnBINOM_DIST: ExcelFunction = args => {
  const numS = toNumber(args[0] as CalcValue);
  if (isError(numS)) {
    return numS;
  }
  const trials = toNumber(args[1] as CalcValue);
  if (isError(trials)) {
    return trials;
  }
  const probS = toNumber(args[2] as CalcValue);
  if (isError(probS)) {
    return probS;
  }
  const cumulative = toBoolean(args[3] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  const k = Math.floor(numS);
  const n = Math.floor(trials);
  if (k < 0 || n < 0 || k > n || probS < 0 || probS > 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const pmf = (ki: number): number => {
    const lnC = lnGamma(n + 1) - lnGamma(ki + 1) - lnGamma(n - ki + 1);
    return Math.exp(lnC + ki * Math.log(probS) + (n - ki) * Math.log(1 - probS));
  };
  if (!cumulative) {
    return pmf(k);
  }
  let sum = 0;
  for (let i = 0; i <= k; i++) {
    sum += pmf(i);
  }
  return sum;
};

const fnBINOM_INV: ExcelFunction = args => {
  const trials = toNumber(args[0] as CalcValue);
  if (isError(trials)) {
    return trials;
  }
  const probS = toNumber(args[1] as CalcValue);
  if (isError(probS)) {
    return probS;
  }
  const alpha = toNumber(args[2] as CalcValue);
  if (isError(alpha)) {
    return alpha;
  }
  const n = Math.floor(trials);
  if (n < 0 || probS < 0 || probS > 1 || alpha < 0 || alpha > 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  let cdf = 0;
  for (let k = 0; k <= n; k++) {
    const lnC = lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1);
    cdf += Math.exp(lnC + k * Math.log(probS) + (n - k) * Math.log(1 - probS));
    if (cdf >= alpha) {
      return k;
    }
  }
  return n;
};

const fnHYPGEOM_DIST: ExcelFunction = args => {
  const sampleS = toNumber(args[0] as CalcValue);
  if (isError(sampleS)) {
    return sampleS;
  }
  const numberSample = toNumber(args[1] as CalcValue);
  if (isError(numberSample)) {
    return numberSample;
  }
  const popS = toNumber(args[2] as CalcValue);
  if (isError(popS)) {
    return popS;
  }
  const numberPop = toNumber(args[3] as CalcValue);
  if (isError(numberPop)) {
    return numberPop;
  }
  const cumulative = toBoolean(args[4] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  const ss = Math.floor(sampleS);
  const ns = Math.floor(numberSample);
  const ps = Math.floor(popS);
  const np = Math.floor(numberPop);
  const pmf = (k: number): number =>
    Math.exp(
      lnGamma(ps + 1) -
        lnGamma(k + 1) -
        lnGamma(ps - k + 1) +
        lnGamma(np - ps + 1) -
        lnGamma(ns - k + 1) -
        lnGamma(np - ps - ns + k + 1) -
        lnGamma(np + 1) +
        lnGamma(ns + 1) +
        lnGamma(np - ns + 1)
    );
  if (!cumulative) {
    return pmf(ss);
  }
  let sum = 0;
  for (let k = 0; k <= ss; k++) {
    sum += pmf(k);
  }
  return sum;
};

const fnNEGBINOM_DIST: ExcelFunction = args => {
  const numF = toNumber(args[0] as CalcValue);
  if (isError(numF)) {
    return numF;
  }
  const numS = toNumber(args[1] as CalcValue);
  if (isError(numS)) {
    return numS;
  }
  const probS = toNumber(args[2] as CalcValue);
  if (isError(probS)) {
    return probS;
  }
  const cumulative = toBoolean(args[3] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  const f = Math.floor(numF);
  const s = Math.floor(numS);
  if (f < 0 || s < 1 || probS < 0 || probS > 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const pmf = (k: number): number => {
    const lnC = lnGamma(k + s) - lnGamma(s) - lnGamma(k + 1);
    return Math.exp(lnC + s * Math.log(probS) + k * Math.log(1 - probS));
  };
  if (!cumulative) {
    return pmf(f);
  }
  let sum = 0;
  for (let k = 0; k <= f; k++) {
    sum += pmf(k);
  }
  return sum;
};

const fnCHISQ_DIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const df = toNumber(args[1] as CalcValue);
  if (isError(df)) {
    return df;
  }
  const cumulative = toBoolean(args[2] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  if (x < 0 || df < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const k = Math.floor(df);
  if (cumulative) {
    return gammaIncomplete(k / 2, x / 2);
  }
  const halfK = k / 2;
  return Math.exp((halfK - 1) * Math.log(x / 2) - x / 2 - lnGamma(halfK)) / 2;
};

const fnCHISQ_INV: ExcelFunction = args => {
  const p = toNumber(args[0] as CalcValue);
  if (isError(p)) {
    return p;
  }
  const df = toNumber(args[1] as CalcValue);
  if (isError(df)) {
    return df;
  }
  if (p < 0 || p >= 1 || df < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  let x = df;
  for (let iter = 0; iter < 100; iter++) {
    const cdf = gammaIncomplete(df / 2, x / 2) as number;
    const halfK = df / 2;
    const pdf = Math.exp((halfK - 1) * Math.log(x / 2) - x / 2 - lnGamma(halfK)) / 2;
    if (Math.abs(pdf) < 1e-15) {
      break;
    }
    const delta = (cdf - p) / pdf;
    x -= delta;
    if (x <= 0) {
      x = 0.001;
    }
    if (Math.abs(delta) < 1e-10) {
      break;
    }
  }
  return x;
};

const fnCHISQ_DIST_RT: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const df = toNumber(args[1] as CalcValue);
  if (isError(df)) {
    return df;
  }
  if (x < 0 || df < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return 1 - (gammaIncomplete(Math.floor(df) / 2, x / 2) as number);
};

const fnF_DIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const df1 = toNumber(args[1] as CalcValue);
  if (isError(df1)) {
    return df1;
  }
  const df2 = toNumber(args[2] as CalcValue);
  if (isError(df2)) {
    return df2;
  }
  const cumulative = toBoolean(args[3] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  if (x < 0 || df1 < 1 || df2 < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const d1 = Math.floor(df1);
  const d2 = Math.floor(df2);
  if (cumulative) {
    return betaIncomplete((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2);
  }
  const num =
    (Math.pow(d1 * x, d1 / 2) * Math.pow(d2, d2 / 2)) / Math.pow(d1 * x + d2, (d1 + d2) / 2);
  const denom = x * Math.exp(lnGamma(d1 / 2) + lnGamma(d2 / 2) - lnGamma((d1 + d2) / 2));
  return denom === 0 ? 0 : num / denom;
};

const fnF_INV: ExcelFunction = args => {
  const p = toNumber(args[0] as CalcValue);
  if (isError(p)) {
    return p;
  }
  const df1 = toNumber(args[1] as CalcValue);
  if (isError(df1)) {
    return df1;
  }
  const df2 = toNumber(args[2] as CalcValue);
  if (isError(df2)) {
    return df2;
  }
  if (p < 0 || p >= 1 || df1 < 1 || df2 < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const d1 = Math.floor(df1);
  const d2 = Math.floor(df2);
  let lo = 0;
  let hi = 1000;
  for (let i = 0; i < 100; i++) {
    const mid = (lo + hi) / 2;
    const cdf = betaIncomplete((d1 * mid) / (d1 * mid + d2), d1 / 2, d2 / 2) as number;
    if (cdf < p) {
      lo = mid;
    } else {
      hi = mid;
    }
    if (hi - lo < 1e-10) {
      break;
    }
  }
  return (lo + hi) / 2;
};

const fnT_DIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const df = toNumber(args[1] as CalcValue);
  if (isError(df)) {
    return df;
  }
  const cumulative = toBoolean(args[2] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  if (df < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const v = Math.floor(df);
  if (cumulative) {
    const t = v / (v + x * x);
    return 1 - 0.5 * (betaIncomplete(t, v / 2, 0.5) as number);
  }
  return (
    Math.exp(lnGamma((v + 1) / 2) - lnGamma(v / 2)) /
    (Math.sqrt(v * Math.PI) * Math.pow(1 + (x * x) / v, (v + 1) / 2))
  );
};

const fnT_INV: ExcelFunction = args => {
  const p = toNumber(args[0] as CalcValue);
  if (isError(p)) {
    return p;
  }
  const df = toNumber(args[1] as CalcValue);
  if (isError(df)) {
    return df;
  }
  if (p <= 0 || p >= 1 || df < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  let x = normSInv(p);
  const v = Math.floor(df);
  for (let iter = 0; iter < 100; iter++) {
    const t = v / (v + x * x);
    const cdf = 1 - 0.5 * (betaIncomplete(t, v / 2, 0.5) as number);
    const pdf =
      Math.exp(lnGamma((v + 1) / 2) - lnGamma(v / 2)) /
      (Math.sqrt(v * Math.PI) * Math.pow(1 + (x * x) / v, (v + 1) / 2));
    if (Math.abs(pdf) < 1e-15) {
      break;
    }
    const delta = (cdf - p) / pdf;
    x -= delta;
    if (Math.abs(delta) < 1e-10) {
      break;
    }
  }
  return x;
};

const fnT_DIST_2T: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const df = toNumber(args[1] as CalcValue);
  if (isError(df)) {
    return df;
  }
  if (x < 0 || df < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const v = Math.floor(df);
  return betaIncomplete(v / (v + x * x), v / 2, 0.5);
};

const fnT_DIST_RT: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const df = toNumber(args[1] as CalcValue);
  if (isError(df)) {
    return df;
  }
  if (df < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const v = Math.floor(df);
  return 0.5 * (betaIncomplete(v / (v + x * x), v / 2, 0.5) as number);
};

const fnT_INV_2T: ExcelFunction = args => {
  const p = toNumber(args[0] as CalcValue);
  if (isError(p)) {
    return p;
  }
  const df = toNumber(args[1] as CalcValue);
  if (isError(df)) {
    return df;
  }
  if (p <= 0 || p > 1 || df < 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const result = fnT_INV([1 - p / 2, df]);
  if (isError(result as CalcValue)) {
    return result;
  }
  return Math.abs(result as number);
};

const fnBETA_DIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const alpha = toNumber(args[1] as CalcValue);
  if (isError(alpha)) {
    return alpha;
  }
  const beta = toNumber(args[2] as CalcValue);
  if (isError(beta)) {
    return beta;
  }
  const cumulative = args.length > 3 ? toBoolean(args[3] as CalcValue) : true;
  if (isError(cumulative)) {
    return cumulative;
  }
  const A = args.length > 4 ? toNumber(args[4] as CalcValue) : 0;
  if (isError(A)) {
    return A;
  }
  const B = args.length > 5 ? toNumber(args[5] as CalcValue) : 1;
  if (isError(B)) {
    return B;
  }
  if (alpha <= 0 || beta <= 0 || B <= A) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const xn = (x - A) / (B - A);
  if (xn < 0 || xn > 1) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  if (cumulative) {
    return betaIncomplete(xn, alpha, beta);
  }
  return (
    Math.exp(
      (alpha - 1) * Math.log(xn) +
        (beta - 1) * Math.log(1 - xn) -
        lnGamma(alpha) -
        lnGamma(beta) +
        lnGamma(alpha + beta)
    ) /
    (B - A)
  );
};

const fnBETA_INV: ExcelFunction = args => {
  const p = toNumber(args[0] as CalcValue);
  if (isError(p)) {
    return p;
  }
  const alpha = toNumber(args[1] as CalcValue);
  if (isError(alpha)) {
    return alpha;
  }
  const beta = toNumber(args[2] as CalcValue);
  if (isError(beta)) {
    return beta;
  }
  const A = args.length > 3 ? toNumber(args[3] as CalcValue) : 0;
  if (isError(A)) {
    return A;
  }
  const B = args.length > 4 ? toNumber(args[4] as CalcValue) : 1;
  if (isError(B)) {
    return B;
  }
  if (p < 0 || p > 1 || alpha <= 0 || beta <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  let lo = 0;
  let hi = 1;
  for (let i = 0; i < 100; i++) {
    const mid = (lo + hi) / 2;
    if ((betaIncomplete(mid, alpha, beta) as number) < p) {
      lo = mid;
    } else {
      hi = mid;
    }
    if (hi - lo < 1e-12) {
      break;
    }
  }
  return A + ((lo + hi) / 2) * (B - A);
};

const fnGAMMA: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n <= 0 && n === Math.floor(n)) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return gammaFn(n);
};

const fnGAMMALN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  if (n <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return lnGamma(n);
};

const fnGAMMA_DIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const alpha = toNumber(args[1] as CalcValue);
  if (isError(alpha)) {
    return alpha;
  }
  const beta = toNumber(args[2] as CalcValue);
  if (isError(beta)) {
    return beta;
  }
  const cumulative = toBoolean(args[3] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  if (x < 0 || alpha <= 0 || beta <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  if (cumulative) {
    return gammaIncomplete(alpha, x / beta);
  }
  return Math.exp((alpha - 1) * Math.log(x) - x / beta - alpha * Math.log(beta) - lnGamma(alpha));
};

const fnGAMMA_INV: ExcelFunction = args => {
  const p = toNumber(args[0] as CalcValue);
  if (isError(p)) {
    return p;
  }
  const alpha = toNumber(args[1] as CalcValue);
  if (isError(alpha)) {
    return alpha;
  }
  const beta = toNumber(args[2] as CalcValue);
  if (isError(beta)) {
    return beta;
  }
  if (p < 0 || p >= 1 || alpha <= 0 || beta <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  let lo = 0;
  let hi = Math.max(alpha * beta * 10, 100);
  for (let i = 0; i < 100; i++) {
    const mid = (lo + hi) / 2;
    if ((gammaIncomplete(alpha, mid / beta) as number) < p) {
      lo = mid;
    } else {
      hi = mid;
    }
    if (hi - lo < 1e-10) {
      break;
    }
  }
  return (lo + hi) / 2;
};

const fnEXPON_DIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const lambda = toNumber(args[1] as CalcValue);
  if (isError(lambda)) {
    return lambda;
  }
  const cumulative = toBoolean(args[2] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  if (x < 0 || lambda <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return cumulative ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x);
};

const fnWEIBULL_DIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const alpha = toNumber(args[1] as CalcValue);
  if (isError(alpha)) {
    return alpha;
  }
  const beta = toNumber(args[2] as CalcValue);
  if (isError(beta)) {
    return beta;
  }
  const cumulative = toBoolean(args[3] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  if (x < 0 || alpha <= 0 || beta <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  if (cumulative) {
    return 1 - Math.exp(-Math.pow(x / beta, alpha));
  }
  return (alpha / beta) * Math.pow(x / beta, alpha - 1) * Math.exp(-Math.pow(x / beta, alpha));
};

const fnLOGNORM_DIST: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const mean = toNumber(args[1] as CalcValue);
  if (isError(mean)) {
    return mean;
  }
  const stddev = toNumber(args[2] as CalcValue);
  if (isError(stddev)) {
    return stddev;
  }
  const cumulative = toBoolean(args[3] as CalcValue);
  if (isError(cumulative)) {
    return cumulative;
  }
  if (x <= 0 || stddev <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const z = (Math.log(x) - mean) / stddev;
  if (cumulative) {
    return normSDist(z);
  }
  return normSPdf(z) / (x * stddev);
};

const fnLOGNORM_INV: ExcelFunction = args => {
  const p = toNumber(args[0] as CalcValue);
  if (isError(p)) {
    return p;
  }
  const mean = toNumber(args[1] as CalcValue);
  if (isError(mean)) {
    return mean;
  }
  const stddev = toNumber(args[2] as CalcValue);
  if (isError(stddev)) {
    return stddev;
  }
  if (p <= 0 || p >= 1 || stddev <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.exp(mean + stddev * normSInv(p));
};

const fnPHI: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  return isError(x) ? x : normSPdf(x);
};

const fnGAUSS: ExcelFunction = args => {
  const z = toNumber(args[0] as CalcValue);
  return isError(z) ? z : normSDist(z) - 0.5;
};

function erfFn(x: number): number {
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(x);
  const t = 1.0 / (1.0 + p * ax);
  const y = 1.0 - ((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * Math.exp(-ax * ax);
  return sign * y;
}

const fnERF: ExcelFunction = args => {
  const lower = toNumber(args[0] as CalcValue);
  if (isError(lower)) {
    return lower;
  }
  if (args.length > 1) {
    const upper = toNumber(args[1] as CalcValue);
    if (isError(upper)) {
      return upper;
    }
    return erfFn(upper) - erfFn(lower);
  }
  return erfFn(lower);
};

const fnERFC: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  return isError(x) ? x : 1 - erfFn(x);
};

const fnSTANDARDIZE: ExcelFunction = args => {
  const x = toNumber(args[0] as CalcValue);
  if (isError(x)) {
    return x;
  }
  const mean = toNumber(args[1] as CalcValue);
  if (isError(mean)) {
    return mean;
  }
  const stddev = toNumber(args[2] as CalcValue);
  if (isError(stddev)) {
    return stddev;
  }
  if (stddev <= 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return (x - mean) / stddev;
};

const fnFREQUENCY: ExcelFunction = args => {
  if (!Array.isArray(args[0]) || !Array.isArray(args[1])) {
    return { error: "#VALUE!" };
  }
  const data = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const bins = flattenNumbers([args[1]]).filter((v): v is number => !isError(v));
  bins.sort((a, b) => a - b);
  const result: CalcValue[][] = [];
  for (let i = 0; i <= bins.length; i++) {
    let count = 0;
    for (const d of data) {
      if (i === 0 && d <= bins[0]) {
        count++;
      } else if (i === bins.length && d > bins[bins.length - 1]) {
        count++;
      } else if (i > 0 && i < bins.length && d > bins[i - 1] && d <= bins[i]) {
        count++;
      }
    }
    result.push([count]);
  }
  return result;
};

const fnGROWTH: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const knownY = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const knownX =
    args.length > 1 && Array.isArray(args[1])
      ? flattenNumbers([args[1]]).filter((v): v is number => !isError(v))
      : knownY.map((_, i) => i + 1);
  const newX =
    args.length > 2 && Array.isArray(args[2])
      ? flattenNumbers([args[2]]).filter((v): v is number => !isError(v))
      : knownX;
  const n = Math.min(knownX.length, knownY.length);
  if (n < 1) {
    return { error: "#VALUE!" };
  }
  let sumX = 0,
    sumLnY = 0,
    sumXLnY = 0,
    sumX2 = 0;
  for (let i = 0; i < n; i++) {
    if (knownY[i] <= 0) {
      return { error: "#NUM!" } as CellErrorValue;
    }
    sumX += knownX[i];
    sumLnY += Math.log(knownY[i]);
    sumXLnY += knownX[i] * Math.log(knownY[i]);
    sumX2 += knownX[i] * knownX[i];
  }
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) {
    return { error: "#DIV/0!" } as CellErrorValue;
  }
  const lnM = (n * sumXLnY - sumX * sumLnY) / denom;
  const lnB = (sumLnY - lnM * sumX) / n;
  return newX.map(x => [Math.exp(lnB + lnM * x)]);
};

const fnTREND: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const knownY = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const knownX =
    args.length > 1 && Array.isArray(args[1])
      ? flattenNumbers([args[1]]).filter((v): v is number => !isError(v))
      : knownY.map((_, i) => i + 1);
  const newX =
    args.length > 2 && Array.isArray(args[2])
      ? flattenNumbers([args[2]]).filter((v): v is number => !isError(v))
      : knownX;
  const n = Math.min(knownX.length, knownY.length);
  if (n < 1) {
    return { error: "#VALUE!" };
  }
  let sumX = 0,
    sumY = 0,
    sumXY = 0,
    sumX2 = 0;
  for (let i = 0; i < n; i++) {
    sumX += knownX[i];
    sumY += knownY[i];
    sumXY += knownX[i] * knownY[i];
    sumX2 += knownX[i] * knownX[i];
  }
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) {
    return { error: "#DIV/0!" } as CellErrorValue;
  }
  const m = (n * sumXY - sumX * sumY) / denom;
  const b = (sumY - m * sumX) / n;
  return newX.map(x => [b + m * x]);
};

const fnLINEST: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const knownY = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const knownX =
    args.length > 1 && Array.isArray(args[1])
      ? flattenNumbers([args[1]]).filter((v): v is number => !isError(v))
      : knownY.map((_, i) => i + 1);
  const n = Math.min(knownX.length, knownY.length);
  if (n < 1) {
    return { error: "#VALUE!" };
  }
  let sumX = 0,
    sumY = 0,
    sumXY = 0,
    sumX2 = 0;
  for (let i = 0; i < n; i++) {
    sumX += knownX[i];
    sumY += knownY[i];
    sumXY += knownX[i] * knownY[i];
    sumX2 += knownX[i] * knownX[i];
  }
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) {
    return { error: "#DIV/0!" } as CellErrorValue;
  }
  return [
    [(n * sumXY - sumX * sumY) / denom, (sumY - ((n * sumXY - sumX * sumY) / denom) * sumX) / n]
  ];
};

const fnLOGEST: ExcelFunction = args => {
  if (!Array.isArray(args[0])) {
    return { error: "#VALUE!" };
  }
  const knownY = flattenNumbers([args[0]]).filter((v): v is number => !isError(v));
  const knownX =
    args.length > 1 && Array.isArray(args[1])
      ? flattenNumbers([args[1]]).filter((v): v is number => !isError(v))
      : knownY.map((_, i) => i + 1);
  const n = Math.min(knownX.length, knownY.length);
  if (n < 1) {
    return { error: "#VALUE!" };
  }
  let sumX = 0,
    sumLnY = 0,
    sumXLnY = 0,
    sumX2 = 0;
  for (let i = 0; i < n; i++) {
    if (knownY[i] <= 0) {
      return { error: "#NUM!" } as CellErrorValue;
    }
    sumX += knownX[i];
    sumLnY += Math.log(knownY[i]);
    sumXLnY += knownX[i] * Math.log(knownY[i]);
    sumX2 += knownX[i] * knownX[i];
  }
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) {
    return { error: "#DIV/0!" } as CellErrorValue;
  }
  const lnM = (n * sumXLnY - sumX * sumLnY) / denom;
  const lnB = (sumLnY - lnM * sumX) / n;
  return [[Math.exp(lnM), Math.exp(lnB)]];
};

// ============================================================================
// Engineering: Complex Numbers, Bit Operations
// ============================================================================

function parseComplex(s: string): [number, number] | null {
  const text = s.trim();
  if (text === "") {
    return null;
  }
  const pureReal = Number(text);
  if (!isNaN(pureReal) && !text.endsWith("i") && !text.endsWith("j")) {
    return [pureReal, 0];
  }
  if (text === "i" || text === "j") {
    return [0, 1];
  }
  if (text === "-i" || text === "-j") {
    return [0, -1];
  }
  if ((text.endsWith("i") || text.endsWith("j")) && !isNaN(Number(text.slice(0, -1)))) {
    return [0, Number(text.slice(0, -1))];
  }
  const re = /^([+-]?[\d.eE+-]*?)([+-][\d.eE]*)?[ij]$/;
  const m = re.exec(text);
  if (!m) {
    return null;
  }
  const realPart = m[1] === "" || m[1] === "+" ? 0 : m[1] === "-" ? 0 : Number(m[1]);
  let imagPart = m[2] === undefined ? 0 : m[2] === "+" ? 1 : m[2] === "-" ? -1 : Number(m[2]);
  if (m[2] === undefined && m[1] !== "") {
    imagPart = m[1] === "" || m[1] === "+" ? 1 : m[1] === "-" ? -1 : Number(m[1]);
    return [0, imagPart];
  }
  return [isNaN(realPart) ? 0 : realPart, isNaN(imagPart) ? 0 : imagPart];
}

function formatComplex(re: number, im: number, suffix: string = "i"): string {
  if (im === 0) {
    return String(re);
  }
  if (re === 0) {
    if (im === 1) {
      return suffix;
    }
    if (im === -1) {
      return "-" + suffix;
    }
    return im + suffix;
  }
  const imStr =
    im === 1 ? "+" + suffix : im === -1 ? "-" + suffix : (im > 0 ? "+" : "") + im + suffix;
  return re + imStr;
}

const fnCOMPLEX: ExcelFunction = args => {
  const re = toNumber(args[0] as CalcValue);
  if (isError(re)) {
    return re;
  }
  const im = toNumber(args[1] as CalcValue);
  if (isError(im)) {
    return im;
  }
  const suffix = args.length > 2 ? toString(args[2] as CalcValue) : "i";
  if (suffix !== "i" && suffix !== "j") {
    return { error: "#VALUE!" } as CellErrorValue;
  }
  return formatComplex(re, im, suffix);
};
const fnIMREAL: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  return c ? c[0] : ({ error: "#NUM!" } as CellErrorValue);
};
const fnIMAGINARY: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  return c ? c[1] : ({ error: "#NUM!" } as CellErrorValue);
};
const fnIMABS: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.sqrt(c[0] * c[0] + c[1] * c[1]);
};
const fnIMARGUMENT: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  if (c[0] === 0 && c[1] === 0) {
    return { error: "#DIV/0!" } as CellErrorValue;
  }
  return Math.atan2(c[1], c[0]);
};
const fnIMCONJUGATE: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return formatComplex(c[0], -c[1]);
};
const fnIMSUM: ExcelFunction = args => {
  let re = 0,
    im = 0;
  for (const a of args) {
    const s = Array.isArray(a) ? toString(a[0]?.[0] ?? null) : toString(a as CalcValue);
    const c = parseComplex(s);
    if (!c) {
      return { error: "#NUM!" } as CellErrorValue;
    }
    re += c[0];
    im += c[1];
  }
  return formatComplex(re, im);
};
const fnIMSUB: ExcelFunction = args => {
  const c1 = parseComplex(toString(args[0] as CalcValue));
  const c2 = parseComplex(toString(args[1] as CalcValue));
  if (!c1 || !c2) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return formatComplex(c1[0] - c2[0], c1[1] - c2[1]);
};
const fnIMPRODUCT: ExcelFunction = args => {
  let re = 1,
    im = 0;
  for (const a of args) {
    const s = Array.isArray(a) ? toString(a[0]?.[0] ?? null) : toString(a as CalcValue);
    const c = parseComplex(s);
    if (!c) {
      return { error: "#NUM!" } as CellErrorValue;
    }
    const nRe = re * c[0] - im * c[1];
    const nIm = re * c[1] + im * c[0];
    re = nRe;
    im = nIm;
  }
  return formatComplex(re, im);
};
const fnIMDIV: ExcelFunction = args => {
  const c1 = parseComplex(toString(args[0] as CalcValue));
  const c2 = parseComplex(toString(args[1] as CalcValue));
  if (!c1 || !c2) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const d = c2[0] * c2[0] + c2[1] * c2[1];
  if (d === 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return formatComplex((c1[0] * c2[0] + c1[1] * c2[1]) / d, (c1[1] * c2[0] - c1[0] * c2[1]) / d);
};
const fnIMPOWER: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const n = toNumber(args[1] as CalcValue);
  if (isError(n)) {
    return n;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  const theta = Math.atan2(c[1], c[0]);
  const rn = Math.pow(r, n);
  return formatComplex(rn * Math.cos(n * theta), rn * Math.sin(n * theta));
};
const fnIMSQRT: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  const theta = Math.atan2(c[1], c[0]);
  const sr = Math.sqrt(r);
  return formatComplex(sr * Math.cos(theta / 2), sr * Math.sin(theta / 2));
};
const fnIMLN: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  if (r === 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return formatComplex(Math.log(r), Math.atan2(c[1], c[0]));
};
const fnIMLOG2: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  if (r === 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const ln2 = Math.log(2);
  return formatComplex(Math.log(r) / ln2, Math.atan2(c[1], c[0]) / ln2);
};
const fnIMLOG10: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  if (r === 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const ln10 = Math.log(10);
  return formatComplex(Math.log(r) / ln10, Math.atan2(c[1], c[0]) / ln10);
};
const fnIMEXP: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  const er = Math.exp(c[0]);
  return formatComplex(er * Math.cos(c[1]), er * Math.sin(c[1]));
};
const fnIMSIN: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return formatComplex(Math.sin(c[0]) * Math.cosh(c[1]), Math.cos(c[0]) * Math.sinh(c[1]));
};
const fnIMCOS: ExcelFunction = args => {
  const c = parseComplex(toString(args[0] as CalcValue));
  if (!c) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return formatComplex(Math.cos(c[0]) * Math.cosh(c[1]), -Math.sin(c[0]) * Math.sinh(c[1]));
};

const fnBITAND: ExcelFunction = args => {
  const a = toNumber(args[0] as CalcValue);
  if (isError(a)) {
    return a;
  }
  const b = toNumber(args[1] as CalcValue);
  if (isError(b)) {
    return b;
  }
  if (a < 0 || b < 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return (Math.floor(a) & Math.floor(b)) >>> 0;
};
const fnBITOR: ExcelFunction = args => {
  const a = toNumber(args[0] as CalcValue);
  if (isError(a)) {
    return a;
  }
  const b = toNumber(args[1] as CalcValue);
  if (isError(b)) {
    return b;
  }
  if (a < 0 || b < 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return (Math.floor(a) | Math.floor(b)) >>> 0;
};
const fnBITXOR: ExcelFunction = args => {
  const a = toNumber(args[0] as CalcValue);
  if (isError(a)) {
    return a;
  }
  const b = toNumber(args[1] as CalcValue);
  if (isError(b)) {
    return b;
  }
  if (a < 0 || b < 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return (Math.floor(a) ^ Math.floor(b)) >>> 0;
};
const fnBITLSHIFT: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const shift = toNumber(args[1] as CalcValue);
  if (isError(shift)) {
    return shift;
  }
  if (num < 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.floor(num) * Math.pow(2, Math.floor(shift));
};
const fnBITRSHIFT: ExcelFunction = args => {
  const num = toNumber(args[0] as CalcValue);
  if (isError(num)) {
    return num;
  }
  const shift = toNumber(args[1] as CalcValue);
  if (isError(shift)) {
    return shift;
  }
  if (num < 0) {
    return { error: "#NUM!" } as CellErrorValue;
  }
  return Math.floor(Math.floor(num) / Math.pow(2, Math.floor(shift)));
};

// ============================================================================
// Information Functions
// ============================================================================

const fnERROR_TYPE: ExcelFunction = args => {
  const v = args[0] as CalcValue;
  if (!isError(v)) {
    return { error: "#N/A" } as CellErrorValue;
  }
  switch (v.error) {
    case "#NULL!":
      return 1;
    case "#DIV/0!":
      return 2;
    case "#VALUE!":
      return 3;
    case "#REF!":
      return 4;
    case "#NAME?":
      return 5;
    case "#NUM!":
      return 6;
    case "#N/A":
      return 7;
    default:
      return { error: "#N/A" } as CellErrorValue;
  }
};

const fnISEVEN: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  return Math.floor(n) % 2 === 0;
};
const fnISODD: ExcelFunction = args => {
  const n = toNumber(args[0] as CalcValue);
  if (isError(n)) {
    return n;
  }
  return Math.floor(n) % 2 !== 0;
};
const fnNA: ExcelFunction = () => ({ error: "#N/A" }) as CellErrorValue;

// Stubs for functions that need runtime context or are rarely used
const fnINFO_STUB: ExcelFunction = () => ({ error: "#N/A" }) as CellErrorValue;
const fnCELL_STUB: ExcelFunction = () => ({ error: "#N/A" }) as CellErrorValue;
const fnISREF_STUB: ExcelFunction = () => false;
const fnSHEET_STUB: ExcelFunction = () => 1;
const fnSHEETS_STUB: ExcelFunction = () => 1;

// Lambda Array Functions (stubs — full invocation requires evaluator context)
const fnMAP: ExcelFunction = () => ({ error: "#CALC!" }) as CellErrorValue;
const fnREDUCE: ExcelFunction = () => ({ error: "#CALC!" }) as CellErrorValue;
const fnSCAN: ExcelFunction = () => ({ error: "#CALC!" }) as CellErrorValue;
const fnMAKEARRAY: ExcelFunction = () => ({ error: "#CALC!" }) as CellErrorValue;
const fnBYROW: ExcelFunction = () => ({ error: "#CALC!" }) as CellErrorValue;
const fnBYCOL: ExcelFunction = () => ({ error: "#CALC!" }) as CellErrorValue;

// ============================================================================
// Function Registry
// ============================================================================

// ============================================================================
// Excel 365 Text Functions: TEXTBEFORE, TEXTAFTER, TEXTSPLIT
// ============================================================================

const fnTEXTBEFORE: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const delimiter = toString(args[1] as CalcValue);
  const instanceNum = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(instanceNum)) {
    return instanceNum;
  }
  const inst = instanceNum as number;
  if (inst === 0) {
    return { error: "#VALUE!" } as CellErrorValue;
  }
  if (delimiter === "") {
    return inst > 0 ? "" : text;
  }
  // Find the nth occurrence
  if (inst > 0) {
    let pos = -1;
    for (let i = 0; i < inst; i++) {
      pos = text.indexOf(delimiter, pos + 1);
      if (pos === -1) {
        // match_mode/if_not_found handling simplified
        return args.length > 5 ? (args[5] as CalcValue) : ({ error: "#N/A" } as CellErrorValue);
      }
    }
    return text.slice(0, pos);
  }
  // Negative: search from end
  let pos = text.length;
  for (let i = 0; i < -inst; i++) {
    pos = text.lastIndexOf(delimiter, pos - 1);
    if (pos === -1) {
      return args.length > 5 ? (args[5] as CalcValue) : ({ error: "#N/A" } as CellErrorValue);
    }
  }
  return text.slice(0, pos);
};

const fnTEXTAFTER: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const delimiter = toString(args[1] as CalcValue);
  const instanceNum = args.length > 2 ? toNumber(args[2] as CalcValue) : 1;
  if (isError(instanceNum)) {
    return instanceNum;
  }
  const inst = instanceNum as number;
  if (inst === 0) {
    return { error: "#VALUE!" } as CellErrorValue;
  }
  if (delimiter === "") {
    return inst > 0 ? text : "";
  }
  if (inst > 0) {
    let pos = -1;
    for (let i = 0; i < inst; i++) {
      pos = text.indexOf(delimiter, pos + 1);
      if (pos === -1) {
        return args.length > 5 ? (args[5] as CalcValue) : ({ error: "#N/A" } as CellErrorValue);
      }
    }
    return text.slice(pos + delimiter.length);
  }
  let pos = text.length;
  for (let i = 0; i < -inst; i++) {
    pos = text.lastIndexOf(delimiter, pos - 1);
    if (pos === -1) {
      return args.length > 5 ? (args[5] as CalcValue) : ({ error: "#N/A" } as CellErrorValue);
    }
  }
  return text.slice(pos + delimiter.length);
};

const fnTEXTSPLIT: ExcelFunction = args => {
  const text = toString(args[0] as CalcValue);
  const colDelimiter = args.length > 1 ? toString(args[1] as CalcValue) : "";
  const rowDelimiter = args.length > 2 && args[2] !== null ? toString(args[2] as CalcValue) : "";

  let rows: string[];
  if (rowDelimiter) {
    rows = text.split(rowDelimiter);
  } else {
    rows = [text];
  }

  const result: CalcValue[][] = [];
  for (const row of rows) {
    if (colDelimiter) {
      result.push(row.split(colDelimiter));
    } else {
      result.push([row]);
    }
  }
  return result;
};

export const FUNCTIONS: Record<string, ExcelFunction> = {
  // Aggregate / Statistical
  SUM: fnSUM,
  AVERAGE: fnAVERAGE,
  MIN: fnMIN,
  MAX: fnMAX,
  COUNT: fnCOUNT,
  COUNTA: fnCOUNTA,
  COUNTBLANK: fnCOUNTBLANK,
  PRODUCT: fnPRODUCT,
  SUMPRODUCT: fnSUMPRODUCT,
  MEDIAN: fnMEDIAN,
  LARGE: fnLARGE,
  SMALL: fnSMALL,
  RANK: fnRANK,
  "RANK.EQ": fnRANK,
  STDEV: fnSTDEV,
  "STDEV.S": fnSTDEV,
  STDEVP: fnSTDEVP,
  "STDEV.P": fnSTDEVP,
  VAR: fnVAR,
  "VAR.S": fnVAR,
  VARP: fnVARP,
  "VAR.P": fnVARP,
  SUMSQ: fnSUMSQ,

  // Conditional Aggregates
  SUMIF: fnSUMIF,
  SUMIFS: fnSUMIFS,
  COUNTIF: fnCOUNTIF,
  COUNTIFS: fnCOUNTIFS,
  AVERAGEIF: fnAVERAGEIF,
  AVERAGEIFS: fnAVERAGEIFS,
  MAXIFS: fnMAXIFS,
  MINIFS: fnMINIFS,

  // Math
  ABS: fnABS,
  CEILING: fnCEILING,
  "CEILING.MATH": fnCEILING,
  FLOOR: fnFLOOR,
  "FLOOR.MATH": fnFLOOR,
  INT: fnINT,
  MOD: fnMOD,
  POWER: fnPOWER,
  ROUND: fnROUND,
  ROUNDDOWN: fnROUNDDOWN,
  ROUNDUP: fnROUNDUP,
  TRUNC: fnTRUNC,
  SQRT: fnSQRT,
  LN: fnLN,
  LOG: fnLOG,
  LOG10: fnLOG10,
  EXP: fnEXP,
  PI: fnPI,
  RAND: fnRAND,
  RANDBETWEEN: fnRANDBETWEEN,
  SIGN: fnSIGN,
  GCD: fnGCD,
  LCM: fnLCM,

  // Logical
  IF: fnIF,
  IFS: fnIFS,
  AND: fnAND,
  OR: fnOR,
  NOT: fnNOT,
  XOR: fnXOR,
  SWITCH: fnSWITCH,
  CHOOSE: fnCHOOSE,
  IFERROR: fnIFERROR,
  IFNA: fnIFNA,

  // Text
  CONCATENATE: fnCONCATENATE,
  CONCAT: fnCONCAT,
  "_XLFN.CONCAT": fnCONCAT,
  TEXTJOIN: fnTEXTJOIN,
  "_XLFN.TEXTJOIN": fnTEXTJOIN,
  LEFT: fnLEFT,
  RIGHT: fnRIGHT,
  MID: fnMID,
  LEN: fnLEN,
  TRIM: fnTRIM,
  LOWER: fnLOWER,
  UPPER: fnUPPER,
  PROPER: fnPROPER,
  SUBSTITUTE: fnSUBSTITUTE,
  REPLACE: fnREPLACE,
  FIND: fnFIND,
  SEARCH: fnSEARCH,
  REPT: fnREPT,
  TEXT: fnTEXT,
  VALUE: fnVALUE,
  EXACT: fnEXACT,
  CODE: fnCODE,
  CHAR: fnCHAR,
  CLEAN: fnCLEAN,
  T: fnT,

  // Information
  ISNUMBER: fnISNUMBER,
  ISTEXT: fnISTEXT,
  ISBLANK: fnISBLANK,
  ISLOGICAL: fnISLOGICAL,
  ISERROR: fnISERROR,
  ISERR: fnISERR,
  ISNA: fnISNA,
  ISNONTEXT: fnISNONTEXT,
  N: fnN,
  TYPE: fnTYPE,

  // Lookup
  ROW: fnROW,
  COLUMN: fnCOLUMN,
  ROWS: fnROWS,
  COLUMNS: fnCOLUMNS,
  INDEX: fnINDEX,
  MATCH: fnMATCH,
  VLOOKUP: fnVLOOKUP,
  HLOOKUP: fnHLOOKUP,
  XLOOKUP: fnXLOOKUP,
  "_XLFN.XLOOKUP": fnXLOOKUP,
  XMATCH: fnXMATCH,
  "_XLFN.XMATCH": fnXMATCH,
  ADDRESS: fnADDRESS,

  // Date / Time
  TODAY: fnTODAY,
  NOW: fnNOW,
  YEAR: fnYEAR,
  MONTH: fnMONTH,
  DAY: fnDAY,
  DATE: fnDATE,
  TIME: fnTIME,
  HOUR: fnHOUR,
  MINUTE: fnMINUTE,
  SECOND: fnSECOND,
  WEEKDAY: fnWEEKDAY,
  EOMONTH: fnEOMONTH,
  EDATE: fnEDATE,
  DATEDIF: fnDATEDIF,
  DAYS: fnDAYS,
  ISOWEEKNUM: fnISOWEEKNUM,
  WEEKNUM: fnWEEKNUM,
  NETWORKDAYS: fnNETWORKDAYS,
  WORKDAY: fnWORKDAY,
  YEARFRAC: fnYEARFRAC,
  DATEVALUE: fnDATEVALUE,
  TIMEVALUE: fnTIMEVALUE,
  NUMBERVALUE: fnNUMBERVALUE,
  "_XLFN.NUMBERVALUE": fnNUMBERVALUE,

  // Dynamic Array
  FILTER: fnFILTER,
  "_XLFN._XLWS.FILTER": fnFILTER,
  SORT: fnSORT,
  "_XLFN._XLWS.SORT": fnSORT,
  UNIQUE: fnUNIQUE,
  "_XLFN._XLWS.UNIQUE": fnUNIQUE,
  SORTBY: fnSORTBY,
  "_XLFN._XLWS.SORTBY": fnSORTBY,
  SEQUENCE: fnSEQUENCE,
  "_XLFN.SEQUENCE": fnSEQUENCE,
  RANDARRAY: fnRANDARRAY,
  "_XLFN.RANDARRAY": fnRANDARRAY,

  // Aggregate / Subtotal
  SUBTOTAL: fnSUBTOTAL,
  AGGREGATE: fnAGGREGATE,

  // Financial
  PMT: fnPMT,
  FV: fnFV,
  PV: fnPV,
  NPV: fnNPV,
  IRR: fnIRR,
  NPER: fnNPER,
  RATE: fnRATE,
  SLN: fnSLN,
  DB: fnDB,
  DDB: fnDDB,
  IPMT: fnIPMT,
  PPMT: fnPPMT,

  // Engineering
  BIN2DEC: fnBIN2DEC,
  DEC2BIN: fnDEC2BIN,
  HEX2DEC: fnHEX2DEC,
  DEC2HEX: fnDEC2HEX,
  OCT2DEC: fnOCT2DEC,
  DEC2OCT: fnDEC2OCT,
  DELTA: fnDELTA,
  GESTEP: fnGESTEP,

  // Advanced Statistics
  "NORM.S.DIST": fnNORMSDIST,
  "_XLFN.NORM.S.DIST": fnNORMSDIST,
  NORMSDIST: fnNORMSDIST,
  "NORM.DIST": fnNORMDIST,
  "_XLFN.NORM.DIST": fnNORMDIST,
  NORMDIST: fnNORMDIST,
  "NORM.S.INV": fnNORMSINV,
  "_XLFN.NORM.S.INV": fnNORMSINV,
  NORMSINV: fnNORMSINV,
  "NORM.INV": fnNORMINV,
  "_XLFN.NORM.INV": fnNORMINV,
  NORMINV: fnNORMINV,
  PERCENTILE: fnPERCENTILE,
  "PERCENTILE.INC": fnPERCENTILE,
  "_XLFN.PERCENTILE.INC": fnPERCENTILE,
  "PERCENTILE.EXC": fnPERCENTILEEXC,
  "_XLFN.PERCENTILE.EXC": fnPERCENTILEEXC,
  QUARTILE: fnQUARTILE,
  "QUARTILE.INC": fnQUARTILE,
  "_XLFN.QUARTILE.INC": fnQUARTILE,
  "QUARTILE.EXC": fnQUARTILEEXC,
  "_XLFN.QUARTILE.EXC": fnQUARTILEEXC,
  MODE: fnMODE,
  "MODE.SNGL": fnMODE,
  "_XLFN.MODE.SNGL": fnMODE,
  CORREL: fnCORREL,
  SLOPE: fnSLOPE,
  INTERCEPT: fnINTERCEPT,
  RSQ: fnRSQ,
  FORECAST: fnFORECAST,
  "FORECAST.LINEAR": fnFORECAST,
  "_XLFN.FORECAST.LINEAR": fnFORECAST,
  FACT: fnFACT,
  FACTDOUBLE: fnFACTDOUBLE,
  COMBIN: fnCOMBIN,
  COMBINA: fnCOMBINA,
  PERMUT: fnPERMUT,
  GEOMEAN: fnGEOMEAN,
  HARMEAN: fnHARMEAN,
  TRIMMEAN: fnTRIMMEAN,
  DEVSQ: fnDEVSQ,
  AVEDEV: fnAVEDEV,
  "CONFIDENCE.NORM": fnCONFIDENCENORM,
  "_XLFN.CONFIDENCE.NORM": fnCONFIDENCENORM,
  CONFIDENCE: fnCONFIDENCENORM,
  FISHER: fnFISHER,
  FISHERINV: fnFISHERINV,
  AVERAGEA: fnAVERAGEA,
  MAXA: fnMAXA,
  MINA: fnMINA,

  // Database
  DSUM: fnDSUM,
  DAVERAGE: fnDAVERAGE,
  DCOUNT: fnDCOUNT,
  DMAX: fnDMAX,
  DMIN: fnDMIN,
  DPRODUCT: fnDPRODUCT,
  DGET: fnDGET,

  // Array manipulation
  TOCOL: fnTOCOL,
  "_XLFN.TOCOL": fnTOCOL,
  TOROW: fnTOROW,
  "_XLFN.TOROW": fnTOROW,
  CHOOSEROWS: fnCHOOSEROWS,
  "_XLFN.CHOOSEROWS": fnCHOOSEROWS,
  CHOOSECOLS: fnCHOOSECOLS,
  "_XLFN.CHOOSECOLS": fnCHOOSECOLS,
  VSTACK: fnVSTACK,
  "_XLFN.VSTACK": fnVSTACK,
  HSTACK: fnHSTACK,
  "_XLFN.HSTACK": fnHSTACK,
  TAKE: fnTAKE,
  "_XLFN.TAKE": fnTAKE,
  DROP: fnDROP,
  "_XLFN.DROP": fnDROP,
  WRAPROWS: fnWRAPROWS,
  "_XLFN.WRAPROWS": fnWRAPROWS,
  WRAPCOLS: fnWRAPCOLS,
  "_XLFN.WRAPCOLS": fnWRAPCOLS,
  EXPAND: fnEXPAND,
  "_XLFN.EXPAND": fnEXPAND,

  // Trigonometric
  SIN: fnSIN,
  COS: fnCOS,
  TAN: fnTAN,
  ASIN: fnASIN,
  ACOS: fnACOS,
  ATAN: fnATAN,
  ATAN2: fnATAN2,
  SINH: fnSINH,
  COSH: fnCOSH,
  TANH: fnTANH,
  ASINH: fnASINH,
  ACOSH: fnACOSH,
  ATANH: fnATANH,

  // More Math
  EVEN: fnEVEN,
  ODD: fnODD,
  MROUND: fnMROUND,
  QUOTIENT: fnQUOTIENT,
  BASE: fnBASE,
  "_XLFN.BASE": fnBASE,
  DECIMAL: fnDECIMAL,
  "_XLFN.DECIMAL": fnDECIMAL,
  ROMAN: fnROMAN,
  ARABIC: fnARABIC,
  "_XLFN.ARABIC": fnARABIC,
  DEGREES: fnDEGREES,
  RADIANS: fnRADIANS,
  SUMX2MY2: fnSUMX2MY2,
  SUMX2PY2: fnSUMX2PY2,
  SUMXMY2: fnSUMXMY2,
  MULTINOMIAL: fnMULTINOMIAL,

  // More Text
  UNICHAR: fnUNICHAR,
  "_XLFN.UNICHAR": fnUNICHAR,
  UNICODE: fnUNICODE,
  "_XLFN.UNICODE": fnUNICODE,
  BAHTTEXT: fnBAHTTEXT,
  DOLLAR: fnDOLLAR,
  FIXED: fnFIXED,
  ASC: fnASC,
  DBCS: fnDBCS,
  JIS: fnJIS,
  PHONETIC: fnPHONETIC,

  // More Lookup
  LOOKUP: fnLOOKUP,
  TRANSPOSE: fnTRANSPOSE,
  AREAS: fnAREAS,

  // More Date/Time
  DAYS360: fnDAYS360,
  "NETWORKDAYS.INTL": fnNETWORKDAYS_INTL,
  "_XLFN.NETWORKDAYS.INTL": fnNETWORKDAYS_INTL,
  "WORKDAY.INTL": fnWORKDAY_INTL,
  "_XLFN.WORKDAY.INTL": fnWORKDAY_INTL,

  // Financial
  EFFECT: fnEFFECT,
  NOMINAL: fnNOMINAL,
  XNPV: fnXNPV,
  XIRR: fnXIRR,
  MIRR: fnMIRR,
  ISPMT: fnISPMT,
  CUMPRINC: fnCUMPRINC,
  CUMIPMT: fnCUMIPMT,
  DOLLARDE: fnDOLLARDE,
  DOLLARFR: fnDOLLARFR,
  DISC: fnDISC,
  PRICEDISC: fnPRICEDISC,
  YIELDDISC: fnYIELDDISC,
  RECEIVED: fnRECEIVED,
  INTRATE: fnINTRATE,

  // More Statistical Distributions
  "POISSON.DIST": fnPOISSON_DIST,
  "_XLFN.POISSON.DIST": fnPOISSON_DIST,
  "BINOM.DIST": fnBINOM_DIST,
  "_XLFN.BINOM.DIST": fnBINOM_DIST,
  "BINOM.INV": fnBINOM_INV,
  "_XLFN.BINOM.INV": fnBINOM_INV,
  "HYPGEOM.DIST": fnHYPGEOM_DIST,
  "_XLFN.HYPGEOM.DIST": fnHYPGEOM_DIST,
  "NEGBINOM.DIST": fnNEGBINOM_DIST,
  "_XLFN.NEGBINOM.DIST": fnNEGBINOM_DIST,
  "CHISQ.DIST": fnCHISQ_DIST,
  "_XLFN.CHISQ.DIST": fnCHISQ_DIST,
  "CHISQ.INV": fnCHISQ_INV,
  "_XLFN.CHISQ.INV": fnCHISQ_INV,
  "CHISQ.DIST.RT": fnCHISQ_DIST_RT,
  "_XLFN.CHISQ.DIST.RT": fnCHISQ_DIST_RT,
  "F.DIST": fnF_DIST,
  "_XLFN.F.DIST": fnF_DIST,
  "F.INV": fnF_INV,
  "_XLFN.F.INV": fnF_INV,
  "T.DIST": fnT_DIST,
  "_XLFN.T.DIST": fnT_DIST,
  "T.INV": fnT_INV,
  "_XLFN.T.INV": fnT_INV,
  "T.DIST.2T": fnT_DIST_2T,
  "_XLFN.T.DIST.2T": fnT_DIST_2T,
  "T.DIST.RT": fnT_DIST_RT,
  "_XLFN.T.DIST.RT": fnT_DIST_RT,
  "T.INV.2T": fnT_INV_2T,
  "_XLFN.T.INV.2T": fnT_INV_2T,
  "BETA.DIST": fnBETA_DIST,
  "_XLFN.BETA.DIST": fnBETA_DIST,
  "BETA.INV": fnBETA_INV,
  "_XLFN.BETA.INV": fnBETA_INV,
  GAMMA: fnGAMMA,
  "_XLFN.GAMMA": fnGAMMA,
  GAMMALN: fnGAMMALN,
  "GAMMALN.PRECISE": fnGAMMALN,
  "_XLFN.GAMMALN.PRECISE": fnGAMMALN,
  "GAMMA.DIST": fnGAMMA_DIST,
  "_XLFN.GAMMA.DIST": fnGAMMA_DIST,
  "GAMMA.INV": fnGAMMA_INV,
  "_XLFN.GAMMA.INV": fnGAMMA_INV,
  "EXPON.DIST": fnEXPON_DIST,
  "_XLFN.EXPON.DIST": fnEXPON_DIST,
  "WEIBULL.DIST": fnWEIBULL_DIST,
  "_XLFN.WEIBULL.DIST": fnWEIBULL_DIST,
  "LOGNORM.DIST": fnLOGNORM_DIST,
  "_XLFN.LOGNORM.DIST": fnLOGNORM_DIST,
  "LOGNORM.INV": fnLOGNORM_INV,
  "_XLFN.LOGNORM.INV": fnLOGNORM_INV,
  PHI: fnPHI,
  "_XLFN.PHI": fnPHI,
  GAUSS: fnGAUSS,
  "_XLFN.GAUSS": fnGAUSS,
  ERF: fnERF,
  "ERF.PRECISE": fnERF,
  "_XLFN.ERF.PRECISE": fnERF,
  ERFC: fnERFC,
  "ERFC.PRECISE": fnERFC,
  "_XLFN.ERFC.PRECISE": fnERFC,
  STANDARDIZE: fnSTANDARDIZE,
  FREQUENCY: fnFREQUENCY,
  GROWTH: fnGROWTH,
  TREND: fnTREND,
  LINEST: fnLINEST,
  LOGEST: fnLOGEST,

  // Engineering
  COMPLEX: fnCOMPLEX,
  IMREAL: fnIMREAL,
  IMAGINARY: fnIMAGINARY,
  IMABS: fnIMABS,
  IMARGUMENT: fnIMARGUMENT,
  IMCONJUGATE: fnIMCONJUGATE,
  IMSUM: fnIMSUM,
  IMSUB: fnIMSUB,
  IMPRODUCT: fnIMPRODUCT,
  IMDIV: fnIMDIV,
  IMPOWER: fnIMPOWER,
  IMSQRT: fnIMSQRT,
  IMLN: fnIMLN,
  IMLOG2: fnIMLOG2,
  IMLOG10: fnIMLOG10,
  IMEXP: fnIMEXP,
  IMSIN: fnIMSIN,
  IMCOS: fnIMCOS,
  BITAND: fnBITAND,
  "_XLFN.BITAND": fnBITAND,
  BITOR: fnBITOR,
  "_XLFN.BITOR": fnBITOR,
  BITXOR: fnBITXOR,
  "_XLFN.BITXOR": fnBITXOR,
  BITLSHIFT: fnBITLSHIFT,
  "_XLFN.BITLSHIFT": fnBITLSHIFT,
  BITRSHIFT: fnBITRSHIFT,
  "_XLFN.BITRSHIFT": fnBITRSHIFT,

  // Information
  "ERROR.TYPE": fnERROR_TYPE,
  INFO: fnINFO_STUB,
  CELL: fnCELL_STUB,
  ISREF: fnISREF_STUB,
  ISEVEN: fnISEVEN,
  "_XLFN.ISEVEN": fnISEVEN,
  ISODD: fnISODD,
  "_XLFN.ISODD": fnISODD,
  NA: fnNA,
  SHEET: fnSHEET_STUB,
  "_XLFN.SHEET": fnSHEET_STUB,
  SHEETS: fnSHEETS_STUB,
  "_XLFN.SHEETS": fnSHEETS_STUB,

  // Lambda Array Functions
  MAP: fnMAP,
  "_XLFN.MAP": fnMAP,
  REDUCE: fnREDUCE,
  "_XLFN.REDUCE": fnREDUCE,
  SCAN: fnSCAN,
  "_XLFN.SCAN": fnSCAN,
  MAKEARRAY: fnMAKEARRAY,
  "_XLFN.MAKEARRAY": fnMAKEARRAY,
  BYROW: fnBYROW,
  "_XLFN.BYROW": fnBYROW,
  BYCOL: fnBYCOL,
  "_XLFN.BYCOL": fnBYCOL,

  // Excel 365 Text Functions
  TEXTBEFORE: fnTEXTBEFORE,
  "_XLFN.TEXTBEFORE": fnTEXTBEFORE,
  TEXTAFTER: fnTEXTAFTER,
  "_XLFN.TEXTAFTER": fnTEXTAFTER,
  TEXTSPLIT: fnTEXTSPLIT,
  "_XLFN.TEXTSPLIT": fnTEXTSPLIT
};
