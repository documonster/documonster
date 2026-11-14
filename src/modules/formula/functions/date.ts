/**
 * Date / Time Functions — Native RuntimeValue implementation.
 *
 * Dates are read as calendar fields of Excel's own calendar (`toParts`) and built back with
 * `partsToSerial`; weekdays come from the serial (`weekdayOf`). No function here goes through a
 * `Date`, which holds the real calendar: it has no 1900-02-29 (serial 60), no 1900-01-00
 * (serial 0), and calls 1900-01-01 a Monday where Excel says Sunday.
 *
 * The exception is the clock: `TODAY` / `NOW` read the host's local wall-clock fields, which is
 * what the user reads too, and pack them into a serial with no timezone.
 */

import { argToNumber, checkError } from "@formula/functions/_shared";
import { calendarParts, yearFraction } from "@formula/functions/year-fraction";
import type { FunctionRuntimeContext } from "@formula/runtime/function-context";
import { DEFAULT_FUNCTION_CONTEXT } from "@formula/runtime/function-context";
import type { RuntimeValue, ErrorValue, NumberValue } from "@formula/runtime/values";
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
  rvBoolean
} from "@formula/runtime/values";
import type { ExcelDateTimeParts } from "@utils/excel-serial";
import {
  excelDaysInMonth,
  gregorianDaysInMonth,
  partsToSerial,
  serialToWeekday
} from "@utils/excel-serial";

// ============================================================================
// Type alias for native function signature
// ============================================================================

// ============================================================================
// Helpers
// ============================================================================

/**
 * An Excel serial as calendar fields on Excel's own calendar, honouring the active date1904 mode.
 *
 * Every date function reads fields and does calendar arithmetic through these parts and
 * `partsToSerial`, never through a `Date`: a `Date` holds the real calendar, which has no
 * 1900-02-29 (serial 60) and no 1900-01-00 (serial 0). The parts carry both as plain numbers —
 * `MONTH(60)` is 2 and `DAY(0)` is 0, as in Excel.
 */
function toParts(serial: number, context: FunctionRuntimeContext): ExcelDateTimeParts {
  return calendarParts(serial, context.date1904);
}

/** Whether a serial lies in Excel's date range: day 0 through 9999-12-31. */
function isSerialInRange(serial: number, context: FunctionRuntimeContext): boolean {
  return serial >= 0 && serial < partsToSerial({ year: 10000, month: 1, day: 1 }, context.date1904);
}

/**
 * A date or time argument as a number in Excel's date range, `#NUM!` outside it — what Excel
 * answers for a serial "out of range for the current date base value" (`WEEKDAY(-1)`,
 * `YEAR(2958466)`), rather than reading a calendar field off a day it has no date for.
 */
function dateArg(arg: RuntimeValue, context: FunctionRuntimeContext): NumberValue | ErrorValue {
  const n = argToNumber(arg);
  if (isError(n)) {
    return n;
  }
  return isSerialInRange(n.value, context) ? n : ERRORS.NUM;
}

/**
 * The day of the week of a serial, 0 = Sunday … 6 = Saturday.
 *
 * Every weekday question goes through here rather than a `Date`'s `getUTCDay()`: a `Date` holds
 * the real calendar, on which 1900-01-01 was a Monday, while Excel — counting a 1900-02-29 that
 * never existed — calls serial 1 a Sunday, and `TEXT(1, "dddd")` already said so.
 */
function weekdayOf(serial: number, context: FunctionRuntimeContext): number {
  return serialToWeekday(serial, context.date1904);
}

/** The serial of January 1st of `year`, honouring the active date1904 mode. */
function januaryFirst(year: number, context: FunctionRuntimeContext): number {
  return partsToSerial({ year, month: 1, day: 1 }, context.date1904);
}

/**
 * Collect holiday serial numbers from a RuntimeValue argument (array or scalar).
 *
 * A holiday is a date like the others: one out of Excel's date range is `#NUM!`, as Microsoft
 * documents for `WORKDAY.INTL`, rather than a number quietly matched against no day.
 */
function collectHolidays(
  arg: RuntimeValue,
  context: FunctionRuntimeContext
): Set<number> | ErrorValue {
  const set = new Set<number>();
  if (isArray(arg)) {
    for (const row of arg.rows) {
      for (const cell of row) {
        // Propagate errors from the holidays list rather than silently
        // skipping them — Excel surfaces `#N/A` from a holiday cell.
        if (cell.kind === RVKind.Error) {
          return cell;
        }
        if (cell.kind === RVKind.Number) {
          if (!isSerialInRange(cell.value, context)) {
            return ERRORS.NUM;
          }
          set.add(Math.floor(cell.value));
        }
      }
    }
  } else {
    if (arg.kind === RVKind.Error) {
      return arg;
    }
    const n = toNumberRV(arg);
    if (n.kind === RVKind.Number) {
      if (!isSerialInRange(n.value, context)) {
        return ERRORS.NUM;
      }
      set.add(Math.floor(n.value));
    }
  }
  return set;
}

// ============================================================================
// Date Functions
// ============================================================================

/**
 * TODAY — today's date on the host's wall clock.
 *
 * The user's "today" is their local date, so the local fields of `new Date()` are read and packed
 * into a serial, which has no timezone.
 */
export function fnTODAY(
  _args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const now = new Date();
  return rvNumber(
    partsToSerial(
      { year: now.getFullYear(), month: now.getMonth() + 1, day: now.getDate() },
      context.date1904
    )
  );
}

/**
 * NOW — the current date and time on the host's wall clock, as Excel reports it.
 *
 * Read from the local fields, like `TODAY`: the UTC instant made `NOW()` disagree with `TODAY()`
 * east or west of UTC — in UTC+8 at 07:00 it was still yesterday evening.
 */
export function fnNOW(
  _args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const now = new Date();
  return rvNumber(
    partsToSerial(
      {
        year: now.getFullYear(),
        month: now.getMonth() + 1,
        day: now.getDate(),
        hour: now.getHours(),
        minute: now.getMinutes(),
        second: now.getSeconds(),
        millisecond: now.getMilliseconds()
      },
      context.date1904
    )
  );
}

export function fnYEAR(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const n = dateArg(args[0], context);
  if (isError(n)) {
    return n;
  }
  return rvNumber(toParts(n.value, context).year);
}

export function fnMONTH(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const n = dateArg(args[0], context);
  if (isError(n)) {
    return n;
  }
  return rvNumber(toParts(n.value, context).month);
}

export function fnDAY(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const n = dateArg(args[0], context);
  if (isError(n)) {
    return n;
  }
  return rvNumber(toParts(n.value, context).day);
}

export function fnDATE(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const year = argToNumber(args[0]);
  if (isError(year)) {
    return year;
  }
  const month = argToNumber(args[1]);
  if (isError(month)) {
    return month;
  }
  const day = argToNumber(args[2]);
  if (isError(day)) {
    return day;
  }

  // Excel's DATE interprets a year in [0, 1899] as (year + 1900). Years below 0
  // or above 9999 are rejected as #NUM! (Excel's documented range).
  let y = Math.trunc(year.value);
  if (y < 0 || y > 9999) {
    return ERRORS.NUM;
  }
  if (y < 1900) {
    // After this coercion `y` is in [1900, 3799], always within Excel's [1900, 9999] range.
    y += 1900;
  }
  // Month and day carry over Excel's own calendar (`DATE(2024,13,1)` is 2025-01-01, and
  // `DATE(1900,2,29)` / `DATE(1900,3,0)` are both serial 60, the 1900-02-29 Excel counts).
  const serial = partsToSerial(
    { year: y, month: Math.trunc(month.value), day: Math.trunc(day.value) },
    context.date1904
  );
  return isSerialInRange(serial, context) ? rvNumber(serial) : ERRORS.NUM;
}

export function fnTIME(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const hour = argToNumber(args[0]);
  if (isError(hour)) {
    return hour;
  }
  const minute = argToNumber(args[1]);
  if (isError(minute)) {
    return minute;
  }
  const second = argToNumber(args[2]);
  if (isError(second)) {
    return second;
  }
  // Excel's TIME rejects negative arguments outright and wraps anything >= 24
  // hours back into the [0, 1) fraction-of-day range. Without the modulo,
  // `TIME(25, 0, 0)` would produce a value > 1, which breaks downstream
  // date-time arithmetic that expects a pure time fraction.
  if (hour.value < 0 || minute.value < 0 || second.value < 0) {
    return ERRORS.NUM;
  }
  const total = (hour.value * 3600 + minute.value * 60 + second.value) / 86400;
  // total could still be >= 1 if e.g. hour=25; wrap into [0, 1).
  return rvNumber(total - Math.floor(total));
}

export function fnHOUR(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const n = dateArg(args[0], context);
  if (isError(n)) {
    return n;
  }
  const totalSeconds = Math.round((n.value % 1) * 86400);
  return rvNumber(Math.floor(totalSeconds / 3600) % 24);
}

export function fnMINUTE(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const n = dateArg(args[0], context);
  if (isError(n)) {
    return n;
  }
  const totalSeconds = Math.round((n.value % 1) * 86400);
  return rvNumber(Math.floor(totalSeconds / 60) % 60);
}

export function fnSECOND(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const n = dateArg(args[0], context);
  if (isError(n)) {
    return n;
  }
  const totalSeconds = Math.round((n.value % 1) * 86400);
  return rvNumber(totalSeconds % 60);
}

export function fnWEEKDAY(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const n = dateArg(args[0], context);
  if (isError(n)) {
    return n;
  }
  // Blank `return_type` → Excel default 1 (Sun=1..Sat=7). Without the
  // blank guard, `argToNumber(BLANK)` coerces to 0 which falls to the
  // default branch and yields a spurious #NUM! for `WEEKDAY(date, )`.
  const returnType =
    args.length > 1 && args[1].kind !== RVKind.Blank ? argToNumber(args[1]) : rvNumber(1);
  if (isError(returnType)) {
    return returnType;
  }
  const day = weekdayOf(n.value, context); // 0=Sun, 6=Sat
  // Excel truncates the code: `WEEKDAY(d, 2.5)` is type 2.
  switch (Math.trunc(returnType.value)) {
    case 1:
      return rvNumber(day + 1); // 1=Sun, 7=Sat
    case 2:
      return rvNumber(day === 0 ? 7 : day); // 1=Mon, 7=Sun
    case 3:
      return rvNumber(day === 0 ? 6 : day - 1); // 0=Mon, 6=Sun
    case 11: // Mon=1..Sun=7
    case 12: // Tue=1..Mon=7
    case 13: // Wed=1..Tue=7
    case 14: // Thu=1..Wed=7
    case 15: // Fri=1..Thu=7
    case 16: // Sat=1..Fri=7
    case 17: // Sun=1..Sat=7
      return rvNumber(((((day - (Math.trunc(returnType.value) - 10)) % 7) + 7) % 7) + 1);
    default:
      return ERRORS.NUM;
  }
}

export function fnEOMONTH(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const startDate = dateArg(args[0], context);
  if (isError(startDate)) {
    return startDate;
  }
  const months = argToNumber(args[1]);
  if (isError(months)) {
    return months;
  }
  // Excel truncates `months` toward zero before doing month arithmetic.
  const m = Math.trunc(months.value);
  const { year, month } = toParts(startDate.value, context);
  // The last day of the target month on the real calendar: Excel's `EOMONTH(1, 1)` is 59.
  const target = month + m;
  const result = partsToSerial(
    { year, month: target, day: gregorianDaysInMonth(year, target) },
    context.date1904
  );
  return isSerialInRange(result, context) ? rvNumber(result) : ERRORS.NUM;
}

export function fnEDATE(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const startDate = dateArg(args[0], context);
  if (isError(startDate)) {
    return startDate;
  }
  const months = argToNumber(args[1]);
  if (isError(months)) {
    return months;
  }
  const m = Math.trunc(months.value);
  const { year, month, day } = toParts(startDate.value, context);
  // Excel clamps to the last day of the target month when the source day would overflow
  // (`EDATE(2024-01-31, 1)` is 2024-02-29) — of the real calendar: `EDATE(1900-01-31, 1)` is 59.
  const lastDay = gregorianDaysInMonth(year, month + m);
  const result = partsToSerial(
    { year, month: month + m, day: Math.min(day, lastDay) },
    context.date1904
  );
  return isSerialInRange(result, context) ? rvNumber(result) : ERRORS.NUM;
}

export function fnDATEDIF(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const startN = dateArg(args[0], context);
  if (isError(startN)) {
    return startN;
  }
  const endN = dateArg(args[1], context);
  if (isError(endN)) {
    return endN;
  }
  // DATEDIF requires end >= start; otherwise #NUM! per Excel.
  if (endN.value < startN.value) {
    return ERRORS.NUM;
  }
  const unit = toStringRV(topLeft(args[2])).toUpperCase();
  const start = toParts(startN.value, context);
  const end = toParts(endN.value, context);
  const sy = start.year;
  const sm = start.month;
  const sd = start.day;
  const ey = end.year;
  const em = end.month;
  const ed = end.day;
  switch (unit) {
    case "Y":
      return rvNumber(ey - sy - (em < sm || (em === sm && ed < sd) ? 1 : 0));
    case "M":
      return rvNumber((ey - sy) * 12 + em - sm - (ed < sd ? 1 : 0));
    case "D":
      return rvNumber(Math.floor(endN.value) - Math.floor(startN.value));
    case "MD": {
      // Days between the dates, ignoring months and years.
      // If end.day >= start.day → ed - sd; otherwise borrow days from the
      // previous month of the end date (its length - start.day + end.day).
      if (ed >= sd) {
        return rvNumber(ed - sd);
      }
      return rvNumber(excelDaysInMonth(ey, em - 1, context.date1904) - sd + ed);
    }
    case "YM": {
      // Months between the dates, ignoring days and years.
      let months = em - sm;
      if (ed < sd) {
        months -= 1;
      }
      if (months < 0) {
        months += 12;
      }
      return rvNumber(months);
    }
    case "YD": {
      // Days between the dates as though they were in the same year, ignoring years.
      // Align the end date to the start's year (or the next one if it falls earlier in the year).
      const startSerial = partsToSerial({ year: sy, month: sm, day: sd }, context.date1904);
      let diff = partsToSerial({ year: sy, month: em, day: ed }, context.date1904) - startSerial;
      if (diff < 0) {
        diff = partsToSerial({ year: sy + 1, month: em, day: ed }, context.date1904) - startSerial;
      }
      return rvNumber(diff);
    }
    default:
      return ERRORS.NUM;
  }
}

export function fnDAYS(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const end = dateArg(args[0], context);
  if (isError(end)) {
    return end;
  }
  const start = dateArg(args[1], context);
  if (isError(start)) {
    return start;
  }
  return rvNumber(Math.floor(end.value) - Math.floor(start.value));
}

export function fnISOWEEKNUM(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const n = dateArg(args[0], context);
  if (isError(n)) {
    return n;
  }
  // ISO 8601: a week belongs to the year that holds its Thursday. Worked on serials so the
  // weekday is Excel's (see `weekdayOf`).
  const day = Math.floor(n.value);
  const mondayBased = (weekdayOf(day, context) + 6) % 7; // 0=Mon … 6=Sun
  const thursday = day - mondayBased + 3;
  const year = toParts(thursday, context).year;
  return rvNumber(Math.floor((thursday - januaryFirst(year, context)) / 7) + 1);
}

export function fnWEEKNUM(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const n = dateArg(args[0], context);
  if (isError(n)) {
    return n;
  }
  // Blank `return_type` → Excel default 1 (Sunday start). See WEEKDAY
  // for the same rationale.
  const returnType =
    args.length > 1 && args[1].kind !== RVKind.Blank ? argToNumber(args[1]) : rvNumber(1);
  if (isError(returnType)) {
    return returnType;
  }
  // Truncated, as WEEKDAY's: `WEEKNUM(d, 2.5)` is type 2.
  const rt = Math.trunc(returnType.value);
  // Type 21 is ISO 8601 week.
  if (rt === 21) {
    return fnISOWEEKNUM(args, context);
  }
  // Excel maps `return_type` to the day-of-week that starts the week:
  //   1  (default) → Sunday
  //   2  or 11     → Monday
  //   12 → Tuesday, 13 → Wednesday, … 17 → Saturday
  //   (16 → Friday, 17 → Saturday)
  // Any other value is #NUM!.
  let startDay: number; // 0 = Sunday … 6 = Saturday
  switch (rt) {
    case 1:
      startDay = 0;
      break;
    case 2:
    case 11:
      startDay = 1;
      break;
    case 12:
      startDay = 2;
      break;
    case 13:
      startDay = 3;
      break;
    case 14:
      startDay = 4;
      break;
    case 15:
      startDay = 5;
      break;
    case 16:
      startDay = 6;
      break;
    case 17:
      startDay = 0;
      break;
    default:
      return ERRORS.NUM;
  }
  const day = Math.floor(n.value);
  const jan1 = januaryFirst(toParts(day, context).year, context);
  const jan1Day = weekdayOf(jan1, context);
  const dayOfYear = day - jan1;
  return rvNumber(Math.floor((dayOfYear + ((jan1Day - startDay + 7) % 7)) / 7) + 1);
}

function networkdaysHelper(
  startN: number,
  endN: number,
  holidays: Set<number>,
  context: FunctionRuntimeContext
): number {
  const s = Math.floor(Math.min(startN, endN));
  const e = Math.floor(Math.max(startN, endN));
  const sign = startN <= endN ? 1 : -1;
  // Closed-form weekday count: partition `[s, e]` into whole weeks plus
  // a tail. Each whole week contributes 5 weekdays, regardless of its
  // starting day-of-week. The tail contributes however many of its
  // remaining days fall on Monday..Friday.
  //
  // `weekdayOf`: Sun=0, Mon=1, …, Sat=6. We compute `dow` once for
  // the start date and then just walk the `tail` days forward by
  // modular arithmetic.
  const totalDays = e - s + 1;
  const weeks = Math.floor(totalDays / 7);
  const tail = totalDays % 7;
  let weekdays = weeks * 5;
  if (tail > 0) {
    const startDow = weekdayOf(s, context);
    for (let i = 0; i < tail; i++) {
      const dow = (startDow + i) % 7;
      if (dow !== 0 && dow !== 6) {
        weekdays++;
      }
    }
  } else {
    // When `totalDays` is an exact multiple of 7 the start DOW still
    // governs whether any holidays from the caller's list fall on a
    // weekday, so we stop here — no tail to walk.
  }
  // Subtract holidays that land on a weekday and fall within [s, e].
  if (holidays.size > 0) {
    for (const h of holidays) {
      if (h >= s && h <= e) {
        const dow = weekdayOf(h, context);
        if (dow !== 0 && dow !== 6) {
          weekdays--;
        }
      }
    }
  }
  return weekdays * sign;
}

export function fnNETWORKDAYS(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const startN = dateArg(args[0], context);
  if (isError(startN)) {
    return startN;
  }
  const endN = dateArg(args[1], context);
  if (isError(endN)) {
    return endN;
  }
  const holidays = args.length > 2 ? collectHolidays(args[2], context) : new Set<number>();
  if (holidays instanceof Set) {
    return rvNumber(networkdaysHelper(startN.value, endN.value, holidays, context));
  }
  return holidays;
}

export function fnWORKDAY(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const startN = dateArg(args[0], context);
  if (isError(startN)) {
    return startN;
  }
  const days = argToNumber(args[1]);
  if (isError(days)) {
    return days;
  }
  const holidays = args.length > 2 ? collectHolidays(args[2], context) : new Set<number>();
  if (!(holidays instanceof Set)) {
    return holidays;
  }
  let current = Math.floor(startN.value);
  // Excel truncates `days` toward zero. Without this, a fractional input
  // like 2.7 would walk extra iterations until the fractional remainder
  // underflowed past zero, silently producing a wrong result.
  const daysInt = Math.trunc(days.value);
  // Each working day is at least one calendar day further on, so when `days` alone reaches past
  // Excel's dates the answer does too: refuse it rather than walk there a day at a time.
  if (!isSerialInRange(current + daysInt, context)) {
    return ERRORS.NUM;
  }
  const step = daysInt >= 0 ? 1 : -1;
  let remaining = Math.abs(daysInt);
  while (remaining > 0) {
    current += step;
    const dow = weekdayOf(current, context);
    if (dow !== 0 && dow !== 6 && !holidays.has(current)) {
      remaining--;
    }
  }
  return isSerialInRange(current, context) ? rvNumber(current) : ERRORS.NUM;
}

export function fnYEARFRAC(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const startN = dateArg(args[0], context);
  if (isError(startN)) {
    return startN;
  }
  const endN = dateArg(args[1], context);
  if (isError(endN)) {
    return endN;
  }
  const basis = args.length > 2 ? argToNumber(args[2]) : rvNumber(0);
  if (isError(basis)) {
    return basis;
  }
  // Basis is truncated (`YEARFRAC(a, b, 1.9)` is basis 1); anything outside 0–4 is #NUM!.
  const b = Math.trunc(basis.value);
  if (b < 0 || b > 4) {
    return ERRORS.NUM;
  }
  return rvNumber(
    yearFraction(
      Math.min(startN.value, endN.value),
      Math.max(startN.value, endN.value),
      b,
      context.date1904
    )
  );
}

/**
 * Whether a serial is the last day of February on Excel's calendar (1900-02-29 included, 02-28
 * not), which is the calendar `DAYS360` uses — unlike `YEARFRAC`'s basis 0.
 */
function isLastOfFebruary(serial: number, context: FunctionRuntimeContext): boolean {
  return (
    toParts(serial, context).month === 2 && toParts(Math.floor(serial) + 1, context).month !== 2
  );
}

export function fnDATEVALUE(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0])).trim();
  const parsed = parseDateOnly(text);
  if (!parsed) {
    return ERRORS.VALUE;
  }
  // Excel's calendar, so "2/29/1900" is serial 60 like any other valid date.
  const serial = partsToSerial(
    { year: parsed.y, month: parsed.m, day: parsed.d },
    context.date1904
  );
  // A date before the epoch (1900-01-01, or 1904-01-01 in the 1904 system) is not a date to Excel.
  return serial >= 1 || (context.date1904 && serial >= 0) ? rvNumber(serial) : ERRORS.VALUE;
}

export function fnTIMEVALUE(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const text = toStringRV(topLeft(args[0])).trim();
  const parsed = parseTimeOnly(text);
  if (parsed === null) {
    return ERRORS.VALUE;
  }
  return rvNumber(parsed);
}

// ============================================================================
// Date / Time parsing (deterministic, independent of host locale)
//
// `new Date(text)` is unreliable across engines (Chrome parses "1/2/3" as
// US MDY, Node varies by version, and locale influences both). We hand-
// roll a small parser that accepts the formats Excel's DATEVALUE does and
// rejects everything else.
// ============================================================================

const MONTH_NAMES: Record<string, number> = {
  jan: 1,
  january: 1,
  feb: 2,
  february: 2,
  mar: 3,
  march: 3,
  apr: 4,
  april: 4,
  may: 5,
  jun: 6,
  june: 6,
  jul: 7,
  july: 7,
  aug: 8,
  august: 8,
  sep: 9,
  sept: 9,
  september: 9,
  oct: 10,
  october: 10,
  nov: 11,
  november: 11,
  dec: 12,
  december: 12
};

/** Parse a date-only string into `{y, m, d}` or `null` on rejection. */
function parseDateOnly(raw: string): { y: number; m: number; d: number } | null {
  const s = raw.trim();
  // Strip a trailing time component (e.g. "2024-01-15 14:30") — but only
  // split at a space that is followed by a digit, so a spelled-out date
  // like "Jan 15, 2024" (which has month-name and day-number separated
  // by spaces) survives for the "Mmm D, YYYY" matcher below.
  // A naive `indexOf(" ")` truncated "Jan 15, 2024" to just "Jan".
  const trailingTime = /^(.+?)\s(\d{1,2}:[\d:]+(?:\s*[AaPp]\.?[Mm]\.?)?)$/.exec(s);
  const datePart = trailingTime ? trailingTime[1].trim() : s;

  // ISO YYYY-MM-DD or YYYY/MM/DD
  let m = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(datePart);
  if (m) {
    return validateYmd(+m[1], +m[2], +m[3]);
  }
  // US M/D/YYYY or M-D-YYYY
  m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/.exec(datePart);
  if (m) {
    let y = +m[3];
    if (y < 100) {
      // Excel's pivot: 00-29 → 2000s, 30-99 → 1900s
      y += y < 30 ? 2000 : 1900;
    }
    return validateYmd(y, +m[1], +m[2]);
  }
  // D-Mmm-YY or D-Mmm-YYYY
  m = /^(\d{1,2})[ /-]([A-Za-z]{3,9})[ /-]?(\d{2,4})?$/.exec(datePart);
  if (m) {
    const month = MONTH_NAMES[m[2].toLowerCase()];
    if (!month) {
      return null;
    }
    // Excel's DATEVALUE substitutes the host's current calendar year when
    // the input omits a year (e.g. "15-Jan"). This matches Excel on the
    // desktop, but note that the return value is not stable across time
    // zones or years — tests that exercise this branch should freeze the
    // clock (or supply a year) if they need reproducibility.
    let y = m[3] ? +m[3] : new Date().getFullYear();
    if (y < 100) {
      y += y < 30 ? 2000 : 1900;
    }
    return validateYmd(y, month, +m[1]);
  }
  // "Mmm D, YYYY" or "Month D, YYYY"
  m = /^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{2,4})$/.exec(datePart);
  if (m) {
    const month = MONTH_NAMES[m[1].toLowerCase()];
    if (!month) {
      return null;
    }
    let y = +m[3];
    if (y < 100) {
      y += y < 30 ? 2000 : 1900;
    }
    return validateYmd(y, month, +m[2]);
  }
  return null;
}

/** Validate a calendar date and return null for out-of-range components. */
function validateYmd(y: number, mo: number, d: number): { y: number; m: number; d: number } | null {
  if (y < 0 || y > 9999 || mo < 1 || mo > 12 || d < 1 || d > 31) {
    return null;
  }
  // Day-of-month range check on Excel's calendar, where 1900 is a leap year. Validated against the
  // 1900 calendar whatever the epoch: under 1904 every 1900 date is before the epoch anyway.
  if (d > excelDaysInMonth(y, mo)) {
    return null;
  }
  return { y, m: mo, d };
}

/** Parse a time-only string into a fraction-of-day in [0, 1). */
function parseTimeOnly(raw: string): number | null {
  // Optional AM/PM suffix; captured case-insensitively.
  const m = /^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}(?:\.\d+)?))?(?:\s*([AaPp])\.?\s*[Mm]\.?)?$/.exec(
    raw
  );
  if (!m) {
    return null;
  }
  let h = +m[1];
  const min = m[2] ? +m[2] : 0;
  const sec = m[3] ? +m[3] : 0;
  if (min >= 60 || sec >= 60 || min < 0 || sec < 0) {
    return null;
  }
  if (m[4]) {
    const pm = m[4].toLowerCase() === "p";
    if (h < 1 || h > 12) {
      return null;
    }
    if (pm && h < 12) {
      h += 12;
    } else if (!pm && h === 12) {
      h = 0;
    }
  } else {
    if (h < 0 || h > 23) {
      return null;
    }
  }
  return (h * 3600 + min * 60 + sec) / 86400;
}

// ============================================================================
// More Date/Time Functions
// ============================================================================

export function fnDAYS360(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const startN = dateArg(args[0], context);
  if (isError(startN)) {
    return startN;
  }
  const endN = dateArg(args[1], context);
  if (isError(endN)) {
    return endN;
  }
  const methodRV = args.length > 2 ? toBooleanRV(topLeft(args[2])) : rvBoolean(false);
  if (isError(methodRV)) {
    return methodRV;
  }
  const method = methodRV.value;
  // Calendar fields from the serial (`toParts`), so 1900-02-29 — Excel's serial 60, which a
  // `Date` turns into March 1st — is read as the last day of February it is to Excel.
  const start = Math.floor(startN.value);
  const end = Math.floor(endN.value);
  const sd = toParts(start, context);
  const ed = toParts(end, context);
  let d1 = sd.day;
  let d2 = ed.day;
  const m1 = sd.month;
  const m2 = ed.month;
  const y1 = sd.year;
  const y2 = ed.year;
  if (method) {
    // European: the 31st is the 30th; February is left alone.
    if (d1 === 31) {
      d1 = 30;
    }
    if (d2 === 31) {
      d2 = 30;
    }
  } else {
    // US (NASD): a start date on the last day of its month — February's included — is the
    // 30th, and an end date on the 31st becomes the 30th only when the start is then the 30th.
    // Unlike YEARFRAC's basis 0, an end date on the last of February is not moved.
    if (d1 === 31 || isLastOfFebruary(start, context)) {
      d1 = 30;
    }
    if (d2 === 31 && d1 >= 30) {
      d2 = 30;
    }
  }
  return rvNumber((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1));
}

/** Weekend days (0=Sun..6=Sat) for the numeric `weekend` codes of the `*.INTL` functions. */
const WEEKEND_CODES: ReadonlyMap<number, readonly number[]> = new Map([
  [1, [6, 0]], // Sat, Sun
  [2, [0, 1]], // Sun, Mon
  [3, [1, 2]], // Mon, Tue
  [4, [2, 3]], // Tue, Wed
  [5, [3, 4]], // Wed, Thu
  [6, [4, 5]], // Thu, Fri
  [7, [5, 6]], // Fri, Sat
  [11, [0]], // Sun only
  [12, [1]], // Mon only
  [13, [2]], // Tue only
  [14, [3]], // Wed only
  [15, [4]], // Thu only
  [16, [5]], // Fri only
  [17, [6]] // Sat only
]);

/**
 * The `weekend` argument of `NETWORKDAYS.INTL` / `WORKDAY.INTL` as the set of weekend days
 * (0=Sun..6=Sat).
 *
 * Excel takes either a code (1–7, 11–17, truncated; anything else is `#NUM!`) or a seven-character mask of
 * `0`/`1` for Monday through Sunday, `1` marking a weekend day (any other string is `#VALUE!`).
 * The mask is a string, never a number: read as one, `"0000011"` became code 11 (Sunday only)
 * and `"1000000"` an unknown code that quietly fell back to Saturday and Sunday. Omitted or blank
 * means code 1.
 */
function parseWeekend(args: RuntimeValue[], index: number): Set<number> | ErrorValue {
  const arg = args.length > index ? topLeft(args[index]) : undefined;
  if (arg === undefined || arg.kind === RVKind.Blank) {
    return new Set(WEEKEND_CODES.get(1));
  }
  if (arg.kind === RVKind.Error) {
    return arg;
  }
  if (arg.kind === RVKind.String) {
    const mask = arg.value;
    if (!/^[01]{7}$/.test(mask)) {
      return ERRORS.VALUE;
    }
    const days = new Set<number>();
    for (let i = 0; i < 7; i++) {
      if (mask[i] === "1") {
        days.add((i + 1) % 7); // mask position 0 is Monday
      }
    }
    return days;
  }
  const code = toNumberRV(arg);
  if (isError(code)) {
    return code;
  }
  // A code is truncated, as Excel does: 1.6 is code 1, 17.5 is 17, and 0.5 is the invalid 0.
  const days = WEEKEND_CODES.get(Math.trunc(code.value));
  return days ? new Set(days) : ERRORS.NUM;
}

export function fnNETWORKDAYS_INTL(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const startN = dateArg(args[0], context);
  if (isError(startN)) {
    return startN;
  }
  const endN = dateArg(args[1], context);
  if (isError(endN)) {
    return endN;
  }
  const weekendDays = parseWeekend(args, 2);
  if (!(weekendDays instanceof Set)) {
    return weekendDays;
  }
  const holidays = args.length > 3 ? collectHolidays(args[3], context) : new Set<number>();
  if (!(holidays instanceof Set)) {
    return holidays;
  }
  const s = Math.floor(Math.min(startN.value, endN.value));
  const e = Math.floor(Math.max(startN.value, endN.value));
  const sign = startN.value <= endN.value ? 1 : -1;
  let count = 0;
  for (let d = s; d <= e; d++) {
    if (!weekendDays.has(weekdayOf(d, context)) && !holidays.has(d)) {
      count++;
    }
  }
  return rvNumber(count * sign);
}

export function fnWORKDAY_INTL(
  args: RuntimeValue[],
  context: FunctionRuntimeContext = DEFAULT_FUNCTION_CONTEXT
): RuntimeValue {
  const startN = dateArg(args[0], context);
  if (isError(startN)) {
    return startN;
  }
  const days = argToNumber(args[1]);
  if (isError(days)) {
    return days;
  }
  const weekendDays = parseWeekend(args, 2);
  if (!(weekendDays instanceof Set)) {
    return weekendDays;
  }
  // A week with no working day has no next working day: Excel refuses `"1111111"` here (while
  // NETWORKDAYS.INTL counts 0 with it), and stepping would never end.
  if (weekendDays.size === 7) {
    return ERRORS.VALUE;
  }
  const holidays = args.length > 3 ? collectHolidays(args[3], context) : new Set<number>();
  if (!(holidays instanceof Set)) {
    return holidays;
  }
  let current = Math.floor(startN.value);
  // Truncate `days` toward zero before stepping — see WORKDAY for the
  // same rationale. A fractional input like 2.7 would otherwise walk
  // extra iterations and silently produce the wrong result.
  const daysInt = Math.trunc(days.value);
  // See WORKDAY: the answer is at least `days` calendar days away.
  if (!isSerialInRange(current + daysInt, context)) {
    return ERRORS.NUM;
  }
  const step = daysInt >= 0 ? 1 : -1;
  let remaining = Math.abs(daysInt);
  while (remaining > 0) {
    current += step;
    if (!weekendDays.has(weekdayOf(current, context)) && !holidays.has(current)) {
      remaining--;
    }
  }
  return isSerialInRange(current, context) ? rvNumber(current) : ERRORS.NUM;
}
