/**
 * Date / Time Functions — Native RuntimeValue implementation.
 */

import { dateToExcel, excelToDate } from "@utils/utils.base";

import type { RuntimeValue, NumberValue, ErrorValue } from "../runtime/values";
import {
  RVKind,
  ERRORS,
  isError,
  isArray,
  toNumberRV,
  toStringRV,
  toBooleanRV,
  rvNumber,
  rvBoolean
} from "../runtime/values";

// ============================================================================
// Type alias for native function signature
// ============================================================================

type NativeFunction = (args: RuntimeValue[]) => RuntimeValue;

// ============================================================================
// Helpers
// ============================================================================

/** Extract a number from a RuntimeValue, returning the raw number or an ErrorValue. */
function numArg(v: RuntimeValue): NumberValue | ErrorValue {
  return toNumberRV(v);
}

/** Collect holiday serial numbers from a RuntimeValue argument (array or scalar). */
function collectHolidays(arg: RuntimeValue): Set<number> {
  const set = new Set<number>();
  if (isArray(arg)) {
    for (const row of arg.rows) {
      for (const cell of row) {
        if (cell.kind === RVKind.Number) {
          set.add(Math.floor(cell.value));
        }
      }
    }
  } else {
    const n = toNumberRV(arg);
    if (n.kind === RVKind.Number) {
      set.add(Math.floor(n.value));
    }
  }
  return set;
}

// ============================================================================
// Date Functions
// ============================================================================

export const fnTODAY: NativeFunction = () => {
  const now = new Date();
  return rvNumber(
    dateToExcel(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())))
  );
};

export const fnNOW: NativeFunction = () => {
  const now = new Date();
  return rvNumber(dateToExcel(now));
};

export const fnYEAR: NativeFunction = args => {
  const n = numArg(args[0]);
  if (isError(n)) {
    return n;
  }
  return rvNumber(excelToDate(n.value).getFullYear());
};

export const fnMONTH: NativeFunction = args => {
  const n = numArg(args[0]);
  if (isError(n)) {
    return n;
  }
  return rvNumber(excelToDate(n.value).getMonth() + 1);
};

export const fnDAY: NativeFunction = args => {
  const n = numArg(args[0]);
  if (isError(n)) {
    return n;
  }
  return rvNumber(excelToDate(n.value).getDate());
};

export const fnDATE: NativeFunction = args => {
  const year = numArg(args[0]);
  if (isError(year)) {
    return year;
  }
  const month = numArg(args[1]);
  if (isError(month)) {
    return month;
  }
  const day = numArg(args[2]);
  if (isError(day)) {
    return day;
  }

  // Lotus 1-2-3 bug: DATE(1900,2,29) should return serial 60 even though
  // 1900 is not a leap year. JavaScript's Date constructor rolls Feb 29, 1900
  // forward to March 1, 1900, so we handle this specially.
  if (year.value === 1900 && month.value === 2 && day.value === 29) {
    return rvNumber(60); // The fictitious Feb 29, 1900
  }

  const d = new Date(Date.UTC(year.value, month.value - 1, day.value));
  if (year.value >= 0 && year.value < 100) {
    d.setUTCFullYear(year.value);
  }
  return rvNumber(dateToExcel(d));
};

export const fnTIME: NativeFunction = args => {
  const hour = numArg(args[0]);
  if (isError(hour)) {
    return hour;
  }
  const minute = numArg(args[1]);
  if (isError(minute)) {
    return minute;
  }
  const second = numArg(args[2]);
  if (isError(second)) {
    return second;
  }
  return rvNumber((hour.value * 3600 + minute.value * 60 + second.value) / 86400);
};

export const fnHOUR: NativeFunction = args => {
  const n = numArg(args[0]);
  if (isError(n)) {
    return n;
  }
  const totalSeconds = Math.round((n.value % 1) * 86400);
  return rvNumber(Math.floor(totalSeconds / 3600) % 24);
};

export const fnMINUTE: NativeFunction = args => {
  const n = numArg(args[0]);
  if (isError(n)) {
    return n;
  }
  const totalSeconds = Math.round((n.value % 1) * 86400);
  return rvNumber(Math.floor(totalSeconds / 60) % 60);
};

export const fnSECOND: NativeFunction = args => {
  const n = numArg(args[0]);
  if (isError(n)) {
    return n;
  }
  const totalSeconds = Math.round((n.value % 1) * 86400);
  return rvNumber(totalSeconds % 60);
};

export const fnWEEKDAY: NativeFunction = args => {
  const n = numArg(args[0]);
  if (isError(n)) {
    return n;
  }
  const d = excelToDate(n.value);
  const returnType = args.length > 1 ? numArg(args[1]) : rvNumber(1);
  if (isError(returnType)) {
    return returnType;
  }
  const day = d.getDay(); // 0=Sun, 6=Sat
  switch (returnType.value) {
    case 1:
      return rvNumber(day + 1); // 1=Sun, 7=Sat
    case 2:
      return rvNumber(day === 0 ? 7 : day); // 1=Mon, 7=Sun
    case 3:
      return rvNumber(day === 0 ? 6 : day - 1); // 0=Mon, 6=Sun
    default:
      return rvNumber(day + 1);
  }
};

export const fnEOMONTH: NativeFunction = args => {
  const startDate = numArg(args[0]);
  if (isError(startDate)) {
    return startDate;
  }
  const months = numArg(args[1]);
  if (isError(months)) {
    return months;
  }
  const d = excelToDate(startDate.value);
  const result = new Date(Date.UTC(d.getFullYear(), d.getMonth() + months.value + 1, 0));
  return rvNumber(dateToExcel(result));
};

export const fnEDATE: NativeFunction = args => {
  const startDate = numArg(args[0]);
  if (isError(startDate)) {
    return startDate;
  }
  const months = numArg(args[1]);
  if (isError(months)) {
    return months;
  }
  const d = excelToDate(startDate.value);
  const result = new Date(Date.UTC(d.getFullYear(), d.getMonth() + months.value, d.getDate()));
  return rvNumber(dateToExcel(result));
};

export const fnDATEDIF: NativeFunction = args => {
  const startN = numArg(args[0]);
  if (isError(startN)) {
    return startN;
  }
  const endN = numArg(args[1]);
  if (isError(endN)) {
    return endN;
  }
  const unit = toStringRV(args[2]).toUpperCase();
  const startD = excelToDate(startN.value);
  const endD = excelToDate(endN.value);
  const sy = startD.getFullYear();
  const sm = startD.getMonth();
  const sd = startD.getDate();
  const ey = endD.getFullYear();
  const em = endD.getMonth();
  const ed = endD.getDate();
  switch (unit) {
    case "Y":
      return rvNumber(ey - sy - (em < sm || (em === sm && ed < sd) ? 1 : 0));
    case "M":
      return rvNumber((ey - sy) * 12 + em - sm - (ed < sd ? 1 : 0));
    case "D":
      return rvNumber(Math.floor((endD.getTime() - startD.getTime()) / 86400000));
    default:
      return ERRORS.NUM;
  }
};

export const fnDAYS: NativeFunction = args => {
  const end = numArg(args[0]);
  if (isError(end)) {
    return end;
  }
  const start = numArg(args[1]);
  if (isError(start)) {
    return start;
  }
  return rvNumber(Math.floor(end.value) - Math.floor(start.value));
};

export const fnISOWEEKNUM: NativeFunction = args => {
  const n = numArg(args[0]);
  if (isError(n)) {
    return n;
  }
  const d = excelToDate(n.value);
  const temp = new Date(d.getTime());
  temp.setDate(temp.getDate() + 3 - ((temp.getDay() + 6) % 7));
  const week1 = new Date(temp.getFullYear(), 0, 4);
  return rvNumber(
    1 +
      Math.round(
        ((temp.getTime() - week1.getTime()) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7
      )
  );
};

export const fnWEEKNUM: NativeFunction = args => {
  const n = numArg(args[0]);
  if (isError(n)) {
    return n;
  }
  const d = excelToDate(n.value);
  const returnType = args.length > 1 ? numArg(args[1]) : rvNumber(1);
  if (isError(returnType)) {
    return returnType;
  }
  // For returnType 21, use ISO week
  if (returnType.value === 21) {
    return fnISOWEEKNUM(args);
  }
  // Simple: week starts on Sunday (type 1) or Monday (type 2)
  const startDay = returnType.value === 2 ? 1 : 0; // 0=Sun, 1=Mon
  const jan1 = new Date(d.getFullYear(), 0, 1);
  const jan1Day = jan1.getDay();
  const dayOfYear = Math.floor((d.getTime() - jan1.getTime()) / 86400000);
  return rvNumber(Math.floor((dayOfYear + jan1Day - startDay + 7) / 7));
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

export const fnNETWORKDAYS: NativeFunction = args => {
  const startN = numArg(args[0]);
  if (isError(startN)) {
    return startN;
  }
  const endN = numArg(args[1]);
  if (isError(endN)) {
    return endN;
  }
  const holidays = args.length > 2 ? collectHolidays(args[2]) : new Set<number>();
  return rvNumber(networkdaysHelper(startN.value, endN.value, holidays));
};

export const fnWORKDAY: NativeFunction = args => {
  const startN = numArg(args[0]);
  if (isError(startN)) {
    return startN;
  }
  const days = numArg(args[1]);
  if (isError(days)) {
    return days;
  }
  const holidays = args.length > 2 ? collectHolidays(args[2]) : new Set<number>();
  let current = Math.floor(startN.value);
  const step = days.value >= 0 ? 1 : -1;
  let remaining = Math.abs(days.value);
  while (remaining > 0) {
    current += step;
    const dt = excelToDate(current);
    const dow = dt.getDay();
    if (dow !== 0 && dow !== 6 && !holidays.has(current)) {
      remaining--;
    }
  }
  return rvNumber(current);
};

export const fnYEARFRAC: NativeFunction = args => {
  const startN = numArg(args[0]);
  if (isError(startN)) {
    return startN;
  }
  const endN = numArg(args[1]);
  if (isError(endN)) {
    return endN;
  }
  const basis = args.length > 2 ? numArg(args[2]) : rvNumber(0);
  if (isError(basis)) {
    return basis;
  }
  const sd = excelToDate(Math.min(startN.value, endN.value));
  const ed = excelToDate(Math.max(startN.value, endN.value));
  const diffDays = Math.abs(Math.floor(endN.value) - Math.floor(startN.value));

  switch (basis.value) {
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
      return rvNumber(days30_360 / 360);
    }
    case 1: {
      // Actual/actual
      const y1 = sd.getFullYear();
      const y2 = ed.getFullYear();
      if (y1 === y2) {
        const yearDays =
          (new Date(y1 + 1, 0, 1).getTime() - new Date(y1, 0, 1).getTime()) / 86400000;
        return rvNumber(diffDays / yearDays);
      }
      // Span across years: average the year lengths
      const totalYearDays =
        (new Date(y2 + 1, 0, 1).getTime() - new Date(y1, 0, 1).getTime()) / 86400000;
      const avgYear = totalYearDays / (y2 - y1 + 1);
      return rvNumber(diffDays / avgYear);
    }
    case 2: // Actual/360
      return rvNumber(diffDays / 360);
    case 3: // Actual/365
      return rvNumber(diffDays / 365);
    case 4: {
      // European 30/360
      const d1 = Math.min(sd.getDate(), 30);
      const d2 = Math.min(ed.getDate(), 30);
      const m1 = sd.getMonth() + 1;
      const m2 = ed.getMonth() + 1;
      const y1 = sd.getFullYear();
      const y2 = ed.getFullYear();
      const days30_360 = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
      return rvNumber(days30_360 / 360);
    }
    default:
      return ERRORS.NUM;
  }
};

export const fnDATEVALUE: NativeFunction = args => {
  const text = toStringRV(args[0]);

  // Lotus 1-2-3 bug: "2/29/1900" or "February 29, 1900" etc. should return 60
  const lotus29 =
    /^(2[/-]29[/-]1900|1900[/-]2[/-]29|1900[/-]02[/-]29|02[/-]29[/-]1900|Feb(ruary)?\s+29[,]?\s+1900)$/i;
  if (lotus29.test(text.trim())) {
    return rvNumber(60);
  }

  const d = new Date(text);
  if (isNaN(d.getTime())) {
    return ERRORS.VALUE;
  }
  return rvNumber(dateToExcel(new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()))));
};

export const fnTIMEVALUE: NativeFunction = args => {
  const text = toStringRV(args[0]);
  const m = /(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(text);
  if (!m) {
    return ERRORS.VALUE;
  }
  const h = parseInt(m[1], 10);
  const min = parseInt(m[2], 10);
  const sec = m[3] ? parseInt(m[3], 10) : 0;
  return rvNumber((h * 3600 + min * 60 + sec) / 86400);
};

// ============================================================================
// More Date/Time Functions
// ============================================================================

export const fnDAYS360: NativeFunction = args => {
  const startN = numArg(args[0]);
  if (isError(startN)) {
    return startN;
  }
  const endN = numArg(args[1]);
  if (isError(endN)) {
    return endN;
  }
  const methodRV = args.length > 2 ? toBooleanRV(args[2]) : rvBoolean(false);
  if (isError(methodRV)) {
    return methodRV;
  }
  const method = methodRV.value;
  const sd = excelToDate(startN.value);
  const ed = excelToDate(endN.value);
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
  return rvNumber((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1));
};

export const fnNETWORKDAYS_INTL: NativeFunction = args => {
  const startN = numArg(args[0]);
  if (isError(startN)) {
    return startN;
  }
  const endN = numArg(args[1]);
  if (isError(endN)) {
    return endN;
  }
  const weekendArg = args.length > 2 ? numArg(args[2]) : rvNumber(1);
  if (isError(weekendArg)) {
    return weekendArg;
  }
  const holidays = args.length > 3 ? collectHolidays(args[3]) : new Set<number>();
  const weekendDays = new Set<number>();
  switch (weekendArg.value) {
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
  const s = Math.floor(Math.min(startN.value, endN.value));
  const e = Math.floor(Math.max(startN.value, endN.value));
  const sign = startN.value <= endN.value ? 1 : -1;
  let count = 0;
  for (let d = s; d <= e; d++) {
    const dt = excelToDate(d);
    if (!weekendDays.has(dt.getDay()) && !holidays.has(d)) {
      count++;
    }
  }
  return rvNumber(count * sign);
};

export const fnWORKDAY_INTL: NativeFunction = args => {
  const startN = numArg(args[0]);
  if (isError(startN)) {
    return startN;
  }
  const days = numArg(args[1]);
  if (isError(days)) {
    return days;
  }
  const weekendArg = args.length > 2 ? numArg(args[2]) : rvNumber(1);
  if (isError(weekendArg)) {
    return weekendArg;
  }
  const holidays = args.length > 3 ? collectHolidays(args[3]) : new Set<number>();
  const weekendDays = new Set<number>();
  switch (weekendArg.value) {
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
  let current = Math.floor(startN.value);
  const step = days.value >= 0 ? 1 : -1;
  let remaining = Math.abs(days.value);
  while (remaining > 0) {
    current += step;
    const dt = excelToDate(current);
    if (!weekendDays.has(dt.getDay()) && !holidays.has(current)) {
      remaining--;
    }
  }
  return rvNumber(current);
};
