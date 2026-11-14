/**
 * Day counts and year fractions for every `basis` Excel takes (0 US 30/360, 1 actual/actual,
 * 2 actual/360, 3 actual/365, 4 European 30/360), shared by `YEARFRAC` and the securities
 * functions.
 *
 * One implementation because Excel has one: measured against Excel for Mac 16.x, `DISC`,
 * `INTRATE`, `RECEIVED`, `PRICEDISC`, `YIELDDISC`, `ACCRINTM`, `PRICEMAT` and `YIELDMAT` divide by
 * exactly the year fraction `YEARFRAC` returns, on every basis — `DISC(2020-06-01, 2023-06-01, 97,
 * 100, 1)` is 0.03 / 2.997947 (1095 days over the average of 2020…2023), not ISDA's split into leap
 * and common days, and `ACCRINTM(2023-02-28, 2023-08-31, 0.1, 1000, 0)` is 50.28 because a start on
 * February's last day counts as the 30th.
 */

import type { ExcelDateTimeParts } from "@utils/excel-serial";
import {
  gregorianDaysInMonth,
  isGregorianLeapYear,
  partsToSerial,
  serialToParts
} from "@utils/excel-serial";

/**
 * An Excel serial as calendar fields on Excel's own calendar. Serial 0 of the 1900 system is
 * Excel's 1900-01-00, not the real 1899-12-31: `MONTH(0)` is 1 and `DAY(0)` is 0.
 */
export function calendarParts(serial: number, date1904?: boolean): ExcelDateTimeParts {
  const parts = serialToParts(serial, date1904);
  return !date1904 && Math.floor(serial) === 0 ? { ...parts, year: 1900, month: 1, day: 0 } : parts;
}

/** Whether the fields are February's last day of the real calendar (Excel's answers for 1900 show it is). */
function isLastOfFebruary({ year, month, day }: ExcelDateTimeParts): boolean {
  return month === 2 && day === gregorianDaysInMonth(year, 2);
}

/**
 * Days from `start` to `end` on the US (NASD) 30/360 calendar.
 *
 * A start on the 31st or on February's last day is the 30th. An end on the 31st is the 30th when
 * the start (as given) is the 30th or 31st, and an end on February's last day is the 30th when the
 * start is one too. `modifyBoth` applies both end rules unconditionally — the form Excel uses for
 * the length of a coupon period, where `COUPDAYSNC(2023-03-15, 2026-08-31, 2, 0)` is 180 - 15 = 165
 * and not the 166 a settlement-to-coupon count gives.
 */
export function days360Us(
  start: number,
  end: number,
  modifyBoth: boolean,
  date1904?: boolean
): number {
  const s = calendarParts(start, date1904);
  const e = calendarParts(end, date1904);
  let sd = s.day;
  let ed = e.day;
  if (isLastOfFebruary(e) && (isLastOfFebruary(s) || modifyBoth)) {
    ed = 30;
  }
  if (ed === 31 && (sd >= 30 || modifyBoth)) {
    ed = 30;
  }
  if (sd === 31 || isLastOfFebruary(s)) {
    sd = 30;
  }
  return (e.year - s.year) * 360 + (e.month - s.month) * 30 + (ed - sd);
}

/** Days from `start` to `end` on the European 30/360 calendar: the 31st is the 30th at either end. */
export function days360Eu(start: number, end: number, date1904?: boolean): number {
  const s = calendarParts(start, date1904);
  const e = calendarParts(end, date1904);
  return (
    (e.year - s.year) * 360 + (e.month - s.month) * 30 + (Math.min(e.day, 30) - Math.min(s.day, 30))
  );
}

/** The days between two serials (`start <= end`) as `basis` counts them: 30/360 for 0 and 4, actual otherwise. */
export function dayCount(start: number, end: number, basis: number, date1904?: boolean): number {
  const s = Math.floor(start);
  const e = Math.floor(end);
  switch (basis) {
    case 0:
      return days360Us(s, e, false, date1904);
    case 4:
      return days360Eu(s, e, date1904);
    default:
      return e - s;
  }
}

/**
 * The year length `basis` divides by for dates `start <= end`: 360 for 0, 2 and 4, 365 for 3, and
 * for 1 Excel's actual/actual rule — not ISDA's.
 *
 * Dates no more than a year apart use 366 when a February 29th falls between them (or the end is
 * one) and 365 otherwise; dates further apart use the average length of every calendar year they
 * touch. ISDA's split into leap and non-leap portions agrees on whole years but not across them —
 * `YEARFRAC(2020-06-01, 2023-06-01, 1)` is 1095 / 365.25 in Excel. Leap years are the real
 * calendar's: Excel takes 1900 as 365 days here.
 */
export function yearLength(start: number, end: number, basis: number, date1904?: boolean): number {
  if (basis !== 1) {
    return basis === 3 ? 365 : 360;
  }
  const startSerial = Math.floor(start);
  const endSerial = Math.floor(end);
  const s = calendarParts(startSerial, date1904);
  const e = calendarParts(endSerial, date1904);
  const isLeap = isGregorianLeapYear;
  const withinAYear =
    s.year === e.year ||
    (s.year + 1 === e.year && (s.month > e.month || (s.month === e.month && s.day >= e.day)));
  if (!withinAYear) {
    let days = 0;
    for (let year = s.year; year <= e.year; year++) {
      days += isLeap(year) ? 366 : 365;
    }
    return days / (e.year - s.year + 1);
  }
  if (s.year === e.year) {
    return isLeap(s.year) ? 366 : 365;
  }
  const leapDayBetween =
    (isLeap(s.year) && startSerial < partsToSerial({ year: s.year, month: 3, day: 1 }, date1904)) ||
    (isLeap(e.year) && endSerial >= partsToSerial({ year: e.year, month: 3, day: 1 }, date1904)) ||
    (e.month === 2 && e.day === 29);
  return leapDayBetween ? 366 : 365;
}

/** `YEARFRAC(start, end, basis)` for `start <= end` and a basis from 0 to 4. */
export function yearFraction(
  start: number,
  end: number,
  basis: number,
  date1904?: boolean
): number {
  return dayCount(start, end, basis, date1904) / yearLength(start, end, basis, date1904);
}
