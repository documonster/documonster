/**
 * The forms a built-in number format takes in the reader's locale.
 *
 * Some built-in ids have no format code: the file names the id, and each Excel shows it in its own locale's form
 * (`kept-num-fmt.ts`). A read keeps the id and writes it back, so this is for display only — a cell's display text,
 * a column fitted to it, an export to PDF or Word — and the caller says whose display it is: `"zh-CN"` shows id 7 as
 * `¥3.50`, `"en-US"` as `$3.50`. A locale changes nothing else: a format with a code shows as its code says.
 *
 * Where each form comes from:
 *
 * - **The currencies 5–8 and 63–66, the accounting formats 41–44, and 23–26** are what Excel for Mac (16.115)
 *   wrote back for each id under each locale with its own interface language, saving a workbook of every one. The
 *   standard defines no form for them.
 * - **The East Asian dates and times 27–36 and 50–58, and the Thai formats 59–81**, are ECMA-376's forms for the
 *   locales it gives them (§18.8.30), which LibreOffice also shows — except where Excel and LibreOffice both show
 *   something else, and then LibreOffice's form ({@link LocaleForms.corrections}). Where only Excel for Mac departs
 *   the departure is its own and is not followed: it shows `[$-404]e` as the Western year where the standard and
 *   LibreOffice give the Republic of China's (112 for 2023), writes the Thai day and month codes in the wrong order,
 *   and shows `上午/下午` as `AMtw/PMtw`. In a locale with no form of its own, these ids are shown as the built-in
 *   they translate, as Excel shows them there (an East Asian date as its short date, 14).
 *
 * Measured with 45000.75 and -1234.5 in a cell of each id, for the nine locales below: every text then matches Excel
 * for Mac or LibreOffice, except a date shown as 14 or 22 — which Excel shows in the system's short date, and this
 * library in the code's own (`mm-dd-yy`) — and an accounting format's `*` fill, which only a renderer that knows the
 * cell's width expands.
 */
import { defaultNumFormats } from "@excel/xlsx/defaultnumformats";

/** The locales a display can be asked for. */
export const DISPLAY_LOCALES = [
  "en-US",
  "en-GB",
  "en-AU",
  "zh-CN",
  "zh-TW",
  "zh-HK",
  "ja-JP",
  "ko-KR",
  "th-TH"
] as const;

/**
 * The locale a cell is displayed in: the form a built-in number format that has no code of its own takes — a
 * currency's symbol, an East Asian date, Thai digits. See {@link DISPLAY_LOCALES} for the ones there are.
 */
export type ExcelDisplayLocale = (typeof DISPLAY_LOCALES)[number];

/** The locale a display is in when none is named: the one every built-in code here is already in. */
export const DEFAULT_DISPLAY_LOCALE: ExcelDisplayLocale = "en-US";

/** How a locale writes its currencies: Western, `"$"#,##0_);\("$"#,##0\)`, or Japanese, `"¥"#,##0;"¥"\-#,##0`. */
type CurrencyStyle = "western" | "japanese";

interface LocaleForms {
  /** The currency symbol of 5–8 and 41–44. */
  readonly symbol: string;
  readonly currencyStyle: CurrencyStyle;
  /** The symbol of the reserved 23–26, which an East Asian Excel shows as US dollar currencies; General when absent. */
  readonly dollar?: string;
  /** ECMA-376's column for 27–36 and 50–81 (`defaultnumformats.ts`), where the locale has one. */
  readonly standard?: "zh-cn" | "zh-tw" | "ja-jp" | "ko-kr" | "th-th";
  /**
   * The forms Excel and LibreOffice both show other than the standard's, as LibreOffice writes them: no space after
   * `上午/下午`, Hong Kong's day before its month, Korean dates of 34 with slashes.
   */
  readonly corrections?: Readonly<Record<number, string>>;
}

/** The same form for each of `ids`. */
function each(ids: readonly number[], code: string): Record<number, string> {
  return Object.fromEntries(ids.map(id => [id, code]));
}

const LOCALE_FORMS: Readonly<Record<ExcelDisplayLocale, LocaleForms>> = {
  "en-US": { symbol: "$", currencyStyle: "western" },
  "en-GB": { symbol: "£", currencyStyle: "western" },
  "en-AU": { symbol: "$", currencyStyle: "western" },
  "zh-CN": {
    symbol: "¥",
    currencyStyle: "western",
    dollar: "\\$",
    standard: "zh-cn",
    corrections: {
      ...each([34, 55], '上午/下午h"时"mm"分"'),
      ...each([35, 56], '上午/下午h"时"mm"分"ss"秒"')
    }
  },
  "zh-TW": {
    symbol: "$",
    currencyStyle: "western",
    dollar: '"US$"',
    standard: "zh-tw",
    corrections: {
      ...each([34, 52, 55], '上午/下午hh"時"mm"分"'),
      ...each([35, 53, 56], '上午/下午hh"時"mm"分"ss"秒"')
    }
  },
  // Hong Kong takes Taiwan's forms — Excel and LibreOffice both give it the `[$-404]` dates — in its own order.
  "zh-HK": {
    symbol: "HK$",
    currencyStyle: "western",
    dollar: '"US$"',
    standard: "zh-tw",
    corrections: {
      ...each([27, 36, 50, 57], "[$-404]d/m/e"),
      ...each([28, 29, 51, 54, 58], '[$-404]d"日"m"月"e"年"'),
      31: 'd"日"m"月"yyyy"年"',
      ...each([34, 52, 55], '上午/下午h"時"mm"分"'),
      ...each([35, 53, 56], '上午/下午h"時"mm"分"ss"秒"')
    }
  },
  "ja-JP": { symbol: "¥", currencyStyle: "japanese", dollar: "\\$", standard: "ja-jp" },
  "ko-KR": {
    symbol: "₩",
    currencyStyle: "western",
    dollar: "\\$",
    standard: "ko-kr",
    corrections: each([34, 52, 55], "yyyy/mm/dd")
  },
  "th-TH": { symbol: "฿", currencyStyle: "western", standard: "th-th" }
};

/** A currency format, 5–8 or 63–66 by its position from 5: whole or two places, plain or red when negative. */
function currencyCode(
  symbol: string,
  style: CurrencyStyle,
  offset: number,
  digits: "" | "t"
): string {
  const number = offset >= 2 ? "#,##0.00" : "#,##0";
  const red = offset % 2 === 1 ? "[Red]" : "";
  return style === "japanese"
    ? `${symbol}${number};${red}${symbol}\\-${number}`
    : `${digits}${symbol}${number}_);${red}\\(${digits}${symbol}${number}\\)`;
}

/** An accounting format, 41–44 by its position from 41: with no symbol or with one, whole or two places. */
function accountingCode(symbol: string, style: CurrencyStyle, offset: number): string {
  const number = offset >= 2 ? "#,##0.00" : "#,##0";
  const dash = offset >= 2 ? '"-"??' : '"-"';
  const shown = offset % 2 === 1 ? `"${symbol}"` : "";
  return style === "japanese"
    ? `_ ${shown}* ${number}_ ;_ ${shown}* \\-${number}_ ;_ ${shown}* ${dash}_ ;_ @_ `
    : `_(${shown}* ${number}_);_(${shown}* \\(${number}\\);_(${shown}* ${dash}_);_(@_)`;
}

/**
 * The built-in id each East Asian and Thai id is shown as in a locale with no form of its own: what Excel in an
 * English locale shows and writes it back as, measured as the currencies were, and what LibreOffice shows too. The
 * East Asian dates are its short date (14) and the times its `h:mm:ss` (21). The Thai ids are not each the Latin
 * format it translates: from 72 on, both shift by one against ECMA-376's Thai forms, so 75 (`ช:นน`, a time) shows as
 * `mmm-yy` and 81 (`d/m/bb`, a date) as `mm:ss.0`. That is followed, because it is what a reader of the file sees.
 */
const SHOWN_AS_BUILT_IN = new Map<number, number>([
  ...[27, 28, 29, 30, 31, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58].map(id => [id, 14] as const),
  ...[32, 33, 34, 35].map(id => [id, 21] as const),
  [59, 1],
  [60, 2],
  [61, 3],
  [62, 4],
  [67, 9],
  [68, 10],
  [69, 12],
  [70, 13],
  [71, 14],
  [72, 14],
  [73, 15],
  [74, 16],
  [75, 17],
  [76, 20],
  [77, 21],
  [78, 22],
  [79, 45],
  [80, 46],
  [81, 47]
]);

/**
 * The code a cell whose format is the kept built-in `id` is shown with in `locale`, or `undefined` for an id no
 * Excel defines a format for (82–163).
 */
export function shownCodeOfKeptId(
  id: number,
  locale: ExcelDisplayLocale = DEFAULT_DISPLAY_LOCALE
): string | undefined {
  const forms = LOCALE_FORMS[locale];
  const symbol = `"${forms.symbol}"`;
  if (id >= 5 && id <= 8) {
    return currencyCode(symbol, forms.currencyStyle, id - 5, "");
  }
  if (id >= 63 && id <= 66) {
    // Thai Excel shows these in Thai digits (`t`); every other locale as 5–8.
    return currencyCode(
      symbol,
      forms.currencyStyle,
      id - 63,
      forms.standard === "th-th" ? "t" : ""
    );
  }
  if (id >= 41 && id <= 44) {
    return accountingCode(forms.symbol, forms.currencyStyle, id - 41);
  }
  if (id >= 23 && id <= 26) {
    return forms.dollar === undefined
      ? "General"
      : currencyCode(forms.dollar, "western", id - 23, "");
  }
  const standard =
    forms.corrections?.[id] ??
    (forms.standard === undefined ? undefined : defaultNumFormats[id]?.[forms.standard]);
  if (standard !== undefined) {
    return standard;
  }
  const builtIn = SHOWN_AS_BUILT_IN.get(id);
  return builtIn === undefined ? undefined : defaultNumFormats[builtIn]?.f;
}
