/**
 * Page-number fields (`PAGE`, `NUMPAGES`, `SECTIONPAGES`) and the number
 * formats they and list markers are drawn in — shared by the layout (headers
 * and footers) and the field engine, so both show the same text.
 */

import type { DocxDocument, NumberFormat, PageNumberFormat, SectionProperties } from "@word/types";

/** A parsed page-number field instruction. */
export interface PageField {
  readonly name: "PAGE" | "NUMPAGES" | "SECTIONPAGES";
  /** The field's own `\\* <format>` switch, if any. */
  readonly format: NumberFormat | undefined;
}

/**
 * Page-number field name and number format, or `undefined` when the
 * instruction is not one this layout computes.
 *
 * Accepts the switches Word writes on these fields: `\\* MERGEFORMAT` and
 * `\\* CHARFORMAT` (presentation only) and the numeric format switches. Any
 * other switch (`\\# picture`, `\\@`) means the field cannot be computed
 * faithfully here, so its cached result is shown instead.
 */
export function parsePageField(instruction: string): PageField | undefined {
  const tokens = instruction.trim().split(/\s+/);
  const name = tokens[0]?.toUpperCase();
  if (name !== "PAGE" && name !== "NUMPAGES" && name !== "SECTIONPAGES") {
    return undefined;
  }
  let format: NumberFormat | undefined;
  for (let i = 1; i < tokens.length; i++) {
    if (tokens[i] !== "\\*" || i + 1 >= tokens.length) {
      return undefined;
    }
    const arg = tokens[++i];
    const mapped = FIELD_NUMBER_FORMATS[arg];
    if (mapped) {
      format = mapped;
    } else if (!/^(MERGEFORMAT|CHARFORMAT)$/i.test(arg)) {
      return undefined;
    }
  }
  return { name, format };
}

/** `\\* <format>` switch → list number format (case matters: `roman` vs `ROMAN`). */
const FIELD_NUMBER_FORMATS: Readonly<Record<string, NumberFormat>> = {
  Arabic: "decimal",
  arabic: "decimal",
  roman: "lowerRoman",
  ROMAN: "upperRoman",
  Roman: "upperRoman",
  alphabetic: "lowerLetter",
  ALPHABETIC: "upperLetter"
};

/** A section page-number format that can be drawn as a number, or `undefined`. */
export function pageFormatAsNumberFormat(
  format: PageNumberFormat | undefined
): NumberFormat | undefined {
  switch (format) {
    case undefined:
    case "decimal":
      return "decimal";
    case "upperRoman":
    case "lowerRoman":
    case "upperLetter":
    case "lowerLetter":
      return format;
    default:
      return undefined;
  }
}

/**
 * The text a page-number field shows for value `n`, or `undefined` when its
 * format cannot be drawn (the caller then keeps the cached result). The field's
 * own switch wins; `PAGE` otherwise follows the section's `w:pgNumType w:fmt`.
 */
export function formatPageField(
  field: PageField,
  n: number,
  sectionFormat: PageNumberFormat | undefined
): string | undefined {
  const format =
    field.format ?? (field.name === "PAGE" ? pageFormatAsNumberFormat(sectionFormat) : "decimal");
  return format === undefined ? undefined : formatListCounter(n, format);
}

/** Each section's properties in body order; the last is `doc.sectionProperties`. */
export function collectSectionProperties(doc: DocxDocument): (SectionProperties | undefined)[] {
  const sections: (SectionProperties | undefined)[] = [];
  for (const item of doc.body) {
    if (item.type === "paragraph" && item.properties?.sectionProperties) {
      sections.push(item.properties.sectionProperties);
    }
  }
  sections.push(doc.sectionProperties);
  return sections;
}

/** Where a physical page sits among the document's sections. */
export interface PageSectionSlot {
  /** Index into `collectSectionProperties(doc)`. */
  readonly sectionIndex: number;
  /**
   * 1-based position within its section. `0` marks the blank page Word inserts
   * before an odd/even-page section break: it precedes the new section's first
   * page, shows no header or footer and counts towards no section.
   */
  readonly pageInSection: number;
}

/**
 * How a section that starts a new page opens, by Word's rules (each checked
 * against Word's own PDF output — see `page-number-sections.test.ts`):
 *
 * - the section's own `w:type` decides the break — the type on a section's
 *   `sectPr` describes how *that* section starts, not how it ends;
 * - its first page shows `w:pgNumType w:start` when set, otherwise the number
 *   after `previousShown`;
 * - an odd/even-page break judges parity on that *displayed* number. When it
 *   is wrong and the section restarts numbering, Word skips a number instead
 *   of a page; when it continues numbering, Word inserts a blank page, which
 *   takes the skipped number.
 */
export function openSection(
  previousShown: number,
  props: SectionProperties | undefined
): { readonly blankPage: boolean; readonly firstShown: number } {
  const start = props?.pageNumbering?.start;
  const candidate = start ?? previousShown + 1;
  const parity =
    props?.breakType === "oddPage" ? 1 : props?.breakType === "evenPage" ? 0 : undefined;
  if (parity === undefined || candidate % 2 === parity) {
    return { blankPage: false, firstShown: candidate };
  }
  return { blankPage: start === undefined, firstShown: candidate + 1 };
}

/**
 * The number `PAGE` shows on each page, in page order (see {@link openSection}).
 * Pages after a section's first continue from the previous page's *displayed*
 * number, not from the physical page number.
 */
export function displayedPageNumbers(
  pages: readonly PageSectionSlot[],
  sections: readonly (SectionProperties | undefined)[]
): number[] {
  const out: number[] = [];
  for (let i = 0; i < pages.length; i++) {
    out.push(
      nextDisplayedPage(i > 0 ? pages[i - 1] : undefined, out[i - 1] ?? 0, pages[i], sections)
    );
  }
  return out;
}

/**
 * One step of `displayedPageNumbers`: the number `page` shows given the page
 * before it (`undefined` for the first page) and the number that one showed.
 * For callers that create pages one at a time.
 */
export function nextDisplayedPage(
  previous: PageSectionSlot | undefined,
  previousShown: number,
  page: PageSectionSlot,
  sections: readonly (SectionProperties | undefined)[]
): number {
  // A parity blank page and every page after a section's first continue the count.
  if (page.pageInSection !== 1) {
    return previousShown + 1;
  }
  const props = sections[page.sectionIndex];
  if (previous === undefined) {
    // The document's first page: no break precedes it, so no parity applies.
    return props?.pageNumbering?.start ?? 1;
  }
  return openSection(previousShown, props).firstShown;
}

/**
 * Number of physical pages per section index, for `SECTIONPAGES`. A parity
 * blank page belongs to no section, as Word counts it.
 */
export function sectionPageCounts(pages: readonly PageSectionSlot[]): Map<number, number> {
  const counts = new Map<number, number>();
  for (const { sectionIndex, pageInSection } of pages) {
    if (pageInSection !== 0) {
      counts.set(sectionIndex, (counts.get(sectionIndex) ?? 0) + 1);
    }
  }
  return counts;
}

/** Format an ordered-list counter per its OOXML number format. */
export function formatListCounter(n: number, format: NumberFormat): string {
  switch (format) {
    case "lowerLetter":
      return toAlpha(n).toLowerCase();
    case "upperLetter":
      return toAlpha(n).toUpperCase();
    case "lowerRoman":
      return toRoman(n).toLowerCase();
    case "upperRoman":
      return toRoman(n).toUpperCase();
    case "decimalZero":
      return n < 10 ? `0${n}` : String(n);
    default:
      // decimal and any non-numeric/locale formats we don't render.
      return String(n);
  }
}

/**
 * Word's alphabetic numbering: 1 → "A", 26 → "Z", 27 → "AA", 28 → "BB" — the
 * letter repeats once per pass through the alphabet (not spreadsheet columns,
 * where 28 would be "AB"). The text grows linearly with `n`, so a value from
 * a hostile `w:start` beyond {@link MAX_ALPHA_REPEAT} passes falls back to
 * digits rather than allocating an unbounded string.
 */
function toAlpha(n: number): string {
  if (n <= 0) {
    return "A";
  }
  const repeat = Math.ceil(n / 26);
  if (repeat > MAX_ALPHA_REPEAT) {
    return String(n);
  }
  return String.fromCharCode(65 + ((n - 1) % 26)).repeat(repeat);
}

/** Longest run of one letter `toAlpha` produces. */
const MAX_ALPHA_REPEAT = 64;

/** Convert a positive integer to a Roman numeral (uppercase). */
function toRoman(n: number): string {
  if (n <= 0) {
    return String(n);
  }
  const table: [number, string][] = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"]
  ];
  let v = n;
  let s = "";
  for (const [val, sym] of table) {
    while (v >= val) {
      s += sym;
      v -= val;
    }
  }
  return s;
}
