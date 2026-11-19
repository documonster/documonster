/**
 * How a format takes what it leaves out from the format it inherits — the workbook's default cell format, or a cell's
 * named style — and how a read says that a format holds a facet plainly. The XLSX and XLSB readers and writers share
 * these rules rather than each holding a copy, so the two containers cannot come to disagree about what a format is:
 * the copies already had once, and an XLSB date written under a default date-time format lost its time.
 *
 * Each container keeps its own shape — a style object on the XLSX side, a cell-format table on the XLSB side — and
 * hands these functions the parts they decide on.
 */
import type { Borders, Fill, Font, Style } from "@excel/types";
import { keepNumFmtId, keptNumFmtId, showsDate } from "@excel/utils/kept-num-fmt";
import { numFmtCode } from "@excel/utils/num-fmt-code";

/** The facets a format takes one by one from the format it inherits, where it does not hold them itself. */
export const INHERITED_FACETS = ["font", "fill", "border", "alignment", "protection"] as const;

/** A number format as both containers hold it: a code, or a built-in id kept for want of one (`kept-num-fmt.ts`). */
export interface NumberFormatRef {
  readonly code?: string | undefined;
  readonly keptId?: number | undefined;
}

/**
 * The number format a format that holds none takes from the one it inherits, or `undefined` when it takes none. A
 * date takes it only where it shows a date — a date under a default `0.00` would open as `45306.00` — and is then left
 * to its container's own date format.
 */
export function inheritedNumberFormat(
  inherited: NumberFormatRef | undefined,
  isDate: boolean
): NumberFormatRef | undefined {
  if (inherited === undefined || (inherited.code === undefined && inherited.keptId === undefined)) {
    return undefined;
  }
  if (isDate && !showsDate(inherited.code, inherited.keptId)) {
    return undefined;
  }
  return inherited.code !== undefined ? { code: inherited.code } : { keptId: inherited.keptId };
}

/** A style's number format as a {@link NumberFormatRef}. */
export function numberFormatOfStyle(style: Partial<Style> | undefined): NumberFormatRef {
  const numFmt = style?.numFmt;
  return numFmt ? { code: numFmtCode(numFmt) } : { keptId: keptNumFmtId(style) };
}

/**
 * `model` with what it leaves out taken from `inherited`: each facet it does not hold, its number format by
 * {@link inheritedNumberFormat}, and its named style. `model` is not changed.
 */
export function withInheritedFacets(
  model: Partial<Style>,
  inherited: Partial<Style>,
  isDate: boolean
): Partial<Style> {
  const merged: Partial<Style> = { ...model };
  const own = numberFormatOfStyle(model);
  if (own.code === undefined && own.keptId === undefined) {
    const taken = inheritedNumberFormat(numberFormatOfStyle(inherited), isDate);
    if (taken?.code !== undefined) {
      merged.numFmt = taken.code;
    } else if (taken?.keptId !== undefined) {
      keepNumFmtId(merged, taken.keptId);
    }
  }
  for (const facet of INHERITED_FACETS) {
    if (merged[facet] === undefined && inherited[facet] !== undefined) {
      (merged as Record<string, unknown>)[facet] = inherited[facet];
    }
  }
  if (!merged.styleName && inherited.styleName) {
    merged.styleName = inherited.styleName;
  }
  return merged;
}

/** What a container reads as a facet's plain value: its first font, its first fill, its first border. */
export interface PlainFacets {
  readonly font: Partial<Font>;
  readonly fill: Fill;
  readonly border: Partial<Borders>;
}

/**
 * Where the default cell format sets a facet that `model` — a format read from the file — leaves plain, say so: General,
 * the plain font, fill and border, no alignment, the default protection, the Normal style. A style that holds
 * nothing is the default format — it is what a cell with no style is written as — so a plain facet left out in a
 * workbook whose default is not plain would read, and be written, as the default's.
 *
 * `ownKeptId` is the format's own kept number-format id, for a reader that keeps it beside `model` rather than on it.
 */
export function statePlainFacets(
  model: Partial<Style>,
  defaults: Partial<Style> | undefined,
  plain: PlainFacets,
  ownKeptId?: number
): void {
  if (defaults === undefined) {
    return;
  }
  const own = numberFormatOfStyle(model);
  const inherited = numberFormatOfStyle(defaults);
  if (
    (inherited.code !== undefined || inherited.keptId !== undefined) &&
    own.code === undefined &&
    own.keptId === undefined &&
    ownKeptId === undefined
  ) {
    model.numFmt = "General";
  }
  if (defaults.font && !model.font) {
    model.font = plain.font;
  }
  if (defaults.fill && !model.fill) {
    model.fill = plain.fill;
  }
  if (defaults.border && !model.border) {
    model.border = plain.border;
  }
  if (defaults.alignment && !model.alignment) {
    model.alignment = {};
  }
  if (defaults.protection && !model.protection) {
    model.protection = { locked: true, hidden: false };
  }
  if (defaults.styleName && !model.styleName) {
    model.styleName = "Normal";
  }
}
