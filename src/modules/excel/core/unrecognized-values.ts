/**
 * The values a read kept because they are outside their field's vocabulary, listed — what
 * `Workbook.readWithDiagnostics` reports as `unrecognizedValues` and `Workbook.listUnrecognizedValues` returns.
 *
 * Two kinds: a sheet setting (`sheet-settings-checks`), and a number format kept by its built-in id on a cell, row,
 * column or styled blank range (`kept-num-fmt`), whose path names it as `Cell`, `Row` and `Column` address it —
 * `A4.style.numFmt`, `4:4.style.numFmt`, `B:B.style.numFmt`, and `A5:C9.style.numFmt` for a styled blank range — or
 * on a workbook-level style, which belongs to no sheet — the default cell format (`defaultStyle.numFmt`), the Normal
 * style (`cellStyles["Normal"].numFmt`) or a named one (`cellStyles["Currency"].numFmt`).
 *
 * A module of its own, with nothing but the workbook core beneath it, because `Workbook.listUnrecognizedValues` is part
 * of the workbook core that every bundle reaching Excel imports — the Word and PDF bridges included. It used to live in
 * `workbook-io-types.ts`, beside `readWorkbookStream` and its `import()` of the XLSX reader: importing it from the core
 * put that `import()` into every such bundle, where the reader's chunk outlived the tree-shaken caller and made a
 * single-file (IIFE) bundle impossible.
 */
import type { NamedStyleEntry, WorkbookData } from "@excel/core/workbook-core";
import { getWorksheets } from "@excel/core/workbook-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import {
  getSheetName,
  getStyledBlankRanges,
  setStyledBlankRanges
} from "@excel/core/worksheet-core";
import type { Style } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { keptNumFmtId } from "@excel/utils/kept-num-fmt";
import { findKeptSheetSettings } from "@excel/utils/sheet-settings-checks";
import { excelNameKey } from "@utils/name-key";
import {
  discardUnrecognizedValues as discardKept,
  type KeptValue
} from "@utils/unrecognized-values";

/**
 * A value a read kept because it is outside its setting's vocabulary: the sheet it is on (absent for a workbook-level
 * setting), the setting's path (`pageSetup.orientation`, `views[0].zoomScale`, `A4.style.numFmt`) and the text the
 * file gave.
 */
export interface UnrecognizedValue {
  readonly sheet?: string;
  readonly path: string;
  readonly value: string;
}

/** The path of a style's `numFmt`, for the style of what `at` names. */
function numFmtPath(at: string): string {
  return `${at}.style.numFmt`;
}

/**
 * Every style on `ws` that keeps a number format by its id, with what it belongs to as `Cell`, `Row` and `Column`
 * address it. `replace` is how a style is changed: a cell's style may be shared, a styled blank range's is the read's
 * own, so neither is edited in place.
 */
function eachKeptNumFmt(
  ws: WorksheetData,
  visit: (at: string, style: Partial<Style>, replace: (style: Partial<Style>) => void) => void
): void {
  for (const column of ws._columns) {
    if (column && keptNumFmtId(column.style) !== undefined) {
      const letter = colCache.n2l(column.number);
      visit(`${letter}:${letter}`, column.style, style => (column.style = style));
    }
  }
  for (const row of ws._rows) {
    if (!row) {
      continue;
    }
    if (keptNumFmtId(row.style) !== undefined) {
      visit(`${row.number}:${row.number}`, row.style, style => (row.style = style));
    }
    for (const cell of row.cells) {
      if (cell && keptNumFmtId(cell.style) !== undefined) {
        visit(cell.address, cell.style, style => (cell.style = style));
      }
    }
  }
  const ranges = getStyledBlankRanges(ws);
  if (ranges?.some(range => keptNumFmtId(range.style) !== undefined)) {
    const replaced = ranges.map(range => {
      let style = range.style;
      if (keptNumFmtId(style) !== undefined) {
        const at = `${colCache.encodeAddress(range.firstRow + 1, range.firstColumn + 1)}:${colCache.encodeAddress(range.lastRow + 1, range.lastColumn + 1)}`;
        visit(at, style, next => (style = next));
      }
      return style === range.style ? range : { ...range, style };
    });
    if (replaced.some((range, i) => range !== ranges[i])) {
      setStyledBlankRanges(ws, replaced);
    }
  }
}

/**
 * The values a read kept on `wb`'s sheets, as they are kept now — the `unrecognizedValues` of a report, and
 * `Workbook.listUnrecognizedValues`. Fresh objects.
 */
export function keptSettingsReport(wb: WorkbookData): UnrecognizedValue[] {
  const entries: UnrecognizedValue[] = [];
  for (const { path, style } of keptWorkbookFormats(wb)) {
    entries.push({ path, value: String(keptNumFmtId(style)) });
  }
  for (const ws of getWorksheets(wb)) {
    const sheet = getSheetName(ws);
    for (const { path, value } of findKeptSheetSettings(ws)) {
      entries.push({ sheet, path, value });
    }
    eachKeptNumFmt(ws, (at, style) => {
      entries.push({ sheet, path: numFmtPath(at), value: String(keptNumFmtId(style)) });
    });
  }
  return entries;
}

/**
 * The workbook-level styles that keep a number format by its id, with the path of that format and how to replace the
 * style: the default cell format (`defaultStyle.numFmt`, what `Workbook.getDefaultStyle` returns), the Normal style
 * (`cellStyles["Normal"].numFmt`) and the named cell styles (`cellStyles["Currency"].numFmt`, by the name
 * `Workbook.getCellStyle` takes). Reported without a sheet.
 */
function keptWorkbookFormats(wb: WorkbookData): {
  readonly path: string;
  readonly style: Partial<Style>;
  readonly replace: (style: Partial<Style>) => void;
}[] {
  const found: {
    path: string;
    style: Partial<Style>;
    replace: (style: Partial<Style>) => void;
  }[] = [];
  if (keptNumFmtId(wb._defaultStyle) !== undefined) {
    found.push({
      path: "defaultStyle.numFmt",
      style: wb._defaultStyle!,
      replace: style => (wb._defaultStyle = style)
    });
  }
  if (keptNumFmtId(wb._normalStyle) !== undefined) {
    found.push({
      path: 'cellStyles["Normal"].numFmt',
      style: wb._normalStyle!,
      replace: style => (wb._normalStyle = style)
    });
  }
  for (const style of wb._cellStyles?.values() ?? []) {
    if (keptNumFmtId(style) !== undefined) {
      found.push({
        path: `cellStyles[${JSON.stringify(style.name)}].numFmt`,
        style,
        replace: copy => wb._cellStyles!.set(excelNameKey(style.name), copy as NamedStyleEntry)
      });
    }
  }
  return found;
}

/**
 * Drop the number formats the workbook-level styles of `wb` keep by their id — all, or those `filter` returns `true`
 * for — and return how many there were.
 */
export function discardKeptWorkbookFormats(
  wb: WorkbookData,
  filter?: (value: KeptValue) => boolean
): number {
  let discarded = 0;
  for (const { path, style, replace } of keptWorkbookFormats(wb)) {
    if (filter !== undefined && !filter({ path, value: String(keptNumFmtId(style)) })) {
      continue;
    }
    const copy = { ...style };
    discarded += discardKept(copy, kept => kept.path === "numFmt");
    replace(copy);
  }
  return discarded;
}

/**
 * Drop the number formats a read kept by their id on `ws` — all, or those `filter` returns `true` for — and return
 * how many there were.
 */
export function discardKeptNumFmts(
  ws: WorksheetData,
  filter: ((value: KeptValue) => boolean) | undefined,
  defaults: Partial<Style> | undefined
): number {
  let discarded = 0;
  // A sheet's format with no number format has the default's (`Workbook.getDefaultStyle`). Where the default sets one,
  // dropping a kept id would hand the format the default's — the kept id itself, often — so it is stated as General:
  // that is what dropping it means, and what it is written as.
  const inherits =
    defaults !== undefined &&
    (defaults.numFmt !== undefined || keptNumFmtId(defaults) !== undefined);
  eachKeptNumFmt(ws, (at, style, replace) => {
    const path = numFmtPath(at);
    if (filter !== undefined && !filter({ path, value: String(keptNumFmtId(style)) })) {
      return;
    }
    const copy = { ...style };
    discarded += discardKept(copy, kept => kept.path === "numFmt");
    if (inherits) {
      copy.numFmt = "General";
    }
    replace(copy);
  });
  return discarded;
}
