/**
 * Rewriting references through a structural edit — the one place that knows what Excel does to a reference when rows
 * or columns are inserted or deleted, a sheet is renamed, or a sheet is removed.
 *
 * Everything that names cells by position goes through here: cell formulas on every sheet, defined names, tables,
 * conditional formats, data validations, chart series, sparklines, pivot sources. Each of them used to be left as it
 * was, so after `Row.insert` a formula summed the wrong cells and a chart plotted the wrong range without anything
 * reporting it.
 *
 * The rules are Excel's, and they apply to absolute and relative references alike — `$` decides what a copy does, not
 * what an insertion does:
 *
 * - **Insert** `n` lines at `at`: a line at or after `at` moves by `n`. Applied to each end of a range on its own, so an
 *   insertion at a range's first line moves the range, one strictly inside it grows it, and one just past its last
 *   line leaves it alone.
 * - **Delete** lines `at … at+n-1`: a reference wholly inside the band becomes `#REF!`; a range that overlaps it
 *   shrinks by the overlap; anything after it moves back by `n`.
 * - A splice that deletes and inserts is the deletion followed by the insertion, at the same place.
 * - A reference spanning the whole axis being edited (`A:A`, or `A1:A1048576`, for a row edit) is left alone — it
 *   still means every row. A range end pushed past the sheet's last line stops there; a reference that would start
 *   past it is `#REF!`.
 * - A 3-D reference (`Jan:Mar!A1`) is not changed by a row or column edit on one of its sheets, as in Excel: it names
 *   the same cells on every sheet of the span, and only one of them moved.
 * - **Rename**: every sheet prefix naming the sheet is rewritten, and quoted as the new name requires for the new name.
 * - **Remove**: a reference qualified with the removed sheet becomes `#REF!`. A 3-D span that ends on it shrinks to
 *   the next sheet inward, and one made of that sheet alone becomes `#REF!`.
 *
 * Where Excel writes a qualified `Sheet1!#REF!`, this writes `#REF!` alone: it evaluates to the same error and it is
 * a formula every reader of this library parses.
 *
 * References are found with the formula tokenizer, never a regular expression, so a string, a structured reference
 * (`Table1[Col]`), a defined name, an external reference (`[1]Sheet1!A1`) or a function named like a cell (`LOG10(`)
 * is never mistaken for one. Text the edit does not touch is kept byte for byte.
 */
import { quoteSheetName } from "@excel/utils/address";
import { colCache } from "@excel/utils/col-cache";
import type { Token } from "@formula/syntax/token-types";
import { isSpillOperator, TokenType } from "@formula/syntax/token-types";
import { tokenizeWithSpans } from "@formula/syntax/tokenizer";
import { excelNameKey, sheetNameKey } from "@utils/name-key";

const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

/** The error a reference to cells that no longer exist becomes. */
export const REF_ERROR = "#REF!";

/** Rows or columns deleted and inserted at one place on one sheet. */
export interface LineEdit {
  readonly kind: "rows" | "columns";
  readonly sheet: string;
  /** First line (1-based) the edit deletes, or before which it inserts. */
  readonly at: number;
  readonly deleted: number;
  readonly inserted: number;
  /** Tables the edit deletes, by `excelNameKey` — a structured reference to one becomes `#REF!`. */
  readonly removedTables?: ReadonlySet<string>;
}

export interface SheetRenameEdit {
  readonly kind: "rename";
  readonly from: string;
  readonly to: string;
}

export interface SheetRemoveEdit {
  readonly kind: "removeSheet";
  readonly sheet: string;
  /** Every sheet's name in tab order, the removed one included — a 3-D span is resolved against it. */
  readonly sheets: readonly string[];
  /** The removed sheet's tables, by `excelNameKey` — see {@link LineEdit.removedTables}. */
  readonly removedTables?: ReadonlySet<string>;
}

export type ReferenceEdit = LineEdit | SheetRenameEdit | SheetRemoveEdit;

// ============================================================================
// Lines and spans
// ============================================================================

function maxOf(edit: LineEdit): number {
  return edit.kind === "rows" ? MAX_ROW : MAX_COL;
}

/** Where line `n` goes, or `undefined` when the edit deletes it or pushes it off the sheet. */
export function editLine(n: number, edit: LineEdit): number | undefined {
  return editSpan(n, n, edit)?.[0];
}

/**
 * Where the lines `lo … hi` go, or `undefined` when the edit deletes all of them — see the module note for the rules.
 * A span covering the whole axis is returned unchanged.
 */
export function editSpan(lo: number, hi: number, edit: LineEdit): [number, number] | undefined {
  const max = maxOf(edit);
  if (lo === 1 && hi === max) {
    return [lo, hi];
  }
  const { at, deleted, inserted } = edit;
  let a = lo;
  let b = hi;
  if (deleted > 0) {
    const end = at + deleted - 1;
    if (a >= at && b <= end) {
      return undefined;
    }
    a = a < at ? a : a <= end ? at : a - deleted;
    b = b < at ? b : b <= end ? at - 1 : b - deleted;
  }
  if (inserted > 0) {
    a = a >= at ? a + inserted : a;
    b = b >= at ? b + inserted : b;
  }
  if (a > max) {
    return undefined;
  }
  return [a, Math.min(b, max)];
}

/** A drawing corner: a 0-based cell index on one axis and an offset into that cell. */
export interface AnchorCorner {
  index: number;
  offset: number;
}

/**
 * A drawing corner through a line edit on its axis.
 *
 * A corner lies on a line — the one its index names — and moves with it; a corner on a deleted line goes to the start
 * of the first line after the band. A bottom-right corner at offset 0 is the far edge of the line before it (that is
 * how `A1:C3` is anchored, with its corner at the top of row 4), so it belongs to that line: inserting at row 4 then
 * leaves it where it was.
 */
export function editAnchorCorner(corner: AnchorCorner, edit: LineEdit, end: boolean): AnchorCorner {
  const edge = end && corner.offset === 0;
  let owner = corner.index + (edge ? 0 : 1);
  if (edit.deleted > 0) {
    const bandEnd = edit.at + edit.deleted - 1;
    if (owner >= edit.at && owner <= bandEnd) {
      return { index: edit.at - 1, offset: 0 };
    }
    if (owner > bandEnd) {
      owner -= edit.deleted;
    }
  }
  if (edit.inserted > 0 && owner >= edit.at) {
    owner += edit.inserted;
  }
  owner = Math.min(owner, maxOf(edit));
  return edge ? { index: owner, offset: 0 } : { index: owner - 1, offset: corner.offset };
}

/** How a drawing follows the cells under it — DrawingML's `editAs`. */
export type AnchorPlacement = "twoCell" | "oneCell" | "absolute";

/**
 * A drawing's corners through a line edit, as Excel moves an object: `twoCell` ("move and size with cells") moves
 * each corner with its own line, so an insertion inside the object makes it taller; `oneCell` ("move but don't
 * size") moves the top-left corner and keeps the size; `absolute` does not move. A drawing with no bottom-right
 * corner (a one-cell anchor, sized by its extent) moves its top-left corner.
 */
export function moveAnchorCorners(
  tl: AnchorCorner,
  br: AnchorCorner | undefined,
  placement: AnchorPlacement,
  edit: LineEdit
): { tl: AnchorCorner; br?: AnchorCorner } {
  if (placement === "absolute") {
    return br === undefined ? { tl } : { tl, br };
  }
  const movedTl = editAnchorCorner(tl, edit, false);
  if (br === undefined) {
    return { tl: movedTl };
  }
  if (placement === "twoCell") {
    return { tl: movedTl, br: editAnchorCorner(br, edit, true) };
  }
  return {
    tl: movedTl,
    br: { index: Math.max(0, br.index + movedTl.index - tl.index), offset: br.offset }
  };
}

// ============================================================================
// A1 references
// ============================================================================

/** One end of a reference: a column and/or a row, each with its `$`. */
interface End {
  col?: number;
  colAbs: boolean;
  row?: number;
  rowAbs: boolean;
}

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
const COL_RE = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_RE = /^(\$?)(\d+)$/;

function parseEnd(text: string): End | undefined {
  let m = CELL_RE.exec(text);
  if (m) {
    const col = colCache.l2n(m[2].toUpperCase());
    const row = Number(m[4]);
    return col <= MAX_COL && row >= 1 && row <= MAX_ROW
      ? { col, colAbs: m[1] === "$", row, rowAbs: m[3] === "$" }
      : undefined;
  }
  m = COL_RE.exec(text);
  if (m) {
    const col = colCache.l2n(m[2].toUpperCase());
    return col <= MAX_COL ? { col, colAbs: m[1] === "$", rowAbs: false } : undefined;
  }
  m = ROW_RE.exec(text);
  if (m) {
    const row = Number(m[2]);
    return row >= 1 && row <= MAX_ROW ? { row, colAbs: false, rowAbs: m[1] === "$" } : undefined;
  }
  return undefined;
}

function formatEnd(end: End): string {
  let out = "";
  if (end.col !== undefined) {
    out += (end.colAbs ? "$" : "") + colCache.n2l(end.col);
  }
  if (end.row !== undefined) {
    out += (end.rowAbs ? "$" : "") + end.row;
  }
  return out;
}

/** An unqualified area: one end, or two of the same shape (`A1:B2`, `A:C`, `1:3`). */
interface Area {
  first: End;
  second?: End;
}

function parseArea(text: string): Area | undefined {
  const colon = text.indexOf(":");
  if (colon < 0) {
    const first = parseEnd(text);
    // A column or a row standing alone is a name, not a reference.
    return first && first.col !== undefined && first.row !== undefined ? { first } : undefined;
  }
  const first = parseEnd(text.slice(0, colon));
  const second = parseEnd(text.slice(colon + 1));
  if (
    !first ||
    !second ||
    (first.col === undefined) !== (second.col === undefined) ||
    (first.row === undefined) !== (second.row === undefined)
  ) {
    return undefined;
  }
  return { first, second };
}

/**
 * The area after a line edit: its text, `REF_ERROR`, or `undefined` when nothing about it changed. The edit is
 * assumed to be on the area's sheet.
 */
function editArea(area: Area, edit: LineEdit): string | undefined {
  const key = edit.kind === "rows" ? "row" : "col";
  const second = area.second ?? area.first;
  const lo = area.first[key];
  const hi = second[key];
  if (lo === undefined || hi === undefined) {
    return undefined; // `A:A` through a row edit, `1:1` through a column edit
  }
  const swapped = lo > hi;
  const span = editSpan(Math.min(lo, hi), Math.max(lo, hi), edit);
  if (span === undefined) {
    return REF_ERROR;
  }
  const [newLo, newHi] = swapped ? [span[1], span[0]] : span;
  if (newLo === lo && newHi === hi) {
    return undefined;
  }
  const first = { ...area.first, [key]: newLo };
  if (area.second === undefined) {
    return formatEnd(first);
  }
  return `${formatEnd(first)}:${formatEnd({ ...area.second, [key]: newHi })}`;
}

/**
 * An unqualified area (`A1`, `$A$1:B2`, `A:C`, `3:5`) on `edit.sheet` after a line edit: the new text, the same text
 * when it is unaffected, or `undefined` when the edit deletes it. Text that is not an area is returned as it is.
 */
export function rewriteArea(text: string, edit: LineEdit): string | undefined {
  const area = parseArea(text.trim());
  if (area === undefined) {
    return text;
  }
  const edited = editArea(area, edit);
  return edited === REF_ERROR ? undefined : (edited ?? text);
}

/**
 * A space-separated list of areas (a conditional format's or a validation's `sqref`) after a line edit on its sheet:
 * the areas that survive, or `undefined` when none does.
 */
export function rewriteSqref(sqref: string, edit: LineEdit): string | undefined {
  const parts = sqref.split(/\s+/).filter(Boolean);
  const kept: string[] = [];
  for (const part of parts) {
    const edited = rewriteArea(part, edit);
    if (edited !== undefined) {
      kept.push(edited);
    }
  }
  return kept.length > 0 ? kept.join(" ") : undefined;
}

// ============================================================================
// Sheet names
// ============================================================================

/** Whether a sheet name has to be quoted in a formula: anything `quoteSheetName` quotes, or a name read as a reference. */
function needsQuotes(name: string): boolean {
  return (
    quoteSheetName(name) !== name ||
    parseArea(name) !== undefined ||
    /^(?:[Rr]\d*[Cc]\d*|[RrCc]\d*)$/.test(name) ||
    /^(?:true|false)$/i.test(name)
  );
}

function escape(name: string): string {
  return name.replace(/'/g, "''");
}

/** `name!` as a formula writes it — quoted, with `'` doubled, when the name requires it. */
export function sheetPrefix(name: string, endName?: string): string {
  if (endName === undefined) {
    return needsQuotes(name) ? `'${escape(name)}'!` : `${name}!`;
  }
  // `C:D!A1` would read as a column range, so a span quotes an end that could be read as a column or row too.
  const quote = (end: string): boolean => needsQuotes(end) || parseEnd(end) !== undefined;
  return quote(name) || quote(endName)
    ? `'${escape(name)}:${escape(endName)}'!`
    : `${name}:${endName}!`;
}

function sameSheet(a: string, b: string): boolean {
  return a === b || sheetNameKey(a) === sheetNameKey(b);
}

/**
 * Whether `text` may name the sheet `edit` concerns. A cheap test that lets a formula on another sheet skip
 * tokenising; it only ever errs towards `true`.
 */
function mayNameSheet(text: string, edit: ReferenceEdit): boolean {
  const sheet = edit.kind === "rename" ? edit.from : edit.sheet;
  // A name outside ASCII may match through Unicode folding that `toLowerCase` does not model.
  if (!/^[\x20-\x7e]*$/.test(sheet)) {
    return true;
  }
  const lower = text.toLowerCase();
  const name = sheet.toLowerCase();
  return lower.includes(name) || (name.includes("'") && lower.includes(escape(name)));
}

/** Whether `text` may name one of `tables`; like {@link mayNameSheet}, it only errs towards `true`. */
function mayNameTable(text: string, tables: ReadonlySet<string> | undefined): boolean {
  if (tables === undefined || tables.size === 0) {
    return false;
  }
  const lower = text.toLowerCase();
  for (const table of tables) {
    if (!/^[\x20-\x7e]*$/.test(table) || lower.includes(table)) {
      return true;
    }
  }
  return false;
}

// ============================================================================
// Formulas
// ============================================================================

function isReferenceToken(token: Token | undefined): boolean {
  return (
    token !== undefined &&
    (token.type === TokenType.CellRef ||
      token.type === TokenType.Range ||
      token.type === TokenType.ColRange ||
      token.type === TokenType.RowRange)
  );
}

/** A token a sheet prefix qualifies: a reference, or a sheet-scoped name. */
function isQualifiable(token: Token | undefined): boolean {
  return isReferenceToken(token) || token?.type === TokenType.Name;
}

interface Replacement {
  start: number;
  end: number;
  text: string;
}

/** What an edit does to a sheet prefix and what it qualifies. */
type QualifiedResult =
  | { readonly kind: "prefix"; readonly text: string }
  | { readonly kind: "reference"; readonly text: string }
  | { readonly kind: "error" }
  | undefined;

/**
 * Rewrite every reference in a formula through `edit`.
 *
 * `host` is the sheet the formula lives on, which an unqualified reference belongs to; pass `undefined` where an
 * unqualified reference has no fixed sheet (a defined name, a chart series), and those are left alone. A leading `=`
 * and everything the edit does not touch are kept exactly; text that cannot be tokenised is returned unchanged.
 */
export function rewriteFormula(formula: string, edit: ReferenceEdit, host?: string): string {
  const local =
    (edit.kind === "rows" || edit.kind === "columns") &&
    host !== undefined &&
    sameSheet(host, edit.sheet);
  const removedTables = edit.kind === "rename" ? undefined : edit.removedTables;
  if (!local && !mayNameSheet(formula, edit) && !mayNameTable(formula, removedTables)) {
    return formula;
  }
  let tokens: Token[];
  let spans: Array<readonly [number, number]>;
  try {
    ({ tokens, spans } = tokenizeWithSpans(formula));
  } catch {
    return formula;
  }
  const replacements: Replacement[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.type === TokenType.SheetRef) {
      const next = tokens[i + 1];
      const qualifies = isQualifiable(next);
      const result = rewriteQualified(token, next, formula, spans[i + 1], edit);
      if (result?.kind === "prefix") {
        replacements.push({ start: spans[i][0], end: spans[i][1], text: result.text });
      } else if (result?.kind === "reference") {
        replacements.push({ start: spans[i + 1][0], end: spans[i + 1][1], text: result.text });
      } else if (result?.kind === "error") {
        replacements.push({
          start: spans[i][0],
          end: qualifies ? spans[i + 1][1] : spans[i][1],
          text: REF_ERROR
        });
      }
      if (qualifies) {
        i++;
      }
      continue;
    }
    if (
      token.type === TokenType.StructuredRef &&
      token.tableName !== "" &&
      removedTables?.has(excelNameKey(token.tableName))
    ) {
      replacements.push({ start: spans[i][0], end: spans[i][1], text: REF_ERROR });
      continue;
    }
    if (local && isReferenceToken(token)) {
      const edited = editReferenceText(formula.slice(spans[i][0], spans[i][1]), edit as LineEdit);
      if (edited !== undefined) {
        replacements.push({ start: spans[i][0], end: spans[i][1], text: edited });
      }
    }
  }
  if (replacements.length === 0) {
    return formula;
  }
  // A spill-range reference `A1#` whose anchor is gone is `#REF!`, not `#REF!#`: the operator goes with its operand.
  for (const replacement of replacements) {
    if (replacement.text !== REF_ERROR) {
      continue;
    }
    const next = spans.findIndex(([start]) => start === replacement.end);
    const token = tokens[next];
    if (token !== undefined && isSpillOperator(token)) {
      replacement.end = spans[next][1];
    }
  }
  let out = "";
  let last = 0;
  for (const { start, end, text } of replacements) {
    out += formula.slice(last, start) + text;
    last = end;
  }
  return out + formula.slice(last);
}

/** A reference token's text after a line edit on its sheet, or `undefined` when it is unaffected. */
function editReferenceText(text: string, edit: LineEdit): string | undefined {
  const area = parseArea(text);
  return area === undefined ? undefined : editArea(area, edit);
}

function rewriteQualified(
  token: Extract<Token, { type: TokenType.SheetRef }>,
  next: Token | undefined,
  formula: string,
  nextSpan: readonly [number, number] | undefined,
  edit: ReferenceEdit
): QualifiedResult {
  const { sheetName, endSheetName } = token;
  switch (edit.kind) {
    case "rename": {
      const start = sameSheet(sheetName, edit.from) ? edit.to : sheetName;
      const end =
        endSheetName !== undefined && sameSheet(endSheetName, edit.from) ? edit.to : endSheetName;
      return start === sheetName && end === endSheetName
        ? undefined
        : { kind: "prefix", text: sheetPrefix(start, end) };
    }
    case "removeSheet":
      if (endSheetName === undefined) {
        return sameSheet(sheetName, edit.sheet) ? { kind: "error" } : undefined;
      }
      return rewriteSpanOnRemoval(sheetName, endSheetName, edit);
    default: {
      // A 3-D reference is left alone: only one of the sheets it spans moved.
      if (
        endSheetName !== undefined ||
        !sameSheet(sheetName, edit.sheet) ||
        !isReferenceToken(next) ||
        nextSpan === undefined
      ) {
        return undefined;
      }
      const edited = editReferenceText(formula.slice(nextSpan[0], nextSpan[1]), edit);
      if (edited === undefined) {
        return undefined;
      }
      return edited === REF_ERROR ? { kind: "error" } : { kind: "reference", text: edited };
    }
  }
}

/** A 3-D span `start:end` after `edit.sheet` is removed. */
function rewriteSpanOnRemoval(start: string, end: string, edit: SheetRemoveEdit): QualifiedResult {
  const isStart = sameSheet(start, edit.sheet);
  const isEnd = sameSheet(end, edit.sheet);
  if (!isStart && !isEnd) {
    return undefined;
  }
  const order = edit.sheets;
  const startIndex = order.findIndex(name => sameSheet(name, start));
  const endIndex = order.findIndex(name => sameSheet(name, end));
  if (startIndex < 0 || endIndex < 0 || (isStart && isEnd)) {
    return { kind: "error" };
  }
  const low = Math.min(startIndex, endIndex);
  const high = Math.max(startIndex, endIndex);
  const removed = isStart ? startIndex : endIndex;
  const newLow = removed === low ? low + 1 : low;
  const newHigh = removed === high ? high - 1 : high;
  if (newLow > newHigh) {
    return { kind: "error" };
  }
  return {
    kind: "prefix",
    text:
      newLow === newHigh ? sheetPrefix(order[newLow]) : sheetPrefix(order[newLow], order[newHigh])
  };
}

// ============================================================================
// Table renames
// ============================================================================

/**
 * `formula` with every reference to a renamed table restated under its new name: the table name of a structured
 * reference (`Sales[Amt]`, `Sales[#All]`) and a table named on its own (`ROWS(Sales)`), compared as Excel compares
 * names. `renames` maps each old name's `excelNameKey` to the new name. String literals, column names inside the
 * brackets and everything else are kept exactly; text that cannot be tokenised is returned unchanged.
 */
export function renameTablesInFormula(
  formula: string,
  renames: ReadonlyMap<string, string>
): string {
  if (renames.size === 0) {
    return formula;
  }
  let tokens: Token[];
  let spans: Array<readonly [number, number]>;
  try {
    ({ tokens, spans } = tokenizeWithSpans(formula));
  } catch {
    return formula;
  }
  let out = "";
  let last = 0;
  tokens.forEach((token, i) => {
    const name =
      token.type === TokenType.StructuredRef
        ? token.tableName
        : token.type === TokenType.Name
          ? token.value
          : "";
    const renamed = name === "" ? undefined : renames.get(excelNameKey(name));
    if (renamed === undefined) {
      return;
    }
    const [start] = spans[i];
    // The table name opens the token's text: `Sales[Amt]`, or the whole of `Sales`.
    if (formula.slice(start, start + name.length) !== name) {
      return;
    }
    out += formula.slice(last, start) + renamed;
    last = start + name.length;
  });
  return last === 0 ? formula : out + formula.slice(last);
}
