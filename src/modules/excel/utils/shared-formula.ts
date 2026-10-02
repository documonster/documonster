import { colCache } from "@excel/utils/col-cache";

function slideFormula(formula: string, fromCell: string, toCell: string): string {
  const from = colCache.decodeAddress(fromCell);
  const to = colCache.decodeAddress(toCell);
  const shift: Shift = { col: to.col - from.col, row: to.row - from.row };
  // A reference is never inside a string (`"A1"`), a quoted sheet name (`'Q1 A1'!B2`) or brackets — a structured
  // reference's column (`Table1[Col A1]`) or an external workbook index (`[1]Sheet1!A1`). Sliding any of those changes
  // what the formula computes, so only the text between them is slid.
  let out = "";
  let plainStart = 0;
  let i = 0;
  while (i < formula.length) {
    const c = formula[i];
    if (c !== '"' && c !== "'" && c !== "[") {
      i++;
      continue;
    }
    out += slideReferences(formula.slice(plainStart, i), shift);
    const literalStart = i;
    i = c === "[" ? skipBrackets(formula, i) : skipQuoted(formula, i, c);
    out += formula.slice(literalStart, i);
    plainStart = i;
  }
  return out + slideReferences(formula.slice(plainStart), shift);
}

/** Index just past the quoted run starting at `start`; a doubled quote is an escaped one. */
function skipQuoted(formula: string, start: number, quote: string): number {
  let i = start + 1;
  while (i < formula.length) {
    if (formula[i] === quote) {
      if (formula[i + 1] !== quote) {
        return i + 1;
      }
      i++;
    }
    i++;
  }
  return i;
}

/** Index just past the bracketed run starting at `start`, nested brackets included; `'` escapes the next character. */
function skipBrackets(formula: string, start: number): number {
  let depth = 0;
  let i = start;
  while (i < formula.length) {
    const c = formula[i];
    if (c === "'") {
      i += 2;
      continue;
    }
    if (c === "[") {
      depth++;
    } else if (c === "]" && --depth === 0) {
      return i + 1;
    }
    i++;
  }
  return i;
}

interface Shift {
  col: number;
  row: number;
}

const MAX_ROW = 1048576;
const MAX_COL = 16384;

/** A run that could be a reference, a name, a function or a number — references are told apart below. */
const WORD_RX = /[A-Za-z0-9_.\\?$]+/g;
const CELL_RX = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/;
const COLUMN_RX = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_RX = /^(\$?)([1-9]\d*)$/;

/**
 * Slide every reference in `text`, which holds no string, quoted name or bracket.
 *
 * A word is a reference only as a whole: not when it is a function (`LOG10(`), a sheet (`AB2!`), part of a name that
 * runs on in a character outside the word class, or a column or row standing alone — `A` and `1` are references only
 * as `A:A` and `1:1`. A reference slid off the sheet becomes `#REF!`, as Excel makes it; for a range, the whole range.
 */
function slideReferences(text: string, shift: Shift): string {
  let out = "";
  let last = 0;
  WORD_RX.lastIndex = 0;
  for (let m = WORD_RX.exec(text); m !== null; m = WORD_RX.exec(text)) {
    const start = m.index;
    let end = start + m[0].length;
    if (!isReferenceWord(text, start, end)) {
      continue;
    }
    let slid: string | undefined;
    if (text[end] === ":") {
      WORD_RX.lastIndex = end + 1;
      const second = WORD_RX.exec(text);
      if (second !== null && second.index === end + 1) {
        const secondEnd = second.index + second[0].length;
        if (isReferenceWord(text, second.index, secondEnd)) {
          slid = slideRange(m[0], second[0], shift);
          if (slid !== undefined) {
            end = secondEnd;
          }
        }
      }
      WORD_RX.lastIndex = end;
    }
    slid ??= slideCell(m[0], shift);
    if (slid !== undefined) {
      out += text.slice(last, start) + slid;
      last = end;
    }
  }
  return out + text.slice(last);
}

/** Whether the word at `start`–`end` stands on its own: not a function, not a sheet, not part of a longer name. */
function isReferenceWord(text: string, start: number, end: number): boolean {
  const next = text.charCodeAt(end);
  return (
    next !== 40 && // (
    next !== 33 && // !
    !(next > 127) &&
    !(text.charCodeAt(start - 1) > 127)
  );
}

function slideCell(word: string, shift: Shift): string | undefined {
  const m = CELL_RX.exec(word);
  if (m === null) {
    return undefined;
  }
  const col = colCache.decodeCol(m[2].toUpperCase());
  if (col > MAX_COL) {
    return undefined; // a name such as `ABCD1`, not a reference
  }
  const newCol = m[1] ? col : col + shift.col;
  const newRow = m[3] ? Number(m[4]) : Number(m[4]) + shift.row;
  if (newCol < 1 || newCol > MAX_COL || newRow < 1 || newRow > MAX_ROW) {
    return "#REF!";
  }
  return m[1] + colCache.n2l(newCol) + m[3] + newRow;
}

/** `A1:B2`, `A:B` or `1:2` slid as a unit; `undefined` when the pair is not one of those. */
function slideRange(first: string, second: string, shift: Shift): string | undefined {
  const cellA = slideCell(first, shift);
  const cellB = slideCell(second, shift);
  if (cellA !== undefined && cellB !== undefined) {
    return cellA === "#REF!" || cellB === "#REF!" ? "#REF!" : `${cellA}:${cellB}`;
  }
  const colA = COLUMN_RX.exec(first);
  const colB = COLUMN_RX.exec(second);
  if (colA && colB) {
    const a = slideLine(colA, shift.col, MAX_COL, letters =>
      colCache.decodeCol(letters.toUpperCase())
    );
    const b = slideLine(colB, shift.col, MAX_COL, letters =>
      colCache.decodeCol(letters.toUpperCase())
    );
    if (a === undefined || b === undefined) {
      return undefined;
    }
    return a === null || b === null
      ? "#REF!"
      : `${a.dollar}${colCache.n2l(a.n)}:${b.dollar}${colCache.n2l(b.n)}`;
  }
  const rowA = ROW_RX.exec(first);
  const rowB = ROW_RX.exec(second);
  if (rowA && rowB) {
    const a = slideLine(rowA, shift.row, MAX_ROW, Number);
    const b = slideLine(rowB, shift.row, MAX_ROW, Number);
    if (a === undefined || b === undefined) {
      return undefined;
    }
    return a === null || b === null ? "#REF!" : `${a.dollar}${a.n}:${b.dollar}${b.n}`;
  }
  return undefined;
}

/** One end of a whole-column or whole-row range: `undefined` if it is not one, `null` if slid off the sheet. */
function slideLine(
  m: RegExpExecArray,
  delta: number,
  max: number,
  parse: (s: string) => number
): { dollar: string; n: number } | null | undefined {
  const n = parse(m[2]);
  if (n > max) {
    return undefined;
  }
  const slid = m[1] ? n : n + delta;
  return slid < 1 || slid > max ? null : { dollar: m[1], n: slid };
}

export { slideFormula };
