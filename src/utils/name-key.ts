/**
 * How Excel compares names that it treats as case-insensitive. Two rules were measured against Excel 16, and they
 * differ by one thing: whether a character may stand for several.
 *
 * - {@link excelNameKey} is used for table names, table column names (and the column of a structured reference) and
 *   defined names. `ß` = `ẞ` = `SS`, `ﬁ` = `FI`.
 * - {@link sheetNameKey} is used for sheet names (and the sheet of a reference) and `LET` / `LAMBDA` parameter names.
 *   A character is folded only to one character, so `ß` = `ẞ` but `ß` ≠ `SS`, and `ﬁ` ≠ `FI`.
 *
 * Both rules agree on everything else. Compatibility forms fold to the plain character (`ａ` = `A`, `x²` = `x2`,
 * `Ⅰ` = `I`, `ｶ` = `カ`). Hiragana equals katakana (`あ` = `ア`), but the iteration marks `ゝ` and `ヽ` stay apart. Case
 * folds one character at a time, so `ς` = `σ` = `Σ`, the Kelvin sign = `k` and `Ǳ` = `ǳ`. Canonically equivalent
 * spellings agree, so `e` + U+0301 = `é`. Accents still count (`é` ≠ `e`), and so does the dotless `ı` (≠ `I`, ≠ `i`).
 * `İ` is not `i` (the name rule makes it `i` + U+0307). A character whose capital is two letters keeps its own identity (`ᾳ` ≠ `ΑΙ`).
 *
 * The measurements came from typing headers, renaming tables and sheets, defining names, and calling `LET` and
 * structured references in Excel.
 */

const HIRAGANA_RE = /[\u3041-\u3096]/g;

function foldKana(text: string): string {
  return text.replace(HIRAGANA_RE, ch => String.fromCharCode(ch.charCodeAt(0) + 0x60));
}

/** The key two table, column or defined names are compared by: equal keys are the same name to Excel. */
export function excelNameKey(name: string): string {
  let key = "";
  for (const ch of name.normalize("NFKC")) {
    // Lowered one character at a time: `toLowerCase` lowers a word-final `Σ` to `ς`.
    key += ch === "ß" || ch === "ẞ" ? "ss" : ch === "ς" ? "σ" : ch.toLowerCase();
  }
  return foldKana(key.normalize("NFC"));
}

/** The key two sheet names, or two `LET` / `LAMBDA` parameter names, are compared by. */
export function sheetNameKey(name: string): string {
  let key = "";
  for (const ch of name.normalize("NFC")) {
    const compatible = ch.normalize("NFKC");
    const single = [...compatible].length === 1 ? compatible : ch;
    const lower = single === "ς" ? "σ" : single.toLowerCase();
    key += [...lower].length === 1 ? lower : single;
  }
  return foldKana(key.normalize("NFC"));
}
