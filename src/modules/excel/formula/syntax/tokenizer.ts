/**
 * Formula Tokenizer
 *
 * Converts an Excel formula string into a stream of tokens.
 * Supports cell references (A1, $A$1), ranges (A1:B10),
 * cross-sheet references (Sheet1!A1), operators, function calls,
 * numbers, strings, booleans, and error literals.
 */

import { TokenType, type Token } from "./token-types";

// ============================================================================
// Character Helpers
// ============================================================================

function isDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

function isAlpha(ch: string): boolean {
  return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z");
}

function isAlphaNumOrUnderscore(ch: string): boolean {
  return isAlpha(ch) || isDigit(ch) || ch === "_" || ch === ".";
}

// ============================================================================
// Cell Reference Detection
// ============================================================================

/**
 * Check if a string is a valid cell reference (like A1, $B$2, XFD1048576).
 * Returns null if not a cell ref, otherwise returns parsed info.
 */
function parseCellRef(
  s: string
): { col: string; row: string; colAbsolute: boolean; rowAbsolute: boolean } | null {
  let i = 0;
  let colAbsolute = false;
  let rowAbsolute = false;

  if (i < s.length && s[i] === "$") {
    colAbsolute = true;
    i++;
  }

  const colStart = i;
  while (i < s.length && isAlpha(s[i])) {
    i++;
  }
  const colPart = s.slice(colStart, i).toUpperCase();
  if (colPart.length === 0 || colPart.length > 3) {
    return null;
  }
  // Validate column is <= XFD
  if (colPart.length === 3 && colPart > "XFD") {
    return null;
  }

  if (i < s.length && s[i] === "$") {
    rowAbsolute = true;
    i++;
  }

  const rowStart = i;
  while (i < s.length && isDigit(s[i])) {
    i++;
  }
  const rowPart = s.slice(rowStart, i);
  if (rowPart.length === 0 || i !== s.length) {
    return null;
  }
  const rowNum = parseInt(rowPart, 10);
  if (rowNum < 1 || rowNum > 1048576) {
    return null;
  }

  return { col: colPart, row: rowPart, colAbsolute, rowAbsolute };
}

// ============================================================================
// Structured Reference Bracket Parser
// ============================================================================

/** Valid special items in structured references */
const STRUCTURED_REF_SPECIALS = new Set(["#All", "#Data", "#Headers", "#Totals", "#This Row"]);

/**
 * Parse the bracketed portion of a structured reference starting at position `pos`.
 * `pos` must point to the opening `[`.
 *
 * Handles:
 *   [Column]                    → simple column
 *   [#Headers]                  → special item
 *   [[#Headers],[Column]]       → nested: special + column
 *   [[#This Row],[Column]]      → nested: special + column
 *   [[Col1]:[Col2]]             → column range
 *   [@Column]                   → shorthand for [#This Row],[Column]
 *   [@[Column]]                 → alternative shorthand
 *
 * Returns parsed specials, columns, and the position after the closing `]`.
 */
function parseStructuredRefBrackets(
  formula: string,
  pos: number
): { specials: string[]; columns: string[]; end: number } {
  const len = formula.length;
  if (pos >= len || formula[pos] !== "[") {
    throw new Error("Expected '[' at position " + pos);
  }

  const specials: string[] = [];
  const columns: string[] = [];
  let i = pos + 1; // skip opening [

  // Skip whitespace inside brackets
  while (i < len && (formula[i] === " " || formula[i] === "\t")) {
    i++;
  }

  // Check for @ shorthand: [@Column] or [@[Column]]
  if (i < len && formula[i] === "@") {
    specials.push("#This Row");
    i++; // skip @
    // Skip whitespace
    while (i < len && (formula[i] === " " || formula[i] === "\t")) {
      i++;
    }
    if (i < len && formula[i] === "[") {
      // [@[Column]] form — inner bracketed column name
      i++; // skip [
      const colName = readBracketedItem(formula, i);
      columns.push(colName.value);
      i = colName.end; // after inner ]
    } else if (i < len && formula[i] !== "]") {
      // [@Column] form — unbracketed column name until ]
      const start = i;
      while (i < len && formula[i] !== "]") {
        i++;
      }
      const name = formula.slice(start, i).trim();
      if (name.length > 0) {
        columns.push(name);
      }
    }
    // Skip whitespace before closing ]
    while (i < len && (formula[i] === " " || formula[i] === "\t")) {
      i++;
    }
    if (i < len && formula[i] === "]") {
      i++; // skip closing ]
    }
    return { specials, columns, end: i };
  }

  // Check for nested brackets: [[...],[...]] or [[Col1]:[Col2]]
  if (i < len && formula[i] === "[") {
    // Parse comma-separated inner bracket items
    while (i < len && formula[i] === "[") {
      const item = readBracketedItem(formula, i + 1);
      const value = item.value;
      i = item.end; // after the inner ]

      if (STRUCTURED_REF_SPECIALS.has(value)) {
        specials.push(value);
      } else {
        columns.push(value);
      }

      // Skip whitespace
      while (i < len && (formula[i] === " " || formula[i] === "\t")) {
        i++;
      }

      // Check for : (column range) or , (next item)
      if (i < len && formula[i] === ":") {
        i++; // skip :
        // Skip whitespace
        while (i < len && (formula[i] === " " || formula[i] === "\t")) {
          i++;
        }
        // Next must be a [Column]
        if (i < len && formula[i] === "[") {
          const item2 = readBracketedItem(formula, i + 1);
          columns.push(item2.value);
          i = item2.end;
        }
        // Skip whitespace
        while (i < len && (formula[i] === " " || formula[i] === "\t")) {
          i++;
        }
      } else if (i < len && formula[i] === ",") {
        i++; // skip ,
        // Skip whitespace
        while (i < len && (formula[i] === " " || formula[i] === "\t")) {
          i++;
        }
      }
    }

    // Skip whitespace before closing ]
    while (i < len && (formula[i] === " " || formula[i] === "\t")) {
      i++;
    }
    if (i < len && formula[i] === "]") {
      i++; // skip outer closing ]
    }
    return { specials, columns, end: i };
  }

  // Check for special item: [#Headers], [#Data], etc.
  if (i < len && formula[i] === "#") {
    const start = i;
    while (i < len && formula[i] !== "]") {
      i++;
    }
    const value = formula.slice(start, i).trim();
    if (STRUCTURED_REF_SPECIALS.has(value)) {
      specials.push(value);
    }
    if (i < len && formula[i] === "]") {
      i++; // skip closing ]
    }
    return { specials, columns, end: i };
  }

  // Simple column reference: [Column]
  const start = i;
  while (i < len && formula[i] !== "]") {
    i++;
  }
  const name = formula.slice(start, i).trim();
  if (name.length > 0) {
    columns.push(name);
  }
  if (i < len && formula[i] === "]") {
    i++; // skip closing ]
  }
  return { specials, columns, end: i };
}

/**
 * Read a single bracketed item starting after the opening `[`.
 * Returns the text content and position after the closing `]`.
 * Handles nested `'` escaping within bracket content (Excel uses `'` to escape
 * special chars like `]`, `[`, `#` inside column names — but this is rare).
 */
function readBracketedItem(formula: string, pos: number): { value: string; end: number } {
  const len = formula.length;
  let i = pos;
  let result = "";
  while (i < len && formula[i] !== "]") {
    if (formula[i] === "'" && i + 1 < len) {
      // Escape: the next character is literal
      result += formula[i + 1];
      i += 2;
    } else {
      result += formula[i];
      i++;
    }
  }
  if (i < len && formula[i] === "]") {
    i++; // skip ]
  }
  return { value: result.trim(), end: i };
}

// ============================================================================
// Tokenizer
// ============================================================================

export function tokenize(formula: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const len = formula.length;

  // Track whether previous token could produce a value
  // (used to distinguish unary vs binary +/-)
  function lastTokenIsValue(): boolean {
    if (tokens.length === 0) {
      return false;
    }
    const last = tokens[tokens.length - 1];
    return (
      last.type === TokenType.Number ||
      last.type === TokenType.String ||
      last.type === TokenType.Boolean ||
      last.type === TokenType.CellRef ||
      last.type === TokenType.Range ||
      last.type === TokenType.CloseParen ||
      last.type === TokenType.CloseBrace ||
      last.type === TokenType.Percent ||
      last.type === TokenType.Name ||
      last.type === TokenType.Error ||
      last.type === TokenType.ColRange ||
      last.type === TokenType.RowRange ||
      last.type === TokenType.StructuredRef
    );
  }

  while (i < len) {
    const ch = formula[i];

    // Skip whitespace
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      i++;
      continue;
    }

    // String literals
    if (ch === '"') {
      i++; // skip opening quote
      let str = "";
      while (i < len) {
        if (formula[i] === '"') {
          if (i + 1 < len && formula[i + 1] === '"') {
            // Escaped quote
            str += '"';
            i += 2;
          } else {
            i++; // skip closing quote
            break;
          }
        } else {
          str += formula[i];
          i++;
        }
      }
      tokens.push({ type: TokenType.String, value: str });
      continue;
    }

    // Error literals (#N/A, #REF!, etc.)
    if (ch === "#") {
      let errStr = "#";
      i++;
      while (
        i < len &&
        formula[i] !== " " &&
        formula[i] !== "," &&
        formula[i] !== ")" &&
        formula[i] !== "+" &&
        formula[i] !== "-" &&
        formula[i] !== "*" &&
        formula[i] !== "/" &&
        formula[i] !== "^" &&
        formula[i] !== "&" &&
        formula[i] !== "=" &&
        formula[i] !== "<" &&
        formula[i] !== ">" &&
        formula[i] !== "}"
      ) {
        errStr += formula[i];
        i++;
      }
      tokens.push({ type: TokenType.Error, value: errStr });
      continue;
    }

    // External workbook reference: [Book.xlsx]Sheet!A1
    // When [ appears and is NOT preceded by a value token (not array literal context),
    // consume the bracketed workbook name and subsequent sheet!ref as a Name token.
    // This allows the engine to fall back to cached results gracefully.
    if (ch === "[" && !lastTokenIsValue()) {
      // Peek: if this looks like [filename]Sheet!... consume as external ref
      const bracketStart = i;
      i++; // skip [
      let hasClose = false;
      while (i < len) {
        if (formula[i] === "]") {
          hasClose = true;
          i++;
          break;
        }
        i++;
      }
      if (hasClose) {
        // Consume the rest as sheet!cellref (Name token for external ref)
        const refStart = bracketStart;
        while (
          i < len &&
          formula[i] !== "+" &&
          formula[i] !== "-" &&
          formula[i] !== "*" &&
          formula[i] !== "/" &&
          formula[i] !== "^" &&
          formula[i] !== "&" &&
          formula[i] !== "=" &&
          formula[i] !== "<" &&
          formula[i] !== ">" &&
          formula[i] !== ")" &&
          formula[i] !== "," &&
          formula[i] !== " " &&
          formula[i] !== "}" &&
          formula[i] !== ":" &&
          formula[i] !== "%" &&
          formula[i] !== ";" &&
          formula[i] !== "{"
        ) {
          i++;
        }
        tokens.push({ type: TokenType.Name, value: formula.slice(refStart, i) });
        continue;
      }
      // Not a valid external ref, backtrack
      i = bracketStart;
    }

    // Array constant braces
    if (ch === "{") {
      tokens.push({ type: TokenType.OpenBrace });
      i++;
      continue;
    }
    if (ch === "}") {
      tokens.push({ type: TokenType.CloseBrace });
      i++;
      continue;
    }

    // Semicolons (array row separator)
    if (ch === ";") {
      tokens.push({ type: TokenType.Semicolon });
      i++;
      continue;
    }

    // Numbers (and potential whole-row ranges like 1:5)
    if (isDigit(ch) || (ch === "." && i + 1 < len && isDigit(formula[i + 1]))) {
      const start = i;
      while (i < len && isDigit(formula[i])) {
        i++;
      }
      // Check for whole-row range: pure integer followed by : and another integer
      const isInteger = i > start && formula[start] !== "." && i < len && formula[i] === ":";
      if (isInteger) {
        const row1Str = formula.slice(start, i);
        const row1 = parseInt(row1Str, 10);
        if (row1 >= 1 && row1 <= 1048576) {
          const colonPos = i;
          i++; // skip :
          // Optional $ before second row number
          const start2 = i;
          if (i < len && formula[i] === "$") {
            i++;
          }
          const rowStart2 = i;
          while (i < len && isDigit(formula[i])) {
            i++;
          }
          if (i > rowStart2) {
            const row2Str = formula.slice(rowStart2, i);
            const row2 = parseInt(row2Str, 10);
            if (row2 >= 1 && row2 <= 1048576) {
              tokens.push({
                type: TokenType.RowRange,
                value: row1Str + ":" + formula.slice(start2, i)
              });
              continue;
            }
          }
          // Not a valid row range, backtrack
          i = colonPos;
        }
      }
      if (i < len && formula[i] === ".") {
        i++;
        while (i < len && isDigit(formula[i])) {
          i++;
        }
      }
      // Scientific notation
      if (i < len && (formula[i] === "E" || formula[i] === "e")) {
        i++;
        if (i < len && (formula[i] === "+" || formula[i] === "-")) {
          i++;
        }
        while (i < len && isDigit(formula[i])) {
          i++;
        }
      }
      tokens.push({ type: TokenType.Number, value: formula.slice(start, i) });
      continue;
    }

    // Quoted sheet name: 'Sheet Name'! or 3D ref 'Sheet1:Sheet3'!
    if (ch === "'") {
      i++; // skip opening quote
      let sheetName = "";
      while (i < len) {
        if (formula[i] === "'") {
          if (i + 1 < len && formula[i + 1] === "'") {
            sheetName += "'";
            i += 2;
          } else {
            i++; // skip closing quote
            break;
          }
        } else {
          sheetName += formula[i];
          i++;
        }
      }
      // Expect ! after
      if (i < len && formula[i] === "!") {
        i++; // skip !
      }
      // Check for 3D reference: sheetName contains an unquoted colon separating two sheet names
      // e.g. 'Sheet1:Sheet3' → startSheet="Sheet1", endSheet="Sheet3"
      const colonIdx = sheetName.indexOf(":");
      if (colonIdx > 0 && colonIdx < sheetName.length - 1) {
        const startSheet = sheetName.slice(0, colonIdx);
        const endSheet = sheetName.slice(colonIdx + 1);
        tokens.push({
          type: TokenType.SheetRef,
          sheetName: startSheet,
          endSheetName: endSheet
        });
      } else {
        tokens.push({
          type: TokenType.SheetRef,
          sheetName
        });
      }
      continue;
    }

    // Standalone structured reference: [@Column] or [[#This Row],[Column]]
    if (ch === "[") {
      const sr = parseStructuredRefBrackets(formula, i);
      tokens.push({
        type: TokenType.StructuredRef,
        tableName: "",
        columns: sr.columns,
        specials: sr.specials
      });
      i = sr.end;
      continue;
    }

    // Identifiers: cell refs, function names, booleans, named ranges, sheet refs
    if (isAlpha(ch) || ch === "$" || ch === "_") {
      const start = i;

      // Collect $-prefixed or alpha-numeric identifier
      while (i < len && (isAlphaNumOrUnderscore(formula[i]) || formula[i] === "$")) {
        i++;
      }

      const word = formula.slice(start, i);

      // Check for sheet reference: WORD!
      if (i < len && formula[i] === "!") {
        i++; // skip !
        tokens.push({
          type: TokenType.SheetRef,
          sheetName: word
        });
        continue;
      }

      // Check for 3D sheet reference: WORD:WORD! (e.g. Sheet1:Sheet3!A1)
      if (i < len && formula[i] === ":") {
        const colonPos3D = i;
        let j = i + 1;
        // Collect the second sheet name (alphanumeric + underscore + .)
        while (j < len && (isAlphaNumOrUnderscore(formula[j]) || formula[j] === "$")) {
          j++;
        }
        if (j > colonPos3D + 1 && j < len && formula[j] === "!") {
          const endSheet = formula.slice(colonPos3D + 1, j);
          // Verify endSheet is not a cell reference (contains alpha chars, not purely dollar+digits)
          const hasAlpha = /[A-Za-z]/.test(endSheet);
          if (hasAlpha) {
            i = j + 1; // skip past !
            tokens.push({
              type: TokenType.SheetRef,
              sheetName: word,
              endSheetName: endSheet
            });
            continue;
          }
        }
      }

      // Check for structured reference: WORD[...]
      if (i < len && formula[i] === "[") {
        const sr = parseStructuredRefBrackets(formula, i);
        tokens.push({
          type: TokenType.StructuredRef,
          tableName: word,
          columns: sr.columns,
          specials: sr.specials
        });
        i = sr.end;
        continue;
      }

      // Check for function call: WORD(
      if (i < len && formula[i] === "(") {
        tokens.push({ type: TokenType.Function, name: word.toUpperCase() });
        // Don't consume the '(' — it will be consumed in the next iteration
        continue;
      }

      // Check for boolean literals
      const upper = word.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") {
        tokens.push({ type: TokenType.Boolean, value: upper as "TRUE" | "FALSE" });
        continue;
      }

      // Check for cell reference
      const ref = parseCellRef(word);
      if (ref) {
        const cellToken: Token = {
          type: TokenType.CellRef,
          col: ref.col,
          row: ref.row,
          colAbsolute: ref.colAbsolute,
          rowAbsolute: ref.rowAbsolute
        };
        tokens.push(cellToken);
        // Check if this is part of a range (A1:B2)
        if (i < len && formula[i] === ":") {
          // Peek ahead for another cell ref
          const colonPos = i;
          i++; // skip :
          const rangeStart = i;
          while (i < len && (isAlphaNumOrUnderscore(formula[i]) || formula[i] === "$")) {
            i++;
          }
          const secondWord = formula.slice(rangeStart, i);
          const ref2 = parseCellRef(secondWord);
          if (ref2) {
            // It's a range — replace the last CellRef token with a Range token
            tokens[tokens.length - 1] = {
              type: TokenType.Range,
              value:
                (ref.colAbsolute ? "$" : "") +
                ref.col +
                (ref.rowAbsolute ? "$" : "") +
                ref.row +
                ":" +
                (ref2.colAbsolute ? "$" : "") +
                ref2.col +
                (ref2.rowAbsolute ? "$" : "") +
                ref2.row
            };
          } else {
            // Not a valid range, push : as operator and backtrack
            i = colonPos; // backtrack — let : be handled normally
          }
        }
        continue;
      }

      // Otherwise it's a named range / defined name — but check for whole-column range first (A:B, $A:$B)
      // A whole-column ref looks like: [pure-alpha or $+alpha] followed by ':'
      if (i < len && formula[i] === ":") {
        // Check if 'word' is a pure column reference (optional $ + letters only, no digits)
        const colRefMatch = /^(\$?)([A-Za-z]{1,3})$/.exec(word);
        if (colRefMatch) {
          const col1Abs = colRefMatch[1] === "$";
          const col1 = colRefMatch[2].toUpperCase();
          // Validate column <= XFD
          if (col1.length <= 3 && (col1.length < 3 || col1 <= "XFD")) {
            const colonPos = i;
            i++; // skip :
            // Read the second column ref
            const start2 = i;
            if (i < len && formula[i] === "$") {
              i++;
            }
            const colStart2 = i;
            while (i < len && isAlpha(formula[i])) {
              i++;
            }
            // Ensure it's ONLY letters (no digits follow = not a cell ref)
            const afterLetters = i;
            const isFollowedByDigit = i < len && isDigit(formula[i]);
            if (!isFollowedByDigit && afterLetters > colStart2) {
              const part2 = formula.slice(start2, i);
              const col2Match = /^(\$?)([A-Za-z]{1,3})$/.exec(part2);
              if (col2Match) {
                const col2 = col2Match[2].toUpperCase();
                if (col2.length <= 3 && (col2.length < 3 || col2 <= "XFD")) {
                  tokens.push({
                    type: TokenType.ColRange,
                    value: (col1Abs ? "$" : "") + col1 + ":" + part2.toUpperCase()
                  });
                  continue;
                }
              }
            }
            // Not a valid column range, backtrack
            i = colonPos;
          }
        }
        // Check for absolute whole-row range: $1:$5
        const absRowMatch = /^(\$)(\d+)$/.exec(word);
        if (absRowMatch) {
          const row1 = parseInt(absRowMatch[2], 10);
          if (row1 >= 1 && row1 <= 1048576) {
            const colonPos = i;
            i++; // skip :
            const start2 = i;
            if (i < len && formula[i] === "$") {
              i++;
            }
            const rowStart2 = i;
            while (i < len && isDigit(formula[i])) {
              i++;
            }
            if (i > rowStart2) {
              const row2 = parseInt(formula.slice(rowStart2, i), 10);
              if (row2 >= 1 && row2 <= 1048576) {
                tokens.push({
                  type: TokenType.RowRange,
                  value: word + ":" + formula.slice(start2, i)
                });
                continue;
              }
            }
            i = colonPos;
          }
        }
      }
      tokens.push({ type: TokenType.Name, value: word });
      continue;
    }

    // Operators and punctuation
    if (ch === "(") {
      tokens.push({ type: TokenType.OpenParen });
      i++;
      continue;
    }
    if (ch === ")") {
      tokens.push({ type: TokenType.CloseParen });
      i++;
      continue;
    }
    if (ch === ",") {
      tokens.push({ type: TokenType.Comma });
      i++;
      continue;
    }
    if (ch === ":") {
      tokens.push({ type: TokenType.Colon });
      i++;
      continue;
    }
    if (ch === "%") {
      tokens.push({ type: TokenType.Percent });
      i++;
      continue;
    }

    // Multi-character operators
    if (ch === "<") {
      if (i + 1 < len && formula[i + 1] === "=") {
        tokens.push({ type: TokenType.Operator, value: "<=" });
        i += 2;
        continue;
      }
      if (i + 1 < len && formula[i + 1] === ">") {
        tokens.push({ type: TokenType.Operator, value: "<>" });
        i += 2;
        continue;
      }
      tokens.push({ type: TokenType.Operator, value: "<" });
      i++;
      continue;
    }
    if (ch === ">") {
      if (i + 1 < len && formula[i + 1] === "=") {
        tokens.push({ type: TokenType.Operator, value: ">=" });
        i += 2;
        continue;
      }
      tokens.push({ type: TokenType.Operator, value: ">" });
      i++;
      continue;
    }
    if (ch === "=") {
      tokens.push({ type: TokenType.Operator, value: "=" });
      i++;
      continue;
    }
    if (ch === "&") {
      tokens.push({ type: TokenType.Operator, value: "&" });
      i++;
      continue;
    }
    if (ch === "^") {
      tokens.push({ type: TokenType.Operator, value: "^" });
      i++;
      continue;
    }
    if (ch === "*") {
      tokens.push({ type: TokenType.Operator, value: "*" });
      i++;
      continue;
    }
    if (ch === "/") {
      tokens.push({ type: TokenType.Operator, value: "/" });
      i++;
      continue;
    }

    // + and - : unary or binary
    if (ch === "+" || ch === "-") {
      if (!lastTokenIsValue()) {
        tokens.push({ type: TokenType.UnaryPrefix, value: ch });
      } else {
        tokens.push({ type: TokenType.Operator, value: ch });
      }
      i++;
      continue;
    }

    // @ implicit intersection prefix (Excel 365)
    if (ch === "@") {
      tokens.push({ type: TokenType.AtSign });
      i++;
      continue;
    }

    // Unknown character — skip
    i++;
  }

  return tokens;
}
