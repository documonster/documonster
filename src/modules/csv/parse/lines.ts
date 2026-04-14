/**
 * CSV Line Splitting Utilities
 *
 * Shared helpers for splitting CSV input into lines while preserving
 * the original line ending length (LF/CR/CRLF).
 */

export interface SplitLine {
  /** Line content without the line ending */
  line: string;
  /** Line ending length (0 when last line has no newline) */
  lineEndingLength: number;
  /** Total length of line + line ending */
  lineLengthWithEnding: number;
}

/**
 * Cache for global-flag versions of RegExp objects.
 * Avoids re-creating `new RegExp(..., 'g')` on every call to splitLinesWithEndings.
 */
const globalRegexCache = new WeakMap<RegExp, RegExp>();

function getCachedGlobalRegex(re: RegExp): RegExp {
  let cached = globalRegexCache.get(re);
  if (!cached) {
    cached = new RegExp(re.source, `${re.flags.replace(/g/g, "")}g`);
    globalRegexCache.set(re, cached);
  }
  return cached;
}

/**
 * Split input into lines using the given linebreak regex and yield per-line
 * metadata including the actual line ending length.
 *
 * Notes:
 * - Works with mixed line endings.
 * - Skips trailing split artifacts (empty string produced by split when the input ends with a newline).
 */
export function* splitLinesWithEndings(
  input: string,
  linebreakRegex: RegExp | string
): Generator<SplitLine, void, undefined> {
  if (input === "") {
    return;
  }

  if (typeof linebreakRegex === "string") {
    const sep = linebreakRegex;
    if (sep === "") {
      yield { line: input, lineEndingLength: 0, lineLengthWithEnding: input.length };
      return;
    }

    let pos = 0;
    while (true) {
      const idx = input.indexOf(sep, pos);
      if (idx === -1) {
        if (pos === input.length) {
          return;
        }
        const line = input.slice(pos);
        yield { line, lineEndingLength: 0, lineLengthWithEnding: line.length };
        return;
      }

      const line = input.slice(pos, idx);
      const lineEndingLength = sep.length;
      const lineLengthWithEnding = line.length + lineEndingLength;
      yield { line, lineEndingLength, lineLengthWithEnding };
      pos = idx + sep.length;
    }
  } else {
    const re = linebreakRegex.global ? linebreakRegex : getCachedGlobalRegex(linebreakRegex);

    let pos = 0;
    re.lastIndex = 0;

    while (true) {
      const match = re.exec(input);
      if (!match) {
        break;
      }

      const start = match.index;
      const end = start + match[0].length;
      const line = input.slice(pos, start);
      const lineEndingLength = match[0].length;
      const lineLengthWithEnding = line.length + lineEndingLength;
      yield { line, lineEndingLength, lineLengthWithEnding };
      pos = end;

      // Safety: avoid infinite loops for zero-length matches.
      if (match[0].length === 0) {
        re.lastIndex++;
      }
    }

    // If input ends with a line ending, don't yield a trailing empty line.
    if (pos === input.length) {
      return;
    }

    const tail = input.slice(pos);
    yield { line: tail, lineEndingLength: 0, lineLengthWithEnding: tail.length };
  }
}
