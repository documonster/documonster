/**
 * Recognising a signature dictionary whose `/Contents` must stay unencrypted.
 *
 * A signature's `/Contents` is written in the clear even in an encrypted file
 * (ISO 32000-1 §7.6.1): it is computed over the file's final bytes, after encryption.
 * The reader must therefore not decrypt it and the writer must not encrypt it, and the
 * two have to agree on which dictionaries that applies to — a dictionary one side treats
 * as a signature and the other does not turns the signature into noise. Both sides use
 * the rule here: the dictionary has a `/Contents` string and a direct `/ByteRange` array
 * of four numbers.
 *
 * The reader works on parsed objects ({@link isSignatureDict}); the writer works on
 * serialised text ({@link findSignatureContents}), which applies the same rule.
 */

import { isPdfArray } from "@pdf/reader/pdf-parser";
import type { PdfDictValue } from "@pdf/reader/pdf-parser";

/** Whether a parsed dictionary is a signature dictionary (see the module comment). */
export function isSignatureDict(dict: PdfDictValue): boolean {
  const range = dict.get("ByteRange");
  return (
    dict.get("Contents") instanceof Uint8Array &&
    isPdfArray(range) &&
    range.length === 4 &&
    range.every(n => typeof n === "number")
  );
}

interface DictFrame {
  kind: "dict";
  expectKey: boolean;
  key: string | null;
  contentsAt: number;
  byteRangeOk: boolean;
}

interface ArrayFrame {
  kind: "array";
  parentKey: string | null;
  numbers: number;
  other: boolean;
}

const DELIMITERS = "()<>[]{}/%";

function isRegular(ch: string): boolean {
  return !DELIMITERS.includes(ch) && !/\s/.test(ch);
}

/**
 * The offsets in serialised object text at which the `/Contents` string of a signature
 * dictionary starts — the `(` or `<` that opens it. Same rule as {@link isSignatureDict},
 * applied to text: a dictionary with a string `/Contents` and a direct `/ByteRange` array
 * of exactly four numbers.
 */
export function findSignatureContents(content: string): Set<number> {
  const found = new Set<number>();
  const stack: Array<DictFrame | ArrayFrame> = [];
  const top = (): DictFrame | ArrayFrame | undefined => stack[stack.length - 1];

  // A complete value has been read in the current container.
  const valueDone = (): void => {
    const frame = top();
    if (frame?.kind === "dict") {
      frame.expectKey = true;
      frame.key = null;
    } else if (frame) {
      frame.other = true;
    }
  };

  let i = 0;
  while (i < content.length) {
    const ch = content[i];
    if (ch === "(") {
      const start = i;
      let depth = 0;
      while (i < content.length) {
        const c = content[i];
        if (c === "\\") {
          i += 2;
          continue;
        }
        i++;
        if (c === "(") {
          depth++;
        } else if (c === ")" && --depth === 0) {
          break;
        }
      }
      onString(start);
    } else if (ch === "<" && content[i + 1] === "<") {
      const frame = top();
      if (frame?.kind === "array") {
        frame.other = true;
      }
      stack.push({ kind: "dict", expectKey: true, key: null, contentsAt: -1, byteRangeOk: false });
      i += 2;
    } else if (ch === ">" && content[i + 1] === ">") {
      const frame = stack.pop();
      if (frame?.kind === "dict" && frame.byteRangeOk && frame.contentsAt >= 0) {
        found.add(frame.contentsAt);
      }
      valueDone();
      i += 2;
    } else if (ch === "<") {
      const start = i;
      const end = content.indexOf(">", i);
      i = end < 0 ? content.length : end + 1;
      onString(start);
    } else if (ch === "[") {
      const frame = top();
      const parentKey = frame?.kind === "dict" && !frame.expectKey ? frame.key : null;
      if (frame?.kind === "array") {
        frame.other = true;
      }
      stack.push({ kind: "array", parentKey, numbers: 0, other: false });
      i++;
    } else if (ch === "]") {
      const frame = stack.pop();
      const parent = top();
      if (frame?.kind === "array" && parent?.kind === "dict" && frame.parentKey === "ByteRange") {
        parent.byteRangeOk = !frame.other && frame.numbers === 4;
      }
      if (parent?.kind === "dict") {
        valueDone();
      }
      i++;
    } else if (ch === "%") {
      while (i < content.length && content[i] !== "\n" && content[i] !== "\r") {
        i++;
      }
    } else if (ch === "/") {
      let j = i + 1;
      while (j < content.length && isRegular(content[j])) {
        j++;
      }
      const name = content.slice(i + 1, j);
      i = j;
      const frame = top();
      if (frame?.kind === "dict" && frame.expectKey) {
        frame.expectKey = false;
        frame.key = name;
      } else {
        valueDone();
      }
    } else if (isRegular(ch)) {
      let j = i + 1;
      while (j < content.length && isRegular(content[j])) {
        j++;
      }
      const token = content.slice(i, j);
      i = j;
      const frame = top();
      if (frame?.kind === "array") {
        if (token === "R") {
          // `n g R`: the two numbers before were a reference, not two numbers.
          frame.numbers -= 2;
          frame.other = true;
        } else if (Number.isFinite(Number(token))) {
          frame.numbers++;
        } else {
          frame.other = true;
        }
      } else if (frame?.kind === "dict" && !frame.expectKey) {
        // A reference's generation and `R` arrive while a key is expected, and are skipped.
        valueDone();
      }
    } else {
      i++;
    }
  }
  return found;

  function onString(start: number): void {
    const frame = top();
    if (frame?.kind === "dict" && !frame.expectKey && frame.key === "Contents") {
      frame.contentsAt = start;
    }
    valueDone();
  }
}
