/**
 * UTF-8 decoding for a byte-fed parser, in chunks, with every byte accounted for.
 *
 * `TextDecoder` replaces a malformed sequence with one U+FFFD however many bytes it spans, which
 * loses the byte positions the parser reports and makes the decoded text depend on where a chunk
 * happened to end. So a malformed byte is decoded instead to the lone low surrogate
 * `U+DC00 + byte` — never produced by well-formed UTF-8, and one UTF-16 unit for one byte. The
 * parser treats a lone surrogate as an invalid XML character, so the document's
 * `invalidCharHandling` decides what becomes of it, and every unit maps back to its bytes.
 *
 * Well-formed input is decoded by the platform's `TextDecoder`; only a chunk it refuses is decoded
 * here, byte by byte.
 */

/** The lone surrogate a malformed byte decodes to; {@link escapedByte} reverses it. */
const ESCAPE_BASE = 0xdc00;

/** The byte a lone low surrogate from {@link Utf8ChunkDecoder} stands for, or `-1`. */
export function escapedByte(code: number): number {
  return code >= 0xdc80 && code <= 0xdcff ? code - ESCAPE_BASE : -1;
}

/**
 * How many bytes of `bytes` end on a character boundary: a sequence whose lead byte is in its last
 * three bytes and that is not complete is held back for the next chunk.
 */
function completeLength(bytes: Uint8Array): number {
  const n = bytes.length;
  const stop = n > 4 ? n - 4 : 0;
  let j = n - 1;
  while (j >= stop && (bytes[j] & 0xc0) === 0x80) {
    j--;
  }
  if (j < stop) {
    return n;
  }
  const lead = bytes[j];
  const need =
    lead >= 0xc2 && lead <= 0xdf
      ? 2
      : lead >= 0xe0 && lead <= 0xef
        ? 3
        : lead >= 0xf0 && lead <= 0xf4
          ? 4
          : 1;
  if (n - j >= need) {
    return n;
  }
  // Held back only while it can still become a character: `E0 80` cannot — after `E0` the next byte
  // is at least `A0` — and is malformed already, to be decoded (as escapes) now rather than taken at
  // the end of the input for a character cut short.
  if (j + 1 < n) {
    const second = bytes[j + 1];
    const low = lead === 0xe0 ? 0xa0 : lead === 0xf0 ? 0x90 : 0x80;
    const high = lead === 0xed ? 0x9f : lead === 0xf4 ? 0x8f : 0xbf;
    if (second < low || second > high) {
      return n;
    }
  }
  return j;
}

/** Decode `bytes`, each byte of a malformed sequence becoming its escape. */
function decodeEscaped(bytes: Uint8Array): string {
  const parts: string[] = [];
  const units: number[] = [];
  const n = bytes.length;
  let i = 0;
  while (i < n) {
    const c = bytes[i];
    if (c < 0x80) {
      units.push(c);
      i++;
    } else {
      let need = 0;
      let lo = 0x80;
      let hi = 0xbf;
      let cp = 0;
      if (c >= 0xc2 && c <= 0xdf) {
        need = 1;
        cp = c & 0x1f;
      } else if (c >= 0xe0 && c <= 0xef) {
        need = 2;
        cp = c & 0x0f;
        lo = c === 0xe0 ? 0xa0 : 0x80;
        hi = c === 0xed ? 0x9f : 0xbf;
      } else if (c >= 0xf0 && c <= 0xf4) {
        need = 3;
        cp = c & 0x07;
        lo = c === 0xf0 ? 0x90 : 0x80;
        hi = c === 0xf4 ? 0x8f : 0xbf;
      }
      let ok = need > 0;
      for (let k = 1; ok && k <= need; k++) {
        const d = i + k < n ? bytes[i + k] : -1;
        if (d < (k === 1 ? lo : 0x80) || d > (k === 1 ? hi : 0xbf)) {
          ok = false;
        } else {
          cp = (cp << 6) | (d & 0x3f);
        }
      }
      if (!ok) {
        units.push(ESCAPE_BASE + c);
        i++;
      } else if (cp >= 0x10000) {
        cp -= 0x10000;
        units.push(0xd800 + (cp >> 10), 0xdc00 + (cp & 0x3ff));
        i += need + 1;
      } else {
        units.push(cp);
        i += need + 1;
      }
    }
    if (units.length >= 8192) {
      parts.push(String.fromCharCode(...units));
      units.length = 0;
    }
  }
  parts.push(String.fromCharCode(...units));
  return parts.join("");
}

/**
 * Decodes a file's bytes chunk by chunk. Each call returns the text of the complete characters
 * received so far; {@link textStart} is where that text begins in the file, and
 * {@link oneBytePerUnit} says whether every UTF-16 unit of it is one byte — so a position in the
 * text is a byte offset by addition alone.
 */
export class Utf8ChunkDecoder {
  /** Offset in the file of the first byte of the text the last {@link decode} returned. */
  textStart = 0;
  /** Offset in the file just past the last byte of the text the last {@link decode} returned. */
  textEnd = 0;
  /** Whether the input ended inside a character: the last call was `final`, with a character not whole. */
  endedInsideCharacter = false;
  /** Whether each UTF-16 unit of the text the last {@link decode} returned stands for one byte. */
  oneBytePerUnit = true;
  /** Offset of the next byte to arrive. */
  next: number;
  /** The start of a character that has not arrived whole. */
  private pending: Uint8Array | undefined;

  /** Whether the start of a character that has not arrived whole is held for the next chunk. */
  get holdsPartialCharacter(): boolean {
    return this.pending !== undefined;
  }
  /**
   * Whether the document declared itself ASCII: from then on a byte past `0x7F` is not a character
   * of its encoding (XML 1.0 §4.3.3), and decodes to its escape as a malformed byte does.
   */
  private ascii = false;
  /** The bytes the text the last {@link decode} returned came from, from {@link textStart}. */
  private lastBytes: Uint8Array = new Uint8Array(0);
  /** Continuation bytes still to drop: a start inside a character has the rest of it first. */
  private skip: number;
  private readonly decoder = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true });

  /**
   * @param byteOffset - Offset in the file of the first byte.
   * @param resume - Whether the first byte may fall inside a character, whose remaining
   *   continuation bytes are then dropped.
   */
  /**
   * @param ascii - Whether the document is known to be ASCII from the first byte — a reader that
   *   starts inside a file whose declaration said so (see `XmlPrologContext.encoding`).
   */
  constructor(byteOffset: number, resume: boolean, ascii = false) {
    this.next = byteOffset;
    this.ascii = ascii;
    // An ASCII document has no character to start inside: every byte is one.
    this.skip = resume && !ascii ? 3 : 0;
  }

  /**
   * Decode the next chunk. With `final` the input ends: a character still not whole is not decoded,
   * and {@link endedInsideCharacter} says so.
   */
  decode(chunk: Uint8Array, final: boolean): string {
    let bytes = chunk;
    let start = this.next;
    this.next += chunk.length;
    if (this.skip > 0) {
      let k = 0;
      while (k < bytes.length && this.skip > 0 && (bytes[k] & 0xc0) === 0x80) {
        k++;
        this.skip--;
      }
      if (k < bytes.length) {
        this.skip = 0;
      }
      if (k > 0) {
        bytes = bytes.subarray(k);
        start += k;
      }
    }
    const pending = this.pending;
    if (pending !== undefined) {
      const merged = new Uint8Array(pending.length + bytes.length);
      merged.set(pending);
      merged.set(bytes, pending.length);
      bytes = merged;
      start -= pending.length;
      this.pending = undefined;
    }
    if (this.ascii) {
      this.textStart = start;
      this.textEnd = start + bytes.length;
      this.oneBytePerUnit = true;
      this.lastBytes = bytes;
      return decodeAscii(bytes);
    }
    const end = completeLength(bytes);
    if (end < bytes.length) {
      if (final) {
        this.endedInsideCharacter = true;
      } else {
        this.pending = bytes.slice(end);
      }
    }
    this.textStart = start;
    this.textEnd = start + end;
    if (end === 0) {
      this.oneBytePerUnit = true;
      return "";
    }
    const view = end === bytes.length ? bytes : bytes.subarray(0, end);
    this.lastBytes = view;
    let text: string;
    try {
      text = this.decoder.decode(view);
    } catch {
      text = decodeEscaped(view);
    }
    this.oneBytePerUnit = text.length === end;
    return text;
  }

  /**
   * The document declared itself ASCII: decode the rest of the last chunk again, from byte `from`,
   * and every later byte, as ASCII — the start of a character held for the next chunk included,
   * which is no ASCII either. Sets {@link textStart} and {@link textEnd} to the text returned.
   */
  switchToAscii(from: number): string {
    this.ascii = true;
    let rest = this.lastBytes.subarray(from - this.textStart);
    const pending = this.pending;
    if (pending !== undefined) {
      const merged = new Uint8Array(rest.length + pending.length);
      merged.set(rest);
      merged.set(pending, rest.length);
      rest = merged;
      this.pending = undefined;
    }
    this.textStart = from;
    this.textEnd = from + rest.length;
    this.oneBytePerUnit = true;
    this.lastBytes = rest;
    return decodeAscii(rest);
  }
}

/** Decode `bytes` as ASCII, each byte past `0x7F` becoming its escape. */
function decodeAscii(bytes: Uint8Array): string {
  let i = 0;
  while (i < bytes.length && bytes[i] < 0x80) {
    i++;
  }
  if (i === bytes.length) {
    return decodeLatin(bytes);
  }
  const units: number[] = [];
  const parts: string[] = [];
  for (const byte of bytes) {
    units.push(byte < 0x80 ? byte : ESCAPE_BASE + byte);
    if (units.length >= 8192) {
      parts.push(String.fromCharCode(...units));
      units.length = 0;
    }
  }
  parts.push(String.fromCharCode(...units));
  return parts.join("");
}

const latin = new TextDecoder("latin1");

/** Bytes all below `0x80`, as text. */
function decodeLatin(bytes: Uint8Array): string {
  return latin.decode(bytes);
}

/**
 * The number of bytes the UTF-16 units of `text` from `from` to `to` take in UTF-8, as
 * {@link Utf8ChunkDecoder} produced them: a pair is four bytes, an escape one.
 */
export function utf8Length(text: string, from: number, to: number): number {
  let bytes = 0;
  let i = from;
  while (i < to) {
    const c = text.charCodeAt(i);
    if (c < 0x80) {
      bytes++;
      i++;
    } else if (c < 0x800) {
      bytes += 2;
      i++;
    } else if (c >= 0xd800 && c <= 0xdbff) {
      bytes += 4;
      i += 2;
    } else if (c >= 0xdc00 && c <= 0xdfff) {
      bytes++;
      i++;
    } else {
      bytes += 3;
      i++;
    }
  }
  return bytes;
}
