/**
 * The byte-fed parser's UTF-8 decoder: whatever the bytes and however they are split, the text is
 * the same, every UTF-16 unit maps back to its bytes, and well-formed input reads as `TextDecoder`
 * reads it.
 */

import { escapedByte, Utf8ChunkDecoder, utf8Length } from "@xml/utf8-decoder";
import { describe, expect, it } from "vitest";

import { random } from "./helpers/generated-documents";

const encoder = new TextEncoder();

/** Decode `bytes` split at `cuts`; the text, and whether it ended inside a character. */
function decode(bytes: Uint8Array, cuts: readonly number[]): { text: string; cut: boolean } {
  const decoder = new Utf8ChunkDecoder(0, false);
  let text = "";
  let at = 0;
  for (const cut of [...cuts, bytes.length]) {
    text += decoder.decode(bytes.subarray(at, cut), false);
    at = cut;
  }
  text += decoder.decode(new Uint8Array(0), true);
  return { text, cut: decoder.endedInsideCharacter };
}

/** Bytes that are mostly UTF-8, with malformed sequences mixed in. */
function bytesOf(seed: number): Uint8Array {
  const next = random(seed);
  const out: number[] = [];
  const pieces = ["a", "é", "€", "😀", "日", "<", "\r\n"];
  const length = 1 + Math.floor(next() * 40);
  for (let i = 0; i < length; i++) {
    const roll = next();
    if (roll < 0.7) {
      out.push(...encoder.encode(pieces[Math.floor(next() * pieces.length)]));
    } else if (roll < 0.85) {
      out.push(0x80 + Math.floor(next() * 0x80)); // a stray byte ≥ 0x80
    } else {
      // A sequence cut short or overlong.
      out.push(
        ...[
          [0xe2, 0x82],
          [0xf0, 0x9f, 0x98],
          [0xc0, 0xaf],
          [0xed, 0xa0, 0x80],
          [0xf4, 0x90, 0x80, 0x80]
        ][Math.floor(next() * 5)]
      );
    }
  }
  return new Uint8Array(out);
}

describe("Utf8ChunkDecoder", () => {
  it("gives the same text however the bytes are split, and accounts for every byte", () => {
    for (let seed = 1; seed <= 2000; seed++) {
      const bytes = bytesOf(seed);
      const whole = decode(bytes, []);
      const next = random(seed * 31);
      for (let trial = 0; trial < 5; trial++) {
        const cuts = [
          ...new Set(Array.from({ length: 4 }, () => Math.floor(next() * bytes.length)))
        ].sort((a, b) => a - b);
        expect(decode(bytes, cuts), `seed ${seed}, cuts ${cuts.join(",")}`).toEqual(whole);
      }
      // Every unit stands for its bytes: a character by its UTF-8 length, an escape by one byte.
      const held = whole.cut
        ? bytes.length - [...bytes].reverse().findIndex(b => (b & 0xc0) !== 0x80) - 1
        : bytes.length;
      expect(utf8Length(whole.text, 0, whole.text.length), `seed ${seed}`).toBe(held);
    }
  });

  it("reads well-formed UTF-8 as TextDecoder does", () => {
    const text = "a é € 😀 日本 \u{10FFFF} \uFEFF";
    const bytes = encoder.encode(text);
    for (let cut = 0; cut <= bytes.length; cut++) {
      expect(decode(bytes, [cut]).text).toBe(text);
    }
  });

  it("decodes each malformed byte to its own escape", () => {
    const { text } = decode(new Uint8Array([0x41, 0xc0, 0xaf, 0xe2, 0x28, 0xa1, 0xff]), []);
    expect([...text].map(ch => escapedByte(ch.charCodeAt(0)))).toEqual([
      -1, 0xc0, 0xaf, 0xe2, -1, 0xa1, 0xff
    ]);
  });

  it("says when the input ends inside a character", () => {
    expect(decode(new Uint8Array([0x41, 0xe2, 0x82]), [2])).toEqual({ text: "A", cut: true });
  });

  it("drops the rest of a character it starts inside, when resuming", () => {
    const bytes = encoder.encode("€a").subarray(1);
    const decoder = new Utf8ChunkDecoder(1, true);
    expect(
      decoder.decode(bytes.subarray(0, 1), false) + decoder.decode(bytes.subarray(1), false)
    ).toBe("a");
    expect(decoder.textStart).toBe(3);
  });
});
