import { SaxParser } from "@xml/sax";
import type { SaxOptions } from "@xml/types";
import { describe, expect, it } from "vitest";

/**
 * `scanAttributes` reads a start tag's attributes in one pass when the whole attribute sits in one chunk, and hands
 * anything else to the state machine. Feeding the same input one character per chunk never satisfies that, so every
 * attribute then goes through the state machine alone — which makes it the oracle: the two runs must agree on every
 * event, every error and every reported position.
 */

function record(xml: string, chunkSize: number, options: SaxOptions): unknown[] {
  const events: unknown[] = [];
  const parser = new SaxParser(options);
  parser.on("opentag", tag =>
    events.push(["open", tag.name, { ...tag.attributes }, tag.isSelfClosing, tag.uri, tag.local])
  );
  parser.on("closetag", tag => events.push(["close", tag.name]));
  parser.on("text", text => events.push(["text", text]));
  parser.on("error", error => events.push(["error", error.message]));
  for (let i = 0; i < xml.length; i += chunkSize) {
    parser.write(xml.slice(i, i + chunkSize));
  }
  parser.close();
  return events;
}

const PIECES = [
  ' a="1"',
  " b='two'",
  ' r="A1"',
  ' s="12"',
  ' t="inlineStr"',
  ' x:y="ns"',
  ' e="a&amp;b"',
  ' gt="1>0"',
  ' sp="  spaced  "',
  ' u="日本"',
  ' tab="a\tb"',
  ' nl="a\nb"',
  '\n  c="after-newline"',
  '\tt2="tab-before"',
  ' a="dup"',
  " novalue",
  " noquote=1",
  ' lt="a<b"',
  ' adj="1"z="2"',
  " empty=''",
  ' q="it\'s"',
  ' ws = "around-equals"',
  " / ",
  ' sl="1"/ ',
  ' cr="a\r\nb"',
  ' ent="&#65;&lt;"',
  ' ctl="a\u0001b"',
  " =",
  ' "orphan"',
  ' é="accent-name"',
  ' x="\uD83D\uDE00"'
];

/** A deterministic generator, so a failure names a seed rather than a flake. */
function random(seed: number): () => number {
  let state = seed >>> 0 || 1;
  return () => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return (state >>> 0) / 0x100000000;
  };
}

function document(seed: number): string {
  const next = random(seed);
  const elements: string[] = [];
  const count = 1 + Math.floor(next() * 6);
  for (let i = 0; i < count; i++) {
    let attributes = "";
    const n = Math.floor(next() * 5);
    for (let j = 0; j < n; j++) {
      attributes += PIECES[Math.floor(next() * PIECES.length)];
    }
    const name = ["c", "row", "v", "x:el", "a.b-c"][Math.floor(next() * 5)];
    const space = next() < 0.3 ? " " : "";
    elements.push(
      next() < 0.4
        ? `<${name}${attributes}${space}/>`
        : `<${name}${attributes}${space}>t${i}</${name}>`
    );
  }
  return `<root xmlns:x="urn:x">${elements.join("")}</root>`;
}

describe("SaxParser attribute scanning", () => {
  for (const options of [
    {},
    { position: false },
    { xmlns: true },
    { invalidCharHandling: "skip" as const }
  ] satisfies SaxOptions[]) {
    it(`agrees with the state machine on 500 generated documents (${JSON.stringify(options)})`, () => {
      for (let seed = 1; seed <= 500; seed++) {
        const xml = document(seed);
        const whole = record(xml, xml.length, options);
        expect(record(xml, 1, options), `seed ${seed}: ${xml}`).toEqual(whole);
        // A chunk boundary falling inside an attribute is the case the fallback exists for.
        expect(record(xml, 7, options), `seed ${seed}: ${xml}`).toEqual(whole);
        const size = 2 + (seed % 13);
        expect(record(xml, size, options), `seed ${seed} / ${size}: ${xml}`).toEqual(whole);
      }
    });
  }

  it("reads the common worksheet shape", () => {
    expect(record('<c r="A1" s="3" t="s"><v>0</v></c>', 1024, {})).toEqual([
      ["open", "c", { r: "A1", s: "3", t: "s" }, false, undefined, undefined],
      ["open", "v", {}, false, undefined, undefined],
      ["text", "0"],
      ["close", "v"],
      ["close", "c"]
    ]);
  });
});
