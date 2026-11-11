/**
 * The DOCTYPE declaration against expat and libxml2, and `Xml.readRoot`.
 */

import {
  XmlArgumentError,
  XmlChunkAmbiguityError,
  XmlLimitExceededError,
  XmlParseError,
  XmlTruncatedError,
  Xml
} from "@xml/index";
import type { XmlReadRecordsResult } from "@xml/index";
import { SaxParser } from "@xml/sax";
import { describe, expect, it } from "vitest";

import { DOCTYPE_CORPUS } from "./helpers/doctype-corpus";

const encoder = new TextEncoder();

function chunked(bytes: Uint8Array, size: number): Uint8Array[] {
  const out: Uint8Array[] = [];
  for (let at = 0; at < bytes.length; at += size) {
    out.push(bytes.subarray(at, at + size));
  }
  return out;
}

describe("DOCTYPE declarations agree with expat and libxml2", () => {
  it.each(DOCTYPE_CORPUS)("%s", (_, xml, wellFormed) => {
    for (const bytes of [false, true]) {
      const errors: string[] = [];
      const opened: string[] = [];
      const parser = new SaxParser({ internalEntities: true });
      parser.on("error", err => errors.push(err.message));
      parser.on("opentag", tag => opened.push(tag.name));
      if (bytes) {
        for (const chunk of chunked(encoder.encode(xml), 1)) {
          parser.write(chunk);
        }
      } else {
        parser.write(xml);
      }
      parser.close();
      if (wellFormed) {
        expect(errors).toEqual([]);
        // Nothing inside the DOCTYPE is read as an element.
        expect(opened).toEqual(["R"]);
      } else {
        expect(errors.length).toBeGreaterThan(0);
      }
    }
  });

  it.each([
    [
      '<!DOCTYPE R [ <!ENTITY e "x&amp;y"> ]><R a="&e;">&e;</R>',
      { R: { "@_a": "x&y", "#text": "x&y" } }
    ],
    ['<!DOCTYPE R [ <!ENTITY a "A"> <!ENTITY b "[&a;]"> ]><R>&b;</R>', { R: "[A]" }],
    ['<!DOCTYPE R [ <!ENTITY e "&#38;amp;"> ]><R>&e;</R>', { R: "&" }],
    [
      '<!DOCTYPE R [ <!ENTITY e "a&#10;b"> ]><R a="&e;">&e;</R>',
      { R: { "@_a": "a b", "#text": "a\nb" } }
    ]
  ])("expands %s as expat does", (xml, expected) => {
    expect(Xml.parseToObject(xml, { internalEntities: true })).toEqual(expected);
  });
});

describe("Xml.readRoot", () => {
  const xml =
    '\uFEFF<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE DSR [ <!-- ] --> <!ENTITY v "1.2"> ]>\n<DSR xmlns="urn:ddex" version="&v;" t=\'a&amp;b\'><Body>never read <<<</Body>';
  const bytes = encoder.encode(xml);

  it("reads the root's name, attributes and the prolog's context", async () => {
    const root = await Xml.readRoot(bytes);
    expect(root.name).toBe("DSR");
    expect({ ...root.attributes }).toEqual({ xmlns: "urn:ddex", version: "1.2", t: "a&b" });
    expect(JSON.parse(JSON.stringify(root.context))).toEqual({
      root: "DSR",
      entities: { v: "1.2" }
    });
  });

  it("gives the same answer however the source is chunked", async () => {
    const expected = await Xml.readRoot(bytes);
    for (const size of [1, 2, 7, 64, 65536]) {
      expect(await Xml.readRoot(chunked(bytes, size))).toEqual(expected);
    }
  });

  it("parses nothing after the root's >, and closes the source", async () => {
    let pulled = 0;
    let closed = false;
    const rootEnd =
      bytes.indexOf(0x3e, encoder.encode(xml.slice(0, xml.indexOf("<DSR"))).length) + 1;
    function* source(): Generator<Uint8Array> {
      try {
        for (const chunk of chunked(bytes, 16)) {
          pulled += chunk.length;
          yield chunk;
        }
      } finally {
        closed = true;
      }
    }
    await Xml.readRoot(source());
    expect(closed).toBe(true);
    // Read no further than the chunk the root's `>` arrived in.
    expect(pulled).toBe(Math.ceil(rootEnd / 16) * 16);
  });

  it.each(DOCTYPE_CORPUS.filter(([, , ok]) => ok))("reads the root after %s", async (_, doc) => {
    const root = await Xml.readRoot(encoder.encode(doc));
    expect(root.name).toBe("R");
  });

  it("throws a parse error before the root, with its byte offset", async () => {
    const err = await Xml.readRoot(encoder.encode("<!-- ok --><!DOCTYPE R [ junk ]><R/>")).catch(
      e => e
    );
    expect(err).toBeInstanceOf(XmlParseError);
    expect(err.byte).toBe(25);
  });

  it("refuses a repeated attribute", async () => {
    await expect(Xml.readRoot(encoder.encode('<R a="1" a="2"/>'))).rejects.toThrow(
      /duplicate attribute: a/
    );
  });

  it.each([
    ["", XmlTruncatedError],
    ["<!-- only a comment -->", XmlTruncatedError],
    ['<R a="1', XmlTruncatedError]
  ])("throws when the input %j has no whole root start tag", async (doc, type) => {
    await expect(Xml.readRoot(encoder.encode(doc))).rejects.toBeInstanceOf(type);
  });

  it("applies the limits to byte positions, whatever the chunking", async () => {
    const doc = encoder.encode(`<!--${"x".repeat(100)}--><Root attribute="${"v".repeat(50)}"/>`);
    const start = 107;
    const tag = doc.length - start;
    const verdict = async (options: object, size: number): Promise<string> =>
      Xml.readRoot(chunked(doc, size), options).then(
        root => root.name,
        (err: Error) => `${err.name}: ${err.message}`
      );
    for (const options of [
      { maxStartByte: start + 1 },
      { maxStartByte: start },
      { maxStartByte: start + 1, maxTagBytes: tag },
      { maxStartByte: start + 1, maxTagBytes: tag - 1 }
    ]) {
      const answers = new Set<string>();
      for (const size of [1, 3, 7, 64, 4096]) {
        answers.add(await verdict(options, size));
      }
      expect(answers.size, JSON.stringify(options)).toBe(1);
    }
    expect(await verdict({ maxStartByte: start + 1, maxTagBytes: tag }, 7)).toBe("Root");
    const tooLate = await Xml.readRoot(doc, { maxStartByte: start }).catch(e => e);
    expect(tooLate).toBeInstanceOf(XmlLimitExceededError);
    expect(tooLate.limit).toBe("rootTag");
    const tooLong = await Xml.readRoot(doc, { maxTagBytes: tag - 1 }).catch(e => e);
    expect(tooLong.message).toMatch(/longer than maxTagBytes|more than maxTagBytes/);
  });

  it("refuses arguments it cannot use", async () => {
    await expect(Xml.readRoot("<R/>" as never)).rejects.toBeInstanceOf(XmlArgumentError);
    await expect(Xml.readRoot(bytes, { maxStartByte: 0 })).rejects.toBeInstanceOf(XmlArgumentError);
  });
});

describe("attribute-list declarations a parser reads (XML 1.0 §5.1)", () => {
  // Each with the attributes expat reports.
  it.each<[string, Record<string, string>]>([
    [
      '<!DOCTYPE R [ <!ENTITY e "E&#9;x"> <!ATTLIST R a CDATA "v &e;" t NMTOKENS #IMPLIED k (x|y) "  y "> <!ATTLIST R a CDATA "ignored"> ]><R t="  a   b  "/>',
      { t: "a b", a: "v E x", k: "y" }
    ],
    [
      '<!DOCTYPE R [ <!ATTLIST R a CDATA "x&#10;y" i ID #IMPLIED> ]><R i=" id1 "/>',
      { i: "id1", a: "x\ny" }
    ],
    ['<!DOCTYPE R [ <!ATTLIST R a CDATA #FIXED "f"> ]><R/>', { a: "f" }]
  ])("gives defaults and normalizes values as expat does: %s", (xml, attributes) => {
    expect({ ...Xml.parse(xml, { internalEntities: true }).root.attributes }).toEqual(attributes);
    const opened: Record<string, string>[] = [];
    const parser = new SaxParser({ internalEntities: true });
    parser.on("opentag", tag => opened.push({ ...tag.attributes }));
    for (const chunk of chunked(encoder.encode(xml), 1)) {
      parser.write(chunk);
    }
    parser.close();
    expect(opened).toEqual([attributes]);
  });

  it("gives no default that refers to an entity declared only where it does not read", () => {
    // expat gives `a` as "" with the entity skipped; it is left out rather than given a value it does not have.
    const xml = '<!DOCTYPE R SYSTEM "r.dtd" [ <!ATTLIST R a CDATA "&missing;" b CDATA "x"> ]><R/>';
    expect({ ...Xml.parse(xml).root.attributes }).toEqual({ b: "x" });
  });

  it("uses the declarations without internalEntities, and says when a default needs an entity", () => {
    const xml = '<!DOCTYPE R [ <!ENTITY e "E"> <!ATTLIST R a CDATA "&e;" t NMTOKEN #IMPLIED> ]>';
    expect({ ...Xml.parse(`${xml}<R a="given" t=" x "/>`).root.attributes }).toEqual({
      a: "given",
      t: "x"
    });
    expect(() => Xml.parse(`${xml}<R/>`)).toThrow(
      /entities declared in a DOCTYPE internal subset are not supported/
    );
  });

  it("declares namespaces through a default", () => {
    const root = Xml.parse('<!DOCTYPE p:r [ <!ATTLIST p:r xmlns:p CDATA "urn:p"> ]><p:r/>', {
      xmlns: true
    }).root;
    expect(root.uri).toBe("urn:p");
  });

  it("takes no declaration after an unread parameter entity, unless the document is standalone", () => {
    const subset = '<!ENTITY % p SYSTEM "x"> %p; <!ATTLIST R a CDATA "d">';
    expect({ ...Xml.parse(`<!DOCTYPE R [ ${subset} ]><R/>`).root.attributes }).toEqual({});
    expect({
      ...Xml.parse(`<?xml version="1.0" standalone="yes"?><!DOCTYPE R [ ${subset} ]><R/>`).root
        .attributes
    }).toEqual({ a: "d" });
  });

  it("carries them in the prolog context, so a range reads its records as the whole file does", async () => {
    let xml = '<!DOCTYPE doc [ <!ATTLIST r kind CDATA "plain" ref IDREFS #IMPLIED> ]>\n<doc>\n';
    for (let i = 0; i < 20; i++) {
      xml += i % 2 ? `  <r ref="  a${i}   b  "/>\n` : `  <r kind="k${i}"><n>${i}</n></r>\n`;
    }
    xml += "</doc>\n";
    const bytes = encoder.encode(xml);
    const root = await Xml.readRoot(bytes);
    expect(root.context.attributes).toEqual({
      r: { kind: { default: "plain" }, ref: { tokenized: true } }
    });
    // Sent to a worker and back.
    const context = JSON.parse(JSON.stringify(root.context)) as typeof root.context;
    const whole: unknown[] = [];
    await Xml.readRecords(bytes, {
      tags: ["r"],
      onRecord: record => void whole.push(record.object)
    });
    expect(whole[1]).toEqual({ r: { "@_ref": "a1 b", "@_kind": "plain" } });
    let agreed = 0;
    for (const size of [17, 64, 200]) {
      const parts: unknown[] = [];
      const results: XmlReadRecordsResult[] = [];
      let ambiguous = false;
      for (let offset = 0; offset < bytes.length; offset += size) {
        try {
          results.push(
            await Xml.readRecords(bytes.subarray(offset), {
              tags: ["r"],
              offset,
              size,
              context,
              onRecord: record => void parts.push(record.object)
            })
          );
        } catch (err) {
          // A cut through markup may be refused; nothing else may fail.
          if (!(err instanceof XmlChunkAmbiguityError)) {
            throw err;
          }
          ambiguous = true;
        }
      }
      if (!ambiguous) {
        Xml.verifyPartition(JSON.parse(JSON.stringify(results)));
        expect(parts, `size ${size}`).toEqual(whole);
        agreed++;
      }
    }
    expect(agreed).toBeGreaterThan(0);
  });

  it("refuses a context whose attribute declarations are not declarations", async () => {
    for (const attributes of [
      { r: {} },
      { r: { a: {} } },
      { r: { a: { tokenized: false } } },
      { "1r": { a: { default: "x" } } }
    ]) {
      await expect(
        Xml.readRecords(encoder.encode("<r/>"), {
          tags: ["r"],
          offset: 1,
          context: { root: "doc", entities: {}, attributes } as never,
          onRecord: () => {}
        }),
        JSON.stringify(attributes)
      ).rejects.toThrow(/attribute declarations/);
    }
  });
});
