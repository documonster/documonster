/**
 * `Xml.readRecords`: one pass that builds each record's object as `Xml.parseToObject` would from
 * its text, and byte ranges of a file that, read independently, give exactly the records of the
 * whole file — or say they cannot.
 */

import {
  AbortError,
  isAbortError,
  XmlArgumentError,
  XmlChunkAmbiguityError,
  XmlLimitExceededError,
  XmlParseError,
  XmlTruncatedError,
  Xml
} from "@xml/index";
import type {
  XmlPartitionStructure,
  XmlPrologContext,
  XmlReadRecordsOptions,
  XmlReadRecordsResult,
  XmlRecord
} from "@xml/index";
import { RecordScanner } from "@xml/records";
import { ObjectBuilder } from "@xml/to-object";
import { describe, expect, it } from "vitest";

import { generateDocument, random } from "./helpers/generated-documents";
import type { GenerateOptions } from "./helpers/generated-documents";

const encoder = new TextEncoder();

/** `bytes` in chunks of `size`. */
function chunked(bytes: Uint8Array, size: number): Uint8Array[] {
  const out: Uint8Array[] = [];
  for (let at = 0; at < bytes.length; at += size) {
    out.push(bytes.subarray(at, at + size));
  }
  return out;
}

type Options = Omit<XmlReadRecordsOptions, "onRecord">;

async function read(
  bytes: Uint8Array,
  options: Options,
  sourceChunk = 65536
): Promise<{ records: XmlRecord[]; endByte: number; structure: XmlPartitionStructure }> {
  const records: XmlRecord[] = [];
  const offset = options.offset ?? 0;
  const result = await Xml.readRecords(chunked(bytes.subarray(offset), sourceChunk), {
    ...options,
    onRecord: record => {
      records.push(record);
    }
  });
  expect(result.records).toBe(records.length);
  return { records, endByte: result.endByte, structure: result.structure };
}

/** What a record is compared on: its range, its name and its object. */
function key(record: XmlRecord): string {
  return `${record.name}@${record.startByte}-${record.endByte}:${JSON.stringify(record.object)}`;
}

/**
 * Read `bytes` as ranges of `size` and check the guarantee: either every range reads without an
 * error, `Xml.verifyPartition` accepts their structures, and their records, in order, are the whole
 * file's — or some range throws `XmlChunkAmbiguityError`. Returns which.
 */
async function partition(
  bytes: Uint8Array,
  size: number,
  whole: readonly string[],
  context: XmlPrologContext,
  options: Options,
  sourceChunk: number
): Promise<"agreed" | "ambiguous"> {
  const got: string[] = [];
  const results: { structure: XmlPartitionStructure }[] = [];
  let ambiguous = false;
  for (let offset = 0; offset < bytes.length; offset += size) {
    try {
      const { records, structure } = await read(
        bytes,
        { ...options, offset, size, context },
        sourceChunk
      );
      results.push({ structure });
      for (const record of records) {
        expect(record.startByte).toBeGreaterThanOrEqual(offset);
        expect(record.startByte).toBeLessThan(offset + size);
        got.push(key(record));
      }
    } catch (err) {
      if (!(err instanceof XmlChunkAmbiguityError)) {
        throw err;
      }
      ambiguous = true;
    }
  }
  if (!ambiguous) {
    try {
      Xml.verifyPartition(results);
    } catch (err) {
      if (!(err instanceof XmlChunkAmbiguityError)) {
        throw err;
      }
      return "ambiguous";
    }
    expect(got).toEqual(whole);
    return "agreed";
  }
  return "ambiguous";
}

async function wholeOf(bytes: Uint8Array, options: Options = {} as Options) {
  const { records } = await read(bytes, { ...options, tags: ["r"] });
  const { context } = await Xml.readRoot(bytes);
  return { whole: records.map(key), context };
}

describe("Xml.readRecords — records", () => {
  it("reads each record's object, range and text", async () => {
    const xml =
      '<?xml version="1.0"?>\n<doc>\n  <r id="1"><n>One</n><x/></r>\n  <r id="2" t="a&amp;b"/>\n  <o><r id="3">ü€😀</r></o>\n</doc>\n';
    const bytes = encoder.encode(xml);
    const { records } = await read(bytes, { tags: ["r"], rawXml: true });
    expect(records.map(r => r.object)).toEqual([
      { r: { "@_id": "1", n: "One", x: "" } },
      { r: { "@_id": "2", "@_t": "a&b" } },
      { r: { "@_id": "3", "#text": "ü€😀" } }
    ]);
    for (const record of records) {
      // The range is the record's bytes, and its text is those bytes decoded.
      expect(new TextDecoder().decode(bytes.subarray(record.startByte, record.endByte))).toBe(
        record.xml
      );
      expect(record.xml!.startsWith("<r")).toBe(true);
    }
    expect(records.map(r => r.xml)).toEqual([
      '<r id="1"><n>One</n><x/></r>',
      '<r id="2" t="a&amp;b"/>',
      '<r id="3">ü€😀</r>'
    ]);
  });

  it("gives no text unless rawXml is set", async () => {
    const { records } = await read(encoder.encode("<d><r/></d>"), { tags: ["r"] });
    expect(records[0].xml).toBeUndefined();
  });

  it("reads a record inside another record as part of it, whatever its name", async () => {
    const { records } = await read(encoder.encode("<d><a><b/><a>x</a></a><b>y</b></d>"), {
      tags: ["a", "b"]
    });
    expect(records.map(r => r.object)).toEqual([{ a: { b: "", a: "x" } }, { b: "y" }]);
  });

  it("reads the root element as a record when asked", async () => {
    const { records } = await read(encoder.encode("<d><r/></d>"), { tags: ["d", "r"] });
    expect(records.map(r => r.name)).toEqual(["d"]);
  });

  it("awaits onRecord before reading on", async () => {
    const order: string[] = [];
    await Xml.readRecords([encoder.encode("<d><r>1</r><r>2</r></d>")], {
      tags: ["r"],
      onRecord: async record => {
        order.push(`start ${String(record.object.r)}`);
        await new Promise(resolve => setTimeout(resolve, 5));
        order.push(`end ${String(record.object.r)}`);
      }
    });
    expect(order).toEqual(["start 1", "end 1", "start 2", "end 2"]);
  });

  it("expands the entities the internal subset declares", async () => {
    const xml = '<!DOCTYPE d [ <!ENTITY e "ä &amp; b"> ]><d><r a="&e;">&e;</r></d>';
    const { records } = await read(encoder.encode(xml), { tags: ["r"] });
    expect(records[0].object).toEqual({ r: { "@_a": "ä & b", "#text": "ä & b" } });
  });

  it.each([
    ["error", undefined],
    ["replace", { r: "a\uFFFDb" }],
    ["skip", { r: "ab" }]
  ] as const)(
    "reads a byte that is not UTF-8 as invalidCharHandling %s says",
    async (mode, object) => {
      const bytes = new Uint8Array([
        ...encoder.encode("<d><r>a"),
        0xff,
        ...encoder.encode("b</r></d>")
      ]);
      if (object === undefined) {
        await expect(read(bytes, { tags: ["r"], invalidCharHandling: mode })).rejects.toThrow(
          /invalid UTF-8: byte 0xff .*\(at byte 7\)/
        );
      } else {
        const { records } = await read(bytes, { tags: ["r"], invalidCharHandling: mode });
        expect(records[0].object).toEqual(object);
      }
    }
  );

  it("matches Xml.parseToObject of each record's text, on generated documents", async () => {
    const conversions = [
      {},
      { alwaysArray: true },
      { attributePrefix: "", textKey: "_" },
      { ignoreAttributes: true, ignoreWhitespaceText: false },
      { isArray: (name: string) => name === "n", preserveCData: false }
    ];
    for (let seed = 1; seed <= 60; seed++) {
      const options: GenerateOptions = {
        decoys: seed % 2 === 0,
        sameNameNesting: seed % 3 === 0,
        oneLine: seed % 5 === 0,
        doctype: seed % 4 === 0
      };
      const { xml, doctype } = generateDocument(seed, options);
      const bytes = encoder.encode(xml);
      const convert = conversions[seed % conversions.length];
      const names: string[] = [];
      const isArray = (convert as { isArray?: (name: string) => boolean }).isArray;
      const { records } = await read(bytes, {
        tags: ["r"],
        rawXml: true,
        ...convert,
        ...(isArray
          ? {
              isArray: (name: string) => {
                names.push(name);
                return isArray(name);
              }
            }
          : {})
      });
      expect(records.length, xml).toBeGreaterThan(0);
      const asked: string[] = [];
      for (const record of records) {
        const expected = Xml.parseToObject(doctype + record.xml!, {
          internalEntities: true,
          ...convert,
          ...(isArray
            ? {
                isArray: (name: string) => {
                  asked.push(name);
                  return isArray(name);
                }
              }
            : {})
        });
        expect(record.object, `seed ${seed}: ${record.xml}`).toEqual(expected);
        expect(Object.getPrototypeOf(record.object)).toBeNull();
      }
      // `isArray` is asked about the same names, in the same order.
      expect(names).toEqual(asked);
    }
  });
});

describe("Xml.readRecords — chunks", () => {
  it("gives the records of the whole file across every partition of a typical file", async () => {
    let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<doc v="1">\n';
    for (let i = 0; i < 12; i++) {
      xml += `  <r id="${i}">\n    <t>Söng ${i} &amp; more</t>\n    <c/>\n  </r>\n`;
    }
    xml += "</doc>\n";
    const bytes = encoder.encode(xml);
    const { whole, context } = await wholeOf(bytes);
    // Byte offsets equal character offsets here only up to the first non-ASCII character, so the
    // root's end tag is found in the bytes.
    for (let size = 1; size <= bytes.length + 1; size++) {
      // A file with no tags in comments or CDATA and no records inside records places every boundary
      // where it can be read from — inside the root's end tag too: the range before says what the
      // `/doc>` the next one reads as text really is.
      expect(
        await partition(bytes, size, whole, context, { tags: ["r"] }, 65536),
        `size ${size}`
      ).toBe("agreed");
    }
  });

  it("holds the guarantee for every range size and source chunking, on generated documents", async () => {
    let agreed = 0;
    let ambiguous = 0;
    for (let seed = 1; seed <= 24; seed++) {
      const { xml } = generateDocument(seed, {
        decoys: seed % 2 === 0,
        sameNameNesting: seed % 3 === 0,
        oneLine: seed % 4 === 0,
        doctype: seed % 5 === 0
      });
      const bytes = encoder.encode(xml);
      const { whole, context } = await wholeOf(bytes);
      for (let size = 1; size <= bytes.length + 1; size++) {
        for (const sourceChunk of [1, 7, 65536]) {
          const outcome = await partition(
            bytes,
            size,
            whole,
            context,
            { tags: ["r"] },
            sourceChunk
          );
          if (outcome === "agreed") {
            agreed++;
          } else {
            ambiguous++;
          }
        }
      }
    }
    // Most partitions place every boundary; the rest say they cannot.
    expect(agreed).toBeGreaterThan(ambiguous);
  }, 120_000);

  it("holds the guarantee on a file that is not all UTF-8, read with invalidCharHandling replace", async () => {
    const bytes = new Uint8Array([
      ...encoder.encode("<doc>\n<r>a"),
      0xc3,
      0x28,
      0xff,
      ...encoder.encode("b</r>\n<r>"),
      0xe2,
      0x82,
      ...encoder.encode("</r>\n<r>ü</r>\n</doc>")
    ]);
    const options = { tags: ["r"], invalidCharHandling: "replace" } as const;
    const { whole, context } = await wholeOf(bytes, options);
    expect(whole).toHaveLength(3);
    for (let size = 1; size <= bytes.length + 1; size++) {
      for (const sourceChunk of [1, 7, 65536]) {
        await partition(bytes, size, whole, context, options, sourceChunk);
      }
    }
  });

  it("does not finish a check while a start tag is half read in a reference of an attribute value", async () => {
    // A range ends inside `&amp;`: the start tag around it may still be a record's.
    const xml = '<R><T>xxxxxxxx<T a="&amp;"/></T></R>';
    const bytes = encoder.encode(xml);
    const whole = (await read(bytes, { tags: ["T"] })).records.map(key);
    const context = { root: "R", entities: {} };
    for (let size = 1; size <= bytes.length + 1; size++) {
      const got: string[] = [];
      let failed = false;
      for (let offset = 0; offset < bytes.length; offset += size) {
        try {
          got.push(...(await read(bytes, { tags: ["T"], offset, size, context })).records.map(key));
        } catch {
          failed = true;
        }
      }
      if (!failed) {
        expect(got, `size ${size}`).toEqual(whole);
      }
    }
  });

  it("refuses same-name nesting across many boundaries with work linear in the file", async () => {
    // Every boundary inside the outer record starts a check that opens a record the reading through
    // it does not; waiting for each to close read the rest of the file once per boundary.
    const write = RecordScanner.prototype.write;
    let scanned = 0;
    RecordScanner.prototype.write = function (this: RecordScanner, bytes: Uint8Array) {
      scanned += bytes.length;
      write.call(this, bytes);
    };
    try {
      for (const size of [3, 4, 5, 7]) {
        for (const n of [256, 1024]) {
          const bytes = encoder.encode("<d>" + "<r>".repeat(n) + "</r>".repeat(n) + "</d>");
          scanned = 0;
          await expect(
            Xml.readRecords(bytes.subarray(3), {
              offset: 3,
              size,
              context: { root: "d", entities: {} } as unknown as XmlPrologContext,
              tags: ["r"],
              maxDepth: 0,
              onRecord() {}
            })
          ).rejects.toThrow(XmlChunkAmbiguityError);
          expect(scanned, `size ${size}, n ${n}`).toBeLessThanOrEqual(4 * bytes.length);
        }
      }
    } finally {
      RecordScanner.prototype.write = write;
    }
  });

  it("holds the guarantee for random partitions into ranges of unequal size", async () => {
    // Ranges need not be equal: any cut points make a partition, and each must keep the promise.
    let agreed = 0;
    for (let seed = 1; seed <= 40; seed++) {
      const { xml } = generateDocument(seed * 7919, {
        decoys: seed % 2 === 0,
        sameNameNesting: seed % 3 !== 0,
        oneLine: seed % 4 === 0,
        doctype: seed % 5 === 0
      });
      const bytes = encoder.encode(xml);
      const { whole, context } = await wholeOf(bytes);
      const next = random(seed);
      for (let trial = 0; trial < 60; trial++) {
        const cuts = new Set<number>([0]);
        const count = 1 + Math.floor(next() * 8);
        while (cuts.size <= count) {
          cuts.add(Math.floor(next() * bytes.length));
        }
        const starts = [...cuts].sort((a, b) => a - b);
        const got: string[] = [];
        const results: { structure: XmlPartitionStructure }[] = [];
        let failed = false;
        for (let i = 0; i < starts.length; i++) {
          const offset = starts[i];
          const size = i + 1 < starts.length ? starts[i + 1] - offset : undefined;
          try {
            const { records, structure } = await read(bytes, {
              tags: ["r"],
              offset,
              size,
              context
            });
            results.push({ structure });
            got.push(...records.map(key));
          } catch (err) {
            expect(err, String(err)).toBeInstanceOf(XmlChunkAmbiguityError);
            failed = true;
          }
        }
        if (!failed) {
          // Joined as a worker's results would be: sent back, then verified.
          try {
            Xml.verifyPartition(JSON.parse(JSON.stringify(results)));
          } catch (err) {
            expect(err, String(err)).toBeInstanceOf(XmlChunkAmbiguityError);
            failed = true;
          }
        }
        if (!failed) {
          expect(got, `seed ${seed}, ranges at ${starts.join(",")}`).toEqual(whole);
          agreed++;
        }
      }
    }
    expect(agreed).toBeGreaterThan(0);
  }, 60_000);

  it("checks later chunks with the entities it read itself, when read from the start", async () => {
    // Read from offset 0 without a context: the range after the boundary uses `&e;`, which the
    // reading from the start declared, so checking it needs no context from the caller.
    const xml = '<!DOCTYPE R [<!ENTITY e "E">]><R><T>xxxx&e;xxxx</T><T/></R>';
    const { records } = await read(encoder.encode(xml), { tags: ["T"], size: 38 });
    expect(records.map(record => record.object)).toEqual([{ T: "xxxxExxxx" }]);
  });

  it("says a chunk is ambiguous when it starts inside a record holding one of the same name", async () => {
    const xml = "<doc><r id='outer'><x>....</x><r id='inner'/><x>....</x></r><r id='next'/></doc>";
    const bytes = encoder.encode(xml);
    const { context } = await Xml.readRoot(bytes);
    const start = xml.indexOf("<x>....</x><r id='inner'");
    await expect(read(bytes, { tags: ["r"], offset: 0, size: start, context })).rejects.toThrow(
      XmlChunkAmbiguityError
    );
    const err = await read(bytes, { tags: ["r"], offset: 0, size: start }).catch(e => e);
    expect(err).toBeInstanceOf(XmlChunkAmbiguityError);
    expect(err.chunkStart).toBe(start);
    expect(err.atByte).toBe(xml.indexOf("<r id='inner'"));
  });

  it("says a chunk is ambiguous when it starts inside a comment that holds a record", async () => {
    const xml = "<doc><!-- old: <r id='gone'>x</r> --><r id='1'/></doc>";
    const bytes = encoder.encode(xml);
    const size = xml.indexOf("old:");
    await expect(read(bytes, { tags: ["r"], offset: 0, size })).rejects.toThrow(
      XmlChunkAmbiguityError
    );
  });

  it("reads past the range only as far as it must", async () => {
    let xml = "<doc>\n";
    for (let i = 0; i < 200; i++) {
      xml += `  <r id="${i}"><t>text ${i}</t></r>\n`;
    }
    xml += "</doc>\n";
    const bytes = encoder.encode(xml);
    const { context } = await Xml.readRoot(bytes);
    for (const size of [100, 1000, 3000]) {
      for (let offset = 0; offset < bytes.length; offset += size) {
        const { endByte } = await read(bytes, { tags: ["r"], offset, size, context });
        expect(endByte - Math.min(bytes.length, offset + size)).toBeLessThanOrEqual(1024);
      }
    }
  });

  it("agrees with the whole read from every starting byte, or says it cannot", async () => {
    for (const seed of [3, 8, 11]) {
      const { xml } = generateDocument(seed, { decoys: true, sameNameNesting: seed === 3 });
      const bytes = encoder.encode(xml);
      const { whole, context } = await wholeOf(bytes);
      for (let offset = 1; offset < bytes.length; offset++) {
        const got: string[] = [];
        let ambiguous = false;
        for (const [start, size] of [
          [0, offset],
          [offset, undefined]
        ] as const) {
          try {
            const { records } = await read(bytes, { tags: ["r"], offset: start, size, context });
            got.push(...records.map(key));
          } catch (err) {
            expect(err).toBeInstanceOf(XmlChunkAmbiguityError);
            ambiguous = true;
          }
        }
        if (!ambiguous) {
          expect(got, `seed ${seed}, offset ${offset}`).toEqual(whole);
        }
      }
    }
  });
});

describe("Xml.readRecords — truncated input", () => {
  it("throws XmlTruncatedError with the record's start when the input ends inside it", async () => {
    const xml = "<doc><r id='1'>one</r><r id='2'>two</r></doc>";
    const cut = xml.indexOf("two");
    const err = await read(encoder.encode(xml.slice(0, cut)), { tags: ["r"] }).catch(e => e);
    expect(err).toBeInstanceOf(XmlTruncatedError);
    expect(err.recordStartByte).toBe(xml.indexOf("<r id='2'"));
  });

  it("says the input is cut when a range read to its end has not seen the root end", async () => {
    const xml = "<doc><r id='1'><x/><r id='2'/><x/><r id='3'/></r></doc>";
    const { context } = await Xml.readRoot(encoder.encode(xml));
    const cut = encoder.encode(xml.slice(0, xml.indexOf("</r></doc>")));
    // Read from inside record 1, which the cut leaves open: its inner records look like records.
    const offset = xml.indexOf("<r id='2'");
    const err = await read(cut, { tags: ["r"], offset, context }).catch(e => e);
    expect(err).toBeInstanceOf(XmlTruncatedError);
    expect(err.message).toMatch(/inside <doc>/);
  });

  it("joins a range that starts inside markup after the root, and finds text there", async () => {
    const join = async (xml: string, offset: number): Promise<unknown> => {
      const bytes = encoder.encode(xml);
      const context = { root: "doc", entities: {} };
      try {
        const first = await read(bytes, { tags: ["r"], offset: 0, size: offset, context });
        const rest = await read(bytes, { tags: ["r"], offset, context });
        Xml.verifyPartition([first, rest]);
        return "joined";
      } catch (err) {
        return err;
      }
    };
    // Inside the root's end tag, a trailing comment or processing instruction, or a root that closes
    // itself: the range reads the end of that markup as text, and the range before says it is none.
    for (const xml of [
      "<doc><r/></doc>\n<!-- trailer -->\n",
      "<doc/>\n<?pi a b?>\n",
      "<doc a='/doc>'/>\n"
    ]) {
      for (let offset = 1; offset < xml.length; offset++) {
        expect(await join(xml, offset), `${JSON.stringify(xml)} at ${offset}`).toBe("joined");
      }
    }
    // Text after the root is found wherever the boundary falls: in the markup before it, or in it.
    const garbage = "<doc><r/></doc>garbage";
    for (let offset = 1; offset < garbage.length; offset++) {
      expect(await join(garbage, offset), `at ${offset}`).toBeInstanceOf(XmlParseError);
    }
    // Without what the range before read past its end, the text cannot be placed, and the join says so.
    expect(() =>
      Xml.verifyPartition([
        { structure: { start: 0, end: 1, final: false, steps: [["e", 0]], open: [] } },
        { structure: { start: 1, end: null, final: true, eof: 9, steps: [["l", 1]], open: [] } }
      ])
    ).toThrow(XmlChunkAmbiguityError);
  });

  it("refuses structures that cannot be a chunk's", () => {
    const whole = {
      start: 0,
      end: null,
      final: true,
      eof: 4,
      steps: [["e", 0]],
      open: []
    } as const;
    for (const structure of [
      { ...whole, start: -1 },
      { start: 0, end: 0, final: true, eof: 4, steps: [], open: [] },
      { ...whole, eof: undefined },
      { ...whole, final: false, end: 9, eof: 4 },
      {
        ...whole,
        steps: [
          ["e", 5],
          ["e", 2]
        ]
      },
      { start: 0, end: 4, final: false, steps: [["e", 9]], open: [] },
      { ...whole, open: ["1bad"] },
      { ...whole, steps: [["/", 0, "a b"]] },
      { start: 0, end: 4, final: false, steps: [], open: [], follow: 2 }
    ]) {
      expect(
        () => Xml.verifyPartition([{ structure } as never]),
        JSON.stringify(structure)
      ).toThrow(XmlArgumentError);
    }
    // What a worker sends back as JSON is what it read.
    expect(() =>
      Xml.verifyPartition(JSON.parse(JSON.stringify([{ structure: whole }])))
    ).not.toThrow();
  });

  it("keeps no more of a chunk's structure than its nesting needs, however many records it reads", async () => {
    const bytes = encoder.encode(`<doc>${"<r/>x".repeat(1000)}</doc>`);
    const { structure } = await read(bytes, {
      tags: ["r"],
      offset: 5,
      size: bytes.length - 11,
      context: { root: "doc", entities: {} }
    });
    expect(structure.steps.length).toBeLessThanOrEqual(3);
  });

  it("does not let a range of white space hide a cut between two records", async () => {
    const xml = "<doc>\n  <r id='1'/>\n  <r id='2'/>\n  <r id='3'/>\n</doc>\n";
    const cut = encoder.encode(xml.slice(0, xml.indexOf("<r id='2'")));
    const { context } = await Xml.readRoot(cut);
    // The last range holds only the indentation before the lost records: it reads, and the join finds
    // the root still open at the end of the file — if the first range has not already found the end
    // of the file, reading on to place its boundary.
    const offset = xml.indexOf("<r id='2'") - 2;
    const failure = await (async () => {
      try {
        const first = await read(cut, { tags: ["r"], offset: 0, size: offset, context });
        const rest = await read(cut, { tags: ["r"], offset, context });
        Xml.verifyPartition([first, rest]);
        return undefined;
      } catch (err) {
        return err;
      }
    })();
    expect(failure).toBeInstanceOf(XmlTruncatedError);
  });

  it("never gives a record the input ends inside, wherever it is cut and however it is split", async () => {
    for (const [seed, options] of [
      [21, { decoys: true }],
      [6, { sameNameNesting: true, doctype: true }],
      [9, { decoys: true, sameNameNesting: true, oneLine: true }]
    ] as const) {
      const { xml } = generateDocument(seed, options);
      const full = encoder.encode(xml);
      const { whole } = await wholeOf(full);
      const wholeRecords = (await read(full, { tags: ["r"] })).records;
      const { context } = await Xml.readRoot(full);
      const rootEnd = full.lastIndexOf(0x3e) + 1;
      for (let length = 1; length < full.length; length++) {
        const bytes = full.subarray(0, length);
        // The whole read of a cut file fails as truncated.
        const wholeError = await read(bytes, { tags: ["r"] }).catch(e => e);
        if (length < xml.lastIndexOf("</doc>") + 6) {
          expect(wholeError, `length ${length}`).toBeInstanceOf(XmlTruncatedError);
        }
        const cutInside = wholeRecords.find(r => r.startByte < length && r.endByte > length);
        for (const size of [1, 7, 13, 50, length + 1]) {
          let owner: unknown;
          let failed = false;
          const delivered: string[] = [];
          const structures: { structure: XmlPartitionStructure }[] = [];
          for (let offset = 0; offset < length; offset += size) {
            try {
              const { records, structure } = await read(bytes, {
                tags: ["r"],
                offset,
                size,
                context
              });
              structures.push({ structure });
              for (const record of records) {
                expect(record.endByte).toBeLessThanOrEqual(length);
                delivered.push(key(record));
              }
            } catch (err) {
              failed = true;
              expect(
                err instanceof XmlTruncatedError || err instanceof XmlChunkAmbiguityError,
                String(err)
              ).toBe(true);
              if (
                cutInside &&
                cutInside.startByte >= offset &&
                cutInside.startByte < offset + size
              ) {
                owner = err;
              }
            }
          }
          if (cutInside) {
            // The range the cut record belongs to does not give it: it says why.
            expect(owner, `length ${length}, size ${size}`).toBeDefined();
          }
          if (!failed) {
            try {
              Xml.verifyPartition(structures);
            } catch (err) {
              failed = true;
              expect(
                err instanceof XmlTruncatedError || err instanceof XmlChunkAmbiguityError,
                String(err)
              ).toBe(true);
            }
          }
          // A cut file has no partition that reads, and joins, without an error: there is never a
          // whole batch to publish.
          if (length < rootEnd) {
            expect(failed, `length ${length}, size ${size}`).toBe(true);
          }
          if (!failed) {
            // Only a batch that read without an error can be published: it gives only the file's records.
            for (const record of delivered) {
              expect(whole).toContain(record);
            }
          }
        }
      }
    }
  }, 120_000);
});

describe("Xml.readRecords — what a partition promises", () => {
  /**
   * Read `bytes` as ranges of `size`, every range to completion: the records of each, or why it
   * failed. What a caller would stage before deciding whether to publish.
   */
  async function rangesOf(bytes: Uint8Array, size: number, context: XmlPrologContext) {
    const out: Array<{ records: string[]; structure: XmlPartitionStructure } | { error: unknown }> =
      [];
    for (let offset = 0; offset < bytes.length; offset += size) {
      try {
        const { records, structure } = await read(bytes, { tags: ["r"], offset, size, context });
        out.push({ records: records.map(key), structure });
      } catch (error) {
        out.push({ error });
      }
    }
    return out;
  }

  /** What stops a partition's batch from being published: the ranges that failed, or else what the join found. */
  function failuresOf(ranges: Awaited<ReturnType<typeof rangesOf>>): unknown[] {
    const errors = ranges
      .filter(range => "error" in range)
      .map(range => (range as { error: unknown }).error);
    if (errors.length > 0) {
      return errors;
    }
    try {
      Xml.verifyPartition(ranges as { structure: XmlPartitionStructure }[]);
      return [];
    } catch (err) {
      return [err];
    }
  }

  /**
   * Every partition of a damaged file — cut short, or not one well-formed document — has a range that
   * fails or does not join: there is never a whole batch to publish.
   */
  async function expectEveryPartitionFails(xml: string, cutAt = xml.length): Promise<void> {
    const full = encoder.encode(xml);
    const { context } = await Xml.readRoot(full);
    const bytes = full.subarray(0, cutAt);
    await expect(read(bytes, { tags: ["r"] })).rejects.toBeInstanceOf(XmlParseError);
    for (let size = 1; size <= bytes.length + 1; size++) {
      const failures = failuresOf(await rangesOf(bytes, size, context));
      expect(failures.length, `size ${size}`).toBeGreaterThan(0);
      for (const error of failures) {
        expect(
          error instanceof XmlParseError || error instanceof XmlChunkAmbiguityError,
          String(error)
        ).toBe(true);
      }
    }
  }

  /** Every partition of a whole file either gives its records exactly, or has a range that says it cannot. */
  async function expectEveryPartitionHolds(xml: string): Promise<void> {
    const bytes = encoder.encode(xml);
    const { whole, context } = await wholeOf(bytes);
    for (let size = 1; size <= bytes.length + 1; size++) {
      const ranges = await rangesOf(bytes, size, context);
      const failures = failuresOf(ranges);
      for (const error of failures) {
        // A whole file is never reported damaged; a range may only be unplaceable.
        expect(error, `size ${size}`).toBeInstanceOf(XmlChunkAmbiguityError);
      }
      if (failures.length === 0) {
        expect(
          ranges.flatMap(range => (range as { records: string[] }).records),
          `size ${size}`
        ).toEqual(whole);
      }
    }
  }

  it("never publishes a file that is not one whole document, however it is split", async () => {
    // Every range of these can read without an error: only the join finds what is wrong.
    await expectEveryPartitionFails("<doc><r/><doc>inner</doc>"); // the root never closes
    await expectEveryPartitionFails("<doc><r/></doc><doc><r/></doc>"); // a second root
    await expectEveryPartitionFails("<doc><x><r/></doc>"); // </doc> closes <x>
    await expectEveryPartitionFails("<doc><r/></doc>trailing text"); // text after the root
    await expectEveryPartitionFails("<doc><r/></doc><r/>"); // an element after the root
    await expectEveryPartitionFails("<doc><r/></doc>trailing text"); // text read as the input closes
    await expectEveryPartitionFails("<doc><r/></doc><!-- a --> b"); // text after markup after the root
    await expectEveryPartitionFails("<doc/> <![CDATA[x]]>"); // a CDATA section after the root
  });

  it("gives records and then rejects: a record handed over does not make the read a success", async () => {
    const xml = "<doc><r id='1'/><r id='2'>cut here";
    const seen: string[] = [];
    const read = Xml.readRecords(encoder.encode(xml), {
      tags: ["r"],
      onRecord: record =>
        void seen.push(String((record.object.r as Record<string, unknown>)["@_id"]))
    });
    await expect(read).rejects.toBeInstanceOf(XmlTruncatedError);
    expect(seen).toEqual(["1"]);
  });

  it("fails some range of every partition of a cut file, even when earlier ranges succeed", async () => {
    let xml = "<doc>\n";
    for (let i = 0; i < 8; i++) {
      xml += `  <r id="${i}"><t>text ${i}</t></r>\n`;
    }
    xml += "</doc>\n";
    const bytes = encoder.encode(xml);
    const { context } = await Xml.readRoot(bytes);
    const cutAt = xml.indexOf('<r id="6"') + 12;
    // With small ranges the first ones succeed on their own: only the batch tells.
    const ranges = await rangesOf(bytes.subarray(0, cutAt), 40, context);
    expect(ranges.some(range => "records" in range && range.records.length > 0)).toBe(true);
    expect(ranges.some(range => "error" in range)).toBe(true);
    for (const at of [cutAt, xml.indexOf("</doc>"), xml.indexOf('<r id="3"'), 30]) {
      await expectEveryPartitionFails(xml, at);
    }
  });

  it("holds for a long record holding records of its own name, cut or whole", async () => {
    const xml =
      "<doc><r id='outer'><n>aaaaaaaa</n><r id='in1'/><n>bbbbbbbb</n><r id='in2'><x/></r><n>cccc</n></r><r id='after'/></doc>";
    await expectEveryPartitionHolds(xml);
    await expectEveryPartitionFails(xml, xml.indexOf("<r id='in2'") + 3);
    await expectEveryPartitionFails(xml, xml.indexOf("<n>cccc"));
  });

  it("is not fooled by an element named like the root", async () => {
    // The end tag of an element named `doc`, with nothing open around it as a later range sees it,
    // must not pass for the end of the document when the document is cut after it.
    const xml = "<doc><r id='1'><doc>inner</doc><t>more text</t></r><r id='2'/></doc>";
    await expectEveryPartitionHolds(xml);
    await expectEveryPartitionFails(xml, xml.indexOf("</doc>") + 6);
    await expectEveryPartitionFails(xml, xml.indexOf("<r id='2'"));
  });

  it("is not fooled by a root end tag in a comment or a CDATA section", async () => {
    const xml =
      "<doc><r id='1'><!-- </doc> --><t>a</t><![CDATA[</doc>]]><t>b</t></r><r id='2'/></doc>";
    await expectEveryPartitionHolds(xml);
    await expectEveryPartitionFails(xml, xml.indexOf("]]>") + 3);
    await expectEveryPartitionFails(xml, xml.indexOf("-->") + 3);
  });

  it("gives the whole file's records, once, when a failed batch is read again whole", async () => {
    const xml = "<doc><r id='outer'><x>....</x><r id='inner'/><x>....</x></r><r id='next'/></doc>";
    const bytes = encoder.encode(xml);
    const { whole, context } = await wholeOf(bytes);
    const ranges = await rangesOf(bytes, xml.indexOf("<x>"), context);
    expect(ranges.some(range => "error" in range)).toBe(true);
    // The failed batch is discarded whole; reading again gives each record exactly once.
    const again = await read(bytes, { tags: ["r"] });
    expect(again.records.map(key)).toEqual(whole);
  });

  it("decodes a range of an ASCII document as the whole read does", async () => {
    // The declaration changes what a later byte means; a range that starts past it learns that from
    // the context, or a byte past 0x7F would be a character there and none in the whole read.
    for (const name of ["ascii", "us-ascii"]) {
      const xml = `<?xml version="1.0" encoding="${name}"?><d><a/><r>é</r></d>`;
      const bytes = encoder.encode(xml);
      const root = await Xml.readRoot(bytes);
      expect(root.context.encoding).toBe("ascii");
      const context = JSON.parse(JSON.stringify(root.context)) as XmlPrologContext;
      const cut = xml.indexOf("<r>");
      const oneByte = function* (offset: number): Generator<Uint8Array> {
        for (let at = offset; at < bytes.length; at++) {
          yield bytes.subarray(at, at + 1);
        }
      };
      for (const invalidCharHandling of ["error", "replace", "skip"] as const) {
        const whole: unknown[] = [];
        const wholeRead = await Xml.readRecords(oneByte(0), {
          tags: ["r"],
          invalidCharHandling,
          onRecord: record => void whole.push(record.object)
        }).catch((error: unknown) => error);
        const parts: unknown[] = [];
        let partition: unknown;
        try {
          const results: XmlReadRecordsResult[] = [];
          for (const [offset, size] of [
            [0, cut],
            [cut, bytes.length - cut]
          ]) {
            results.push(
              await Xml.readRecords(oneByte(offset), {
                offset,
                size,
                context,
                tags: ["r"],
                invalidCharHandling,
                onRecord: record => void parts.push(record.object)
              })
            );
          }
          Xml.verifyPartition(JSON.parse(JSON.stringify(results)));
        } catch (error) {
          partition = error;
        }
        if (invalidCharHandling === "error") {
          expect(wholeRead).toBeInstanceOf(XmlParseError);
          expect(partition).toBeInstanceOf(XmlChunkAmbiguityError);
        } else {
          expect(partition).toBeUndefined();
          expect(parts, `${name} ${invalidCharHandling}`).toEqual(whole);
        }
      }
    }
  });
});

describe("Xml.readRecords — holding one record, and sources", () => {
  it("builds the next record only once a callback's promise has settled", async () => {
    // Reaching into the builder is the only way to see work done ahead of the consumer.
    const built: number[] = [];
    let builds = 0;
    const openTag = ObjectBuilder.prototype.openTag;
    ObjectBuilder.prototype.openTag = function (this: ObjectBuilder, tag) {
      builds++;
      return openTag.call(this, tag);
    };
    try {
      await Xml.readRecords(encoder.encode(`<d>${"<r/>".repeat(500)}</d>`), {
        tags: ["r"],
        onRecord: async () => {
          built.push(builds);
          await Promise.resolve();
        }
      });
    } finally {
      ObjectBuilder.prototype.openTag = openTag;
    }
    // When each callback runs, exactly its own record has been built.
    expect(built).toEqual(Array.from({ length: 500 }, (_, i) => i + 1));
  });

  it("reads a source that fills the same buffer for every chunk", async () => {
    async function* reused(): AsyncGenerator<Uint8Array> {
      const buffer = new Uint8Array(3);
      for (const text of ["<R>", "<T>", "abc", "</T", "></", "R> "]) {
        buffer.set(encoder.encode(text));
        yield buffer;
      }
    }
    const objects: unknown[] = [];
    await Xml.readRecords(reused(), {
      tags: ["T"],
      onRecord: record => void objects.push(record.object)
    });
    expect(objects).toEqual([{ T: "abc" }]);
  });

  it("gives a whole record at once, without waiting for more of the source", async () => {
    const controller = new AbortController();
    let pulls = 0;
    let calls = 0;
    const quiet: AsyncIterable<Uint8Array> = {
      [Symbol.asyncIterator]: () => ({
        next: () =>
          ++pulls === 1
            ? Promise.resolve({ done: false, value: encoder.encode("<R><T/>") })
            : new Promise<IteratorResult<Uint8Array>>(() => {}),
        return: async () => ({ done: true, value: undefined })
      })
    };
    const reading = Xml.readRecords(quiet, {
      tags: ["T"],
      signal: controller.signal,
      onRecord: () => void calls++
    }).catch(e => e);
    await new Promise(resolve => setTimeout(resolve, 20));
    expect(calls).toBe(1);
    controller.abort();
    expect(await reading).toBeInstanceOf(AbortError);
  });

  it("passes a callback's error out as it is, from any start", async () => {
    const bytes = encoder.encode("<R><T/></R>");
    const context = { root: "R", entities: {} };
    const thrown = new XmlParseError("the callback's own");
    for (const offset of [0, 3]) {
      const sync = await Xml.readRecords(bytes.subarray(offset), {
        tags: ["T"],
        offset,
        context,
        onRecord: () => {
          throw thrown;
        }
      }).catch(e => e);
      expect(sync, `offset ${offset}`).toBe(thrown);
      const rejected = await Xml.readRecords(bytes.subarray(offset), {
        tags: ["T"],
        offset,
        context,
        onRecord: () => Promise.reject(thrown)
      }).catch(e => e);
      expect(rejected, `offset ${offset}`).toBe(thrown);
    }
  });

  it("leaves no unhandled rejection behind an abort", async () => {
    const unhandled: unknown[] = [];
    const record = (reason: unknown): void => void unhandled.push(reason);
    const process = (
      globalThis as {
        process?: {
          on(e: string, f: (r: unknown) => void): void;
          off(e: string, f: (r: unknown) => void): void;
        };
      }
    ).process;
    const onWindow = (event: PromiseRejectionEvent): void => record(event.reason);
    process?.on("unhandledRejection", record);
    (globalThis as { addEventListener?: typeof addEventListener }).addEventListener?.(
      "unhandledrejection",
      onWindow
    );
    try {
      const fromCallback = new AbortController();
      await Xml.readRecords(encoder.encode("<R><T/></R>"), {
        tags: ["T"],
        signal: fromCallback.signal,
        onRecord: () => {
          fromCallback.abort();
          return Promise.reject(new Error("callback"));
        }
      }).catch(() => undefined);
      const fromSource = new AbortController();
      await Xml.readRoot(
        {
          [Symbol.asyncIterator]: () => ({
            next: () => {
              fromSource.abort();
              return Promise.reject(new Error("source"));
            }
          })
        },
        { signal: fromSource.signal }
      ).catch(() => undefined);
      await new Promise(resolve => setTimeout(resolve, 20));
      expect(unhandled).toEqual([]);
    } finally {
      process?.off("unhandledRejection", record);
      (globalThis as { removeEventListener?: typeof removeEventListener }).removeEventListener?.(
        "unhandledrejection",
        onWindow
      );
    }
  });

  it("leaves no listener on the signal once a read has ended, however it ended", async () => {
    const controller = new AbortController();
    const signal = controller.signal;
    let listeners = 0;
    const add = signal.addEventListener.bind(signal);
    const remove = signal.removeEventListener.bind(signal);
    signal.addEventListener = ((
      type: string,
      listener: EventListener,
      options?: AddEventListenerOptions
    ) => {
      listeners++;
      add(type, listener, options);
    }) as typeof signal.addEventListener;
    signal.removeEventListener = ((type: string, listener: EventListener) => {
      listeners--;
      remove(type, listener);
    }) as typeof signal.removeEventListener;
    const async = (): Promise<void> => Promise.resolve();
    await Xml.readRecords(encoder.encode(`<R>${"<T/>".repeat(100)}</R>`), {
      tags: ["T"],
      signal,
      onRecord: async
    });
    expect(listeners).toBe(0);
    await Xml.readRecords(encoder.encode("<R><T/><1/></R>"), {
      tags: ["T"],
      signal,
      onRecord: async
    }).catch(() => undefined);
    expect(listeners).toBe(0);
    await Xml.readRoot(encoder.encode("<R/>"), { signal });
    expect(listeners).toBe(0);
  });

  it("parses nothing once the signal has fired after a chunk arrived", async () => {
    const controller = new AbortController();
    queueMicrotask(() => controller.abort());
    await expect(
      Xml.readRoot(encoder.encode("<R/>"), { signal: controller.signal })
    ).rejects.toBeInstanceOf(AbortError);
    const records = new AbortController();
    let calls = 0;
    queueMicrotask(() => records.abort());
    await expect(
      Xml.readRecords(encoder.encode("<R><T/></R>"), {
        tags: ["T"],
        signal: records.signal,
        onRecord: () => void calls++
      })
    ).rejects.toBeInstanceOf(AbortError);
    expect(calls).toBe(0);
  });

  it("keeps the read's error, and no listener, when looking up the source's return() throws", async () => {
    const controller = new AbortController();
    const signal = controller.signal;
    let listeners = 0;
    const add = signal.addEventListener.bind(signal);
    const remove = signal.removeEventListener.bind(signal);
    signal.addEventListener = ((...args: Parameters<typeof add>) => {
      listeners++;
      add(...args);
    }) as typeof signal.addEventListener;
    signal.removeEventListener = ((...args: Parameters<typeof remove>) => {
      listeners--;
      remove(...args);
    }) as typeof signal.removeEventListener;
    const source = (bytes: string): AsyncIterable<Uint8Array> => ({
      [Symbol.asyncIterator]: () => ({
        next: async () => ({ done: false, value: encoder.encode(bytes) }),
        get return(): never {
          throw new Error("return getter");
        }
      })
    });
    await expect(Xml.readRoot(source("<1/>"), { signal })).rejects.toBeInstanceOf(XmlParseError);
    expect(listeners).toBe(0);
    await expect(Xml.readRoot(source("<R/>"), { signal })).rejects.toThrow("return getter");
    expect(listeners).toBe(0);
  });

  it("looks the source's return() up once", async () => {
    let lookups = 0;
    let calls = 0;
    const source: AsyncIterable<Uint8Array> = {
      [Symbol.asyncIterator]: () => ({
        next: async () => ({ done: false, value: encoder.encode("<R/>") }),
        get return() {
          return ++lookups === 1
            ? () => {
                calls++;
                return Promise.resolve({ done: true as const, value: undefined });
              }
            : undefined;
        }
      })
    };
    await Xml.readRoot(source);
    expect([lookups, calls]).toEqual([1, 1]);
  });

  it("passes out what isArray throws, as it is, from any start", async () => {
    const own = new XmlParseError("the caller's own");
    for (const offset of [0, 3]) {
      const err = await Xml.readRecords(encoder.encode("<R><T><C/></T></R>").subarray(offset), {
        tags: ["T"],
        offset,
        context: { root: "R", entities: {} },
        isArray: () => {
          throw own;
        },
        onRecord: () => {}
      }).catch(e => e);
      expect(err, `offset ${offset}`).toBe(own);
    }
  });

  it("refuses input whose byte positions would pass Number.MAX_SAFE_INTEGER", async () => {
    await expect(
      Xml.readRecords(encoder.encode("<T/></R>"), {
        tags: ["T"],
        offset: Number.MAX_SAFE_INTEGER - 2,
        context: { root: "R", entities: {} },
        onRecord: () => {}
      })
    ).rejects.toBeInstanceOf(XmlArgumentError);
  });

  it("reports a root start tag too long as soon as that is known, however it is chunked", async () => {
    const waiting: AsyncIterable<Uint8Array> = {
      [Symbol.asyncIterator]: () => {
        let pulls = 0;
        return {
          next: () =>
            ++pulls === 1
              ? Promise.resolve({ done: false, value: encoder.encode('<R a="123456789') })
              : new Promise<IteratorResult<Uint8Array>>(() => {})
        };
      }
    };
    await expect(Xml.readRoot(waiting, { maxTagBytes: 4 })).rejects.toBeInstanceOf(
      XmlLimitExceededError
    );
    // A syntax error past the limit is the limit, whether it arrives with the tag or after it.
    const bytes = encoder.encode('<R a="123456789 <"/>');
    for (const size of [1, 3, 7, bytes.length]) {
      const parts: Uint8Array[] = [];
      for (let at = 0; at < bytes.length; at += size) {
        parts.push(bytes.subarray(at, at + size));
      }
      await expect(
        Xml.readRoot(parts, { maxTagBytes: 4 }),
        `chunks of ${size}`
      ).rejects.toBeInstanceOf(XmlLimitExceededError);
    }
  });

  it("leaves no unhandled rejection when closing fires the signal and fails", async () => {
    const unhandled: unknown[] = [];
    const record = (reason: unknown): void => void unhandled.push(reason);
    const host = (
      globalThis as {
        process?: {
          on(e: string, f: (r: unknown) => void): void;
          off(e: string, f: (r: unknown) => void): void;
        };
      }
    ).process;
    const onWindow = (event: PromiseRejectionEvent): void => record(event.reason);
    host?.on("unhandledRejection", record);
    (globalThis as { addEventListener?: typeof addEventListener }).addEventListener?.(
      "unhandledrejection",
      onWindow
    );
    try {
      const controller = new AbortController();
      const source: AsyncIterable<Uint8Array> = {
        [Symbol.asyncIterator]: () => ({
          next: async () => ({ done: false, value: encoder.encode("<R/>") }),
          return: () => {
            controller.abort();
            return Promise.reject(new Error("close"));
          }
        })
      };
      await expect(Xml.readRoot(source, { signal: controller.signal })).rejects.toBeInstanceOf(
        AbortError
      );
      await new Promise(resolve => setTimeout(resolve, 20));
      expect(unhandled).toEqual([]);
    } finally {
      host?.off("unhandledRejection", record);
      (globalThis as { removeEventListener?: typeof removeEventListener }).removeEventListener?.(
        "unhandledrejection",
        onWindow
      );
    }
  });

  it("gives out no record once isArray has fired the signal", async () => {
    const controller = new AbortController();
    let calls = 0;
    await expect(
      Xml.readRecords(encoder.encode("<R><T><C/></T></R>"), {
        tags: ["T"],
        signal: controller.signal,
        isArray: () => {
          controller.abort();
          return false;
        },
        onRecord: () => void calls++
      })
    ).rejects.toBeInstanceOf(AbortError);
    expect(calls).toBe(0);
  });

  it("refuses a source whose iterator gives no result object", async () => {
    for (const value of [null, undefined, 7]) {
      const source = {
        [Symbol.asyncIterator]: () => ({ next: () => Promise.resolve(value) })
      } as unknown as AsyncIterable<Uint8Array>;
      const err = await Xml.readRoot(source).catch(e => e);
      expect(err).toBeInstanceOf(XmlArgumentError);
      expect(err.field).toBe("source");
    }
  });

  it("reports no root before maxStartByte without waiting for more of the source", async () => {
    const waiting: AsyncIterable<Uint8Array> = {
      [Symbol.asyncIterator]: () => {
        let pulls = 0;
        return {
          next: () =>
            ++pulls === 1
              ? Promise.resolve({ done: false, value: encoder.encode(" ".repeat(20)) })
              : new Promise<IteratorResult<Uint8Array>>(() => {})
        };
      }
    };
    await expect(Xml.readRoot(waiting, { maxStartByte: 4 })).rejects.toBeInstanceOf(
      XmlLimitExceededError
    );
  });

  it("gives the same root, or the same error, under every limit however the file is chunked", async () => {
    // A limit and a parse error are told apart by which comes first in the file.
    const files = [
      '<R a="123456789" a="x"/>',
      '<R a="123456789" bad?>',
      '<R a="1&missing;2"/>',
      "<!--c--><R/>",
      '  <?pi x?> <R x="é€😀"/>',
      "<1/>",
      '<R\r\n a="1"/>',
      "\uFEFF<R/>"
    ].map(text => encoder.encode(text));
    files.push(
      new Uint8Array([...encoder.encode('<R a="'), 0xe2, 0x3c, 0x22, 0x2f, 0x3e]),
      new Uint8Array([0x20, 0xff, 0x3c, 0x52, 0x2f, 0x3e])
    );
    for (const bytes of files) {
      for (let maxStartByte = 1; maxStartByte <= bytes.length + 1; maxStartByte++) {
        for (let maxTagBytes = 1; maxTagBytes <= bytes.length + 1; maxTagBytes++) {
          const outcomes = new Set<string>();
          for (const size of [1, 2, 3, bytes.length]) {
            const parts: Uint8Array[] = [];
            for (let at = 0; at < bytes.length; at += size) {
              parts.push(bytes.subarray(at, at + size));
            }
            outcomes.add(
              await Xml.readRoot(parts, { maxStartByte, maxTagBytes }).then(
                root => `ok ${root.name}`,
                (e: Error) => `${e.name}: ${e.message}`
              )
            );
          }
          expect([...outcomes], `${bytes.join(",")} ${maxStartByte}/${maxTagBytes}`).toHaveLength(
            1
          );
        }
      }
    }
  });

  it("checks a parser built while the call builds its own, and uses the context it checked", async () => {
    let accepted = 0;
    const context = {
      root: "R",
      get entities() {
        try {
          new Xml.SaxParser({ maxDepth: -1 });
          accepted++;
        } catch {
          // refused, as it should be
        }
        return {};
      }
    };
    await Xml.readRecords(encoder.encode("<T/></R>"), {
      offset: 3,
      context,
      tags: ["T"],
      onRecord: () => {}
    }).catch(() => undefined);
    expect(accepted).toBe(0);
    // A getter answering differently later cannot hand the parsers what was never checked.
    let reads = 0;
    const changing = {
      root: "R",
      get entities(): Record<string, string> {
        return ++reads === 1 ? { e: "E" } : { e: "<bad" };
      }
    };
    const objects: unknown[] = [];
    await Xml.readRecords(encoder.encode("<T>&e;</T></R>"), {
      offset: 3,
      context: changing,
      tags: ["T"],
      onRecord: record => void objects.push(record.object)
    });
    expect([reads, objects]).toEqual([1, [{ T: "E" }]]);
  });

  it("reads a character that straddles the root's limits whole before it judges them", async () => {
    // `<😀/>` is seven bytes: the emoji runs past maxStartByte + maxTagBytes = 4.
    const bytes = encoder.encode("<😀/>");
    for (const size of [1, 2, bytes.length]) {
      const parts: Uint8Array[] = [];
      for (let at = 0; at < bytes.length; at += size) {
        parts.push(bytes.subarray(at, at + size));
      }
      await expect(
        Xml.readRoot(parts, { maxStartByte: 1, maxTagBytes: 3 }),
        `chunks of ${size}`
      ).rejects.toBeInstanceOf(XmlLimitExceededError);
      expect((await Xml.readRoot(parts, { maxStartByte: 1, maxTagBytes: 7 })).name).toBe("😀");
    }
  });

  it("refuses a source whose iterator has no next(), and takes next() once", async () => {
    for (const iterator of [{}, undefined, 7]) {
      const source = {
        [Symbol.asyncIterator]: () => iterator
      } as unknown as AsyncIterable<Uint8Array>;
      const err = await Xml.readRoot(source).catch(e => e);
      expect(err).toBeInstanceOf(XmlArgumentError);
      expect(err.field).toBe("source");
    }
    let lookups = 0;
    const data = [encoder.encode("<R>"), encoder.encode("<T/></R>")];
    const source = {
      [Symbol.iterator]: () => {
        let index = 0;
        return {
          get next() {
            lookups++;
            return lookups === 1
              ? () =>
                  index < data.length
                    ? { done: false, value: data[index++] }
                    : { done: true, value: undefined }
              : undefined;
          }
        };
      }
    } as unknown as Iterable<Uint8Array>;
    const { records } = await Xml.readRecords(source, { tags: ["T"], onRecord: () => {} });
    expect([records, lookups]).toEqual([1, 1]);
  });

  it("reads the options it takes as the check does, and nothing else", async () => {
    const bytes = encoder.encode("<R><T/></R>");
    const inherited = Object.create({ tags: ["T"], onRecord: () => {} }) as XmlReadRecordsOptions;
    expect((await Xml.readRecords(bytes, inherited)).records).toBe(1);
    const hidden = Object.defineProperties({} as XmlReadRecordsOptions, {
      tags: { value: ["T"] },
      onRecord: { value: () => {} }
    });
    expect((await Xml.readRecords(bytes, hidden)).records).toBe(1);
    const array = Object.assign([], { tags: ["T"], onRecord: () => {} });
    await expect(
      Xml.readRecords(bytes, array as unknown as XmlReadRecordsOptions)
    ).rejects.toBeInstanceOf(XmlArgumentError);
    let unrelated = 0;
    const options = {
      tags: ["T"],
      onRecord: () => {},
      get unrelated(): never {
        unrelated++;
        throw new Error("not an option");
      }
    };
    expect((await Xml.readRecords(bytes, options)).records).toBe(1);
    expect(unrelated).toBe(0);
  });

  it("stops converting a record as soon as isArray fires the signal", async () => {
    const controller = new AbortController();
    const asked: string[] = [];
    await expect(
      Xml.readRecords(encoder.encode("<R><T><A/><B/><C/></T></R>"), {
        tags: ["T"],
        signal: controller.signal,
        isArray: name => {
          asked.push(name);
          if (name === "A") {
            controller.abort();
          }
          return false;
        },
        onRecord: () => {}
      })
    ).rejects.toBeInstanceOf(AbortError);
    expect(asked).toEqual(["A"]);
  });

  it("decides the root once the whole budget has arrived, a CR before a split character included", async () => {
    // `<` CR, then three bytes of a four-byte character, and nothing more from the source.
    const stalled = (bytes: Uint8Array): AsyncIterable<Uint8Array> => ({
      [Symbol.asyncIterator]: () => {
        let pulls = 0;
        return {
          next: () =>
            pulls++ === 0
              ? Promise.resolve({ done: false, value: bytes })
              : new Promise<IteratorResult<Uint8Array>>(() => {})
        };
      }
    });
    const options = { maxStartByte: 1, maxTagBytes: 1 };
    await expect(
      Xml.readRoot(stalled(new Uint8Array([0x3c, 0x0d, 0xf0, 0x9f, 0x98])), options)
    ).rejects.toBeInstanceOf(XmlParseError);
    await expect(
      Xml.readRoot(stalled(new Uint8Array([0x20, 0x0d, 0xf0, 0x9f, 0x98])), options)
    ).rejects.toBeInstanceOf(XmlLimitExceededError);
  });

  it("reports text before the root where it begins, however the file is chunked", async () => {
    const bytes = encoder.encode("x          <R/>");
    for (const size of [1, bytes.length]) {
      const parts: Uint8Array[] = [];
      for (let at = 0; at < bytes.length; at += size) {
        parts.push(bytes.subarray(at, at + size));
      }
      const err = await Xml.readRoot(parts, { maxStartByte: 4 }).catch(e => e);
      expect(err, `chunks of ${size}`).toBeInstanceOf(XmlParseError);
      expect(err).not.toBeInstanceOf(XmlLimitExceededError);
      expect(err.byte).toBe(0);
    }
  });

  it("takes iterators and results that are functions, as the protocol does", async () => {
    const bytes = encoder.encode("<R><T/></R>");
    const source = {
      [Symbol.iterator]: () => {
        let done = false;
        return Object.assign(() => {}, {
          next: () => {
            if (done) {
              return { done: true, value: undefined };
            }
            done = true;
            return Object.assign(() => {}, { done: false, value: bytes });
          }
        });
      }
    } as unknown as Iterable<Uint8Array>;
    expect((await Xml.readRecords(source, { tags: ["T"], onRecord: () => {} })).records).toBe(1);
  });

  it("readRoot uses the options it checked, read once", async () => {
    let reads = 0;
    const options = {
      get maxStartByte(): number {
        return ++reads === 1 ? 10 : -1;
      }
    };
    expect((await Xml.readRoot(encoder.encode("<R/>"), options)).name).toBe("R");
    expect(reads).toBe(1);
  });

  it("waits for the then a callback's thenable had when it was read, read once", async () => {
    let reads = 0;
    let waited = 0;
    const thenable = {
      get then() {
        return ++reads === 1
          ? (resolve: () => void) => {
              waited++;
              resolve();
            }
          : undefined;
      }
    };
    await Xml.readRecords(encoder.encode("<R><T/></R>"), {
      tags: ["T"],
      onRecord: () => thenable as never
    });
    expect([reads, waited]).toEqual([1, 1]);
  });

  it("uses the tags it checked, read once", async () => {
    let reads = 0;
    const tags: string[] = [];
    Object.defineProperty(tags, "0", {
      enumerable: true,
      get: () => (++reads === 1 ? "T" : "U")
    });
    const names: string[] = [];
    await Xml.readRecords(encoder.encode("<R><T/><U/></R>"), {
      tags,
      onRecord: record => void names.push(record.name)
    });
    expect([reads, names]).toEqual([1, ["T"]]);
  });

  it("opens a source with the method it checked, and takes the iterator protocol as for-await does", async () => {
    let reads = 0;
    const changing = {
      get [Symbol.iterator]() {
        return ++reads === 1
          ? function* (): Generator<Uint8Array> {
              yield encoder.encode("<R/>");
            }
          : undefined;
      }
    } as unknown as Iterable<Uint8Array>;
    expect((await Xml.readRoot(changing)).name).toBe("R");
    expect(reads).toBe(1);
    // A function can be an iterable.
    const callable = Object.assign(() => {}, {
      *[Symbol.iterator](): Generator<Uint8Array> {
        yield encoder.encode("<R/>");
      }
    }) as unknown as Iterable<Uint8Array>;
    expect((await Xml.readRoot(callable)).name).toBe("R");
    // An async method that is there and is no function is refused, not passed over.
    const wrongAsync = {
      [Symbol.asyncIterator]: 4,
      *[Symbol.iterator](): Generator<Uint8Array> {
        yield encoder.encode("<R/>");
      }
    } as unknown as Iterable<Uint8Array>;
    await expect(Xml.readRoot(wrongAsync)).rejects.toMatchObject({
      field: "source[Symbol.asyncIterator]"
    });
    // return() answers with an iterator result.
    const badReturn = {
      [Symbol.iterator]: () => ({
        next: () => ({ done: false, value: encoder.encode("<R/>") }),
        return: () => 42
      })
    } as unknown as Iterable<Uint8Array>;
    await expect(Xml.readRoot(badReturn)).rejects.toBeInstanceOf(XmlArgumentError);
  });

  it("lets readRoot read a root whose entities expand past the default limits, when asked", async () => {
    const xml = `<!DOCTYPE R [<!ENTITY e "${"x".repeat(60000)}">]><R a="${"&e;".repeat(167)}"><T/></R>`;
    const bytes = encoder.encode(xml);
    await expect(Xml.readRoot(bytes)).rejects.toMatchObject({ limit: "entityExpansionLength" });
    const root = await Xml.readRoot(bytes, { maxEntityExpansionLength: 0 });
    expect(root.attributes.a.length).toBe(10_020_000);
  });

  it("reads a range with a large context as fast as with none", async () => {
    // Each boundary a range checks builds a parser from the context: its entities are shared, not copied.
    const entities: Record<string, string> = {};
    for (let i = 0; i < 2000; i++) {
      entities[`e${i}`] = `v${i}`;
    }
    const bytes = encoder.encode(`<R><T>${"x".repeat(20000)}</T><T>&e7;</T></R>`);
    const objects: unknown[] = [];
    await Xml.readRecords(bytes.subarray(3), {
      tags: ["T"],
      offset: 3,
      size: 32,
      context: { root: "R", entities },
      onRecord: record => void objects.push(record.object)
    });
    expect(objects).toHaveLength(1);
    const later: unknown[] = [];
    const second = `<R><T>${"x".repeat(20000)}</T>`.length;
    await Xml.readRecords(bytes.subarray(second), {
      tags: ["T"],
      offset: second,
      context: { root: "R", entities },
      onRecord: record => void later.push(record.object)
    });
    expect(later).toEqual([{ T: "v7" }]);
  });

  it("calls no callback once the signal has fired while a callback's promise was pending", async () => {
    for (let delay = 0; delay < 8; delay++) {
      const controller = new AbortController();
      const fire = (left: number): void => {
        if (left === 0) {
          controller.abort();
        } else {
          queueMicrotask(() => fire(left - 1));
        }
      };
      let calls = 0;
      let afterAbort = 0;
      await Xml.readRecords(encoder.encode("<R><T/><T/></R>"), {
        tags: ["T"],
        signal: controller.signal,
        onRecord: () => {
          calls++;
          if (controller.signal.aborted) {
            afterAbort++;
          }
          return calls === 1 ? Promise.resolve().then(() => fire(delay)) : undefined;
        }
      }).catch(() => undefined);
      expect(afterAbort, `abort ${delay} microtasks on`).toBe(0);
    }
  });

  it("asks the source to close even when the signal has fired", async () => {
    const controller = new AbortController();
    let returned = 0;
    const source: AsyncIterable<Uint8Array> = {
      [Symbol.asyncIterator]: () => ({
        next: () => {
          queueMicrotask(() => controller.abort());
          return Promise.resolve({ done: false, value: encoder.encode("<R/>") });
        },
        return: () => {
          returned++;
          return Promise.resolve({ done: true, value: undefined });
        }
      })
    };
    await Xml.readRoot(source, { signal: controller.signal }).catch(() => undefined);
    expect(returned).toBe(1);
  });

  it("reports a root past its limits without waiting for a source that never closes", async () => {
    const never: AsyncIterable<Uint8Array> = {
      [Symbol.asyncIterator]: () => ({
        next: () =>
          Promise.resolve({ done: false, value: encoder.encode('<R attribute="123456"/>') }),
        return: () => new Promise<IteratorResult<Uint8Array>>(() => {})
      })
    };
    await expect(Xml.readRoot(never, { maxTagBytes: 4, maxStartByte: 30 })).rejects.toBeInstanceOf(
      XmlLimitExceededError
    );
  });

  it("passes out what a callback's then getter throws, as it is", async () => {
    const own = new XmlParseError("the caller's own");
    const thenable = {
      get then(): never {
        throw own;
      }
    };
    for (const offset of [0, 3]) {
      const err = await Xml.readRecords(encoder.encode("<R><T/></R>").subarray(offset), {
        tags: ["T"],
        offset,
        context: { root: "R", entities: {} },
        onRecord: () => thenable as never
      }).catch(e => e);
      expect(err, `offset ${offset}`).toBe(own);
    }
  });

  it("does not let a source that never finishes closing keep an abort from being reported", async () => {
    const controller = new AbortController();
    const slow: AsyncIterable<Uint8Array> = {
      [Symbol.asyncIterator]: () => ({
        next: async () => ({ done: false, value: encoder.encode("<R><T/>") }),
        return: () => new Promise<IteratorResult<Uint8Array>>(() => {})
      })
    };
    const err = await Xml.readRecords(slow, {
      tags: ["T"],
      signal: controller.signal,
      onRecord: () => controller.abort()
    }).catch(e => e);
    expect(err).toBeInstanceOf(AbortError);
  });

  it("reports a source that fails to close when the read itself succeeded", async () => {
    const source: Iterable<Uint8Array> = {
      [Symbol.iterator]: () => ({
        next: () => ({ done: false, value: encoder.encode("<d><r/></d>") }),
        return: () => {
          throw new Error("cannot release");
        }
      })
    };
    await expect(Xml.readRoot(source)).rejects.toThrow("cannot release");
    // A read that fails anyway reports why it failed, not the release.
    const failing: Iterable<Uint8Array> = {
      [Symbol.iterator]: () => ({
        next: () => ({ done: false, value: encoder.encode("<1/>") }),
        return: () => {
          throw new Error("cannot release");
        }
      })
    };
    await expect(Xml.readRoot(failing)).rejects.toBeInstanceOf(XmlParseError);
  });
});

describe("Xml.readRecords — cancelling and arguments", () => {
  it("stops, closes the source and calls nothing more once the signal fires", async () => {
    const controller = new AbortController();
    let closed = false;
    let calls = 0;
    async function* source(): AsyncGenerator<Uint8Array> {
      try {
        yield encoder.encode("<d>");
        for (let i = 0; ; i++) {
          yield encoder.encode(`<r>${i}</r>`);
        }
      } finally {
        closed = true;
      }
    }
    const reason = new Error("stop");
    const read = Xml.readRecords(source(), {
      tags: ["r"],
      signal: controller.signal,
      onRecord: () => {
        calls++;
        if (calls === 3) {
          controller.abort(reason);
        }
      }
    });
    const err = await read.catch(e => e);
    expect(err).toBeInstanceOf(AbortError);
    expect(isAbortError(err)).toBe(true);
    expect(err.cause).toBe(reason);
    expect(closed).toBe(true);
    const after = calls;
    await new Promise(resolve => setTimeout(resolve, 10));
    expect(calls).toBe(after);
  });

  it("interrupts a source that is waiting for its next chunk, and does not wait for it to answer", async () => {
    for (const reader of ["records", "root"] as const) {
      const controller = new AbortController();
      let released = false;
      async function* quiet(): AsyncGenerator<Uint8Array> {
        try {
          yield encoder.encode("<d a='1"); // half a start tag: neither reader can finish yet
          await new Promise(() => {}); // a source that has gone quiet and never answers
        } finally {
          released = true;
        }
      }
      const pending =
        reader === "records"
          ? Xml.readRecords(quiet(), { tags: ["r"], signal: controller.signal, onRecord: () => {} })
          : Xml.readRoot(quiet(), { signal: controller.signal });
      setTimeout(() => controller.abort(new Error("late")), 10);
      const err = await pending.catch(e => e);
      expect(err, reader).toBeInstanceOf(AbortError);
      expect(err.cause.message).toBe("late");
      // The generator cannot run its `finally` while its `next()` is pending: it was asked, not awaited.
      expect(released).toBe(false);
    }
  });

  it("does not miss an abort that the source's next() fires itself", async () => {
    const controller = new AbortController();
    let returned = false;
    const source: AsyncIterable<Uint8Array> = {
      [Symbol.asyncIterator]: () => ({
        next: () => {
          controller.abort(new Error("from next"));
          return new Promise<IteratorResult<Uint8Array>>(() => {});
        },
        return: async () => {
          returned = true;
          return { done: true, value: undefined };
        }
      })
    };
    for (const pending of [
      Xml.readRoot(source, { signal: controller.signal }),
      Xml.readRecords(source, { tags: ["r"], signal: controller.signal, onRecord: () => {} })
    ]) {
      const err = await pending.catch(e => e);
      expect(err).toBeInstanceOf(AbortError);
    }
    // The source is asked to stop, though not waited for.
    expect(returned).toBe(true);
  });

  it("stops waiting for a callback's promise when the signal fires", async () => {
    const controller = new AbortController();
    let closed = false;
    async function* source(): AsyncGenerator<Uint8Array> {
      try {
        yield encoder.encode("<d><r/><r/></d>");
      } finally {
        closed = true;
      }
    }
    let calls = 0;
    const err = await Xml.readRecords(source(), {
      tags: ["r"],
      signal: controller.signal,
      onRecord: () => {
        calls++;
        controller.abort(new Error("stop"));
        return new Promise<void>(() => {}); // never settles
      }
    }).catch(e => e);
    expect(err).toBeInstanceOf(AbortError);
    expect(calls).toBe(1);
    expect(closed).toBe(true);
  });

  it("calls no callback after one that fires the signal", async () => {
    const controller = new AbortController();
    let calls = 0;
    const err = await Xml.readRecords(encoder.encode(`<d>${"<r/>".repeat(1000)}</d>`), {
      tags: ["r"],
      signal: controller.signal,
      onRecord: () => {
        if (++calls === 3) {
          controller.abort();
        }
      }
    }).catch(e => e);
    expect(err).toBeInstanceOf(AbortError);
    expect(calls).toBe(3);
  });

  it("rejects at once when the signal has already fired", async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(
      Xml.readRecords(encoder.encode("<d/>"), {
        tags: ["r"],
        signal: controller.signal,
        onRecord: () => {}
      })
    ).rejects.toBeInstanceOf(AbortError);
  });

  it("closes the source when it stops early", async () => {
    let closed = false;
    function* source(): Generator<Uint8Array> {
      try {
        yield encoder.encode("<d><r>1</r>");
        yield encoder.encode("<r>2</r>".repeat(100));
        yield encoder.encode("</d>");
      } finally {
        closed = true;
      }
    }
    const result = await Xml.readRecords(source(), { tags: ["r"], size: 5, onRecord: () => {} });
    expect(result.records).toBe(1);
    expect(closed).toBe(true);
  });

  it.each([
    [{ tags: [], onRecord: () => {} }, "options.tags"],
    [{ tags: ["1x"], onRecord: () => {} }, "options.tags"],
    [{ tags: ["r"] }, "options.onRecord"],
    [{ tags: ["r"], onRecord: () => {}, offset: -1 }, "options.offset"],
    [{ tags: ["r"], onRecord: () => {}, size: 0 }, "options.size"],
    [{ tags: ["r"], onRecord: () => {}, offset: 10 }, "options.context"],
    [{ tags: ["r"], onRecord: () => {}, offset: 2 ** 53 }, "options.offset"],
    [
      {
        tags: ["r"],
        onRecord: () => {},
        offset: 2 ** 52,
        size: 2 ** 52,
        context: { root: "d", entities: {} }
      },
      "options.size"
    ],
    [
      {
        tags: ["r"],
        onRecord: () => {},
        offset: 1,
        context: { root: "d", entities: { e: "\u0001" } }
      },
      "options.context"
    ],
    [
      {
        tags: ["r"],
        onRecord: () => {},
        offset: 1,
        maxContextBytes: 30,
        context: { root: "d", entities: { e: "x".repeat(10) } }
      },
      "options.context"
    ],
    [
      {
        tags: ["r"],
        onRecord: () => {},
        offset: 10,
        context: { root: "d", entities: { "1": "x" } }
      },
      "options.context"
    ]
  ])("refuses %j", async (options, field) => {
    const err = await Xml.readRecords(new Uint8Array(0), options as XmlReadRecordsOptions).catch(
      e => e
    );
    expect(err).toBeInstanceOf(XmlArgumentError);
    expect(err.field.startsWith(field)).toBe(true);
    expect(err.message).toMatch(/^Xml\.readRecords: /);
  });

  it("refuses a source that is not bytes", async () => {
    const err = await Xml.readRecords("<d/>" as never, { tags: ["r"], onRecord: () => {} }).catch(
      e => e
    );
    expect(err).toBeInstanceOf(XmlArgumentError);
    expect(err.field).toBe("source");
    const chunkErr = await Xml.readRecords(["<d/>"] as never, {
      tags: ["r"],
      onRecord: () => {}
    }).catch(e => e);
    expect(chunkErr).toBeInstanceOf(XmlArgumentError);
  });
});

describe("Xml.readRecords — what a chunk reports past its end, and what it holds", () => {
  const context = { root: "R", entities: {} };

  it("finds the text after a self-closing tag that the chunk's end falls inside", async () => {
    const bytes = encoder.encode("<R><T/>xxx</R>");
    const { structure } = await read(bytes, { tags: ["T"], offset: 3, size: 1, context });
    expect(structure.follow).toBe(7);
  });

  it("joins a chunk that stopped inside a comment or processing instruction past its end", async () => {
    for (const xml of ["<R></R><!--x-->", "<R></R><?p?>"]) {
      const bytes = encoder.encode(xml);
      for (const sourceChunk of [1, 2, 3, 65536]) {
        const first = await read(bytes, { tags: ["T"], offset: 0, size: 4 }, sourceChunk);
        const rest = await read(bytes, { tags: ["T"], offset: 4, context }, sourceChunk);
        expect(first.structure.follow, `${xml} / ${sourceChunk}`).toBe(null);
        expect(() => Xml.verifyPartition([first, rest])).not.toThrow();
      }
    }
  });

  it("reads a chunk that is all text no further than its end, and still joins it", async () => {
    // Ranges inside a long comment after the root: each, read as text, has nothing to place past its
    // end — the range before the comment reads through all of it.
    const comment = "x".repeat(4096);
    const xml = `<R><T/></R><!--${comment}-->\n`;
    const bytes = encoder.encode(xml);
    const size = 64;
    const results: Awaited<ReturnType<typeof read>>[] = [];
    let readPast = 0;
    for (let offset = 0; offset < bytes.length; offset += size) {
      const result = await read(bytes, {
        tags: ["T"],
        offset,
        size,
        context: offset === 0 ? undefined : context
      });
      readPast += result.endByte - Math.min(bytes.length, offset + size);
      results.push(result);
    }
    expect(() => Xml.verifyPartition(results)).not.toThrow();
    // The range before the comment reads to its end; the rest no further than their own.
    expect(readPast).toBeLessThan(xml.length + 2 * size);
    // Text there is still found: by a range's own reading, or by the join.
    const bad = encoder.encode(`<R><T/></R><!--${comment}--> text`);
    const outcome = await (async () => {
      const badResults: Awaited<ReturnType<typeof read>>[] = [];
      for (let offset = 0; offset < bad.length; offset += size) {
        badResults.push(
          await read(bad, {
            tags: ["T"],
            offset,
            size,
            context: offset === 0 ? undefined : context
          })
        );
      }
      Xml.verifyPartition(badResults);
    })().catch((err: unknown) => err);
    expect(outcome).toBeInstanceOf(XmlParseError);
  });

  it("joins a run of chunks read as text in time that grows with the run, not its square", async () => {
    const bytes = encoder.encode(`<R/><!--${"x".repeat(2000)}-->`);
    const results: Awaited<ReturnType<typeof read>>[] = [];
    for (let offset = 0; offset < bytes.length; offset++) {
      results.push(
        await read(bytes, {
          tags: ["T"],
          offset,
          size: 1,
          context: offset === 0 ? undefined : context
        })
      );
    }
    let reads = 0;
    const counted = results.map(result => ({
      get structure() {
        reads++;
        return result.structure;
      }
    }));
    expect(() => Xml.verifyPartition(counted)).not.toThrow();
    // Each result is read once.
    expect(reads).toBe(results.length);
  });

  it("refuses ranges past the end of the file, and a last range that stops short of it", async () => {
    const bytes = encoder.encode("<R/>");
    const first = await read(bytes, { tags: ["T"], offset: 0, size: 10 });
    const ghost = await read(bytes, { tags: ["T"], offset: 10, size: 10, context });
    expect(first.structure.eof).toBe(4);
    expect(() => Xml.verifyPartition([first])).not.toThrow();
    expect(() => Xml.verifyPartition([first, ghost])).toThrow(XmlArgumentError);
    // A last range whose read reached the end, but whose own range did not.
    expect(() =>
      Xml.verifyPartition([
        { structure: { start: 0, end: 2, final: true, eof: 4, steps: [["e", 0]], open: [] } }
      ])
    ).toThrow(XmlArgumentError);
  });

  it("joins what it checked: a structure read twice is read once", () => {
    let reads = 0;
    const good = { start: 0, end: null, final: true, eof: 4, steps: [["e", 0]], open: [] };
    const bad = {
      ...good,
      steps: [
        ["e", 0],
        ["t", 2]
      ]
    };
    const result = {
      get structure() {
        return reads++ === 0 ? bad : good;
      }
    };
    expect(() => Xml.verifyPartition([result as never])).toThrow(XmlParseError);
    expect(reads).toBe(1);
  });

  it("names the record when the input ends inside its start tag", async () => {
    for (const xml of ['<R><T a="oops', '<R><T a="x"/', '<R><T a="&am']) {
      await expect(read(encoder.encode(xml), { tags: ["T"] }), xml).rejects.toMatchObject({
        name: "XmlTruncatedError",
        recordStartByte: 3
      });
    }
    // A name not yet whole, or another element's, is no record's.
    for (const xml of ["<R><T", '<R><U a="1"']) {
      await expect(read(encoder.encode(xml), { tags: ["T"] })).rejects.toMatchObject({
        recordStartByte: undefined
      });
    }
  });

  it("holds no markup outside a record for rawXml, and all of a record's start tag", async () => {
    const xml = `<R><!--${"x".repeat(70000)}--><T a="${"y".repeat(5000)}"/></R>`;
    const bytes = encoder.encode(xml);
    for (const sourceChunk of [7, 4096]) {
      const { records } = await read(bytes, { tags: ["T"], rawXml: true }, sourceChunk);
      expect(records[0].xml).toBe(`<T a="${"y".repeat(5000)}"/>`);
    }
  });
});

describe("Xml.readRecords — a source that fails as it is opened", () => {
  it("is refused for a return that is not a method", async () => {
    const source = {
      [Symbol.iterator]: () => ({
        next: () => ({ done: false, value: encoder.encode("<R/>") }),
        return: 7
      })
    };
    await expect(Xml.readRoot(source as never)).rejects.toMatchObject({
      name: "XmlArgumentError",
      field: "source"
    });
  });

  it("is released when its next cannot be read, and not opened once the signal has fired", async () => {
    let opened = 0;
    let released = 0;
    const source = {
      [Symbol.iterator]: () => {
        opened++;
        return {
          get next(): never {
            throw new Error("next getter");
          },
          return: () => {
            released++;
            return { done: true, value: undefined };
          }
        };
      }
    };
    await expect(Xml.readRoot(source as never)).rejects.toThrow("next getter");
    expect([opened, released]).toEqual([1, 1]);
    const controller = new AbortController();
    controller.abort();
    await expect(
      Xml.readRoot(source as never, { signal: controller.signal })
    ).rejects.toMatchObject({
      name: "AbortError"
    });
    expect(opened).toBe(1);
  });
});
