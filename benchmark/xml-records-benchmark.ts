/**
 * `Xml.readRecords` and the byte-fed SAX parser, on a generated music-usage report shaped like the
 * ones streaming services deliver (a DDEX-style DSR: a header, then a long run of formatted records
 * holding nested elements, attributes, entity references and non-ASCII text).
 *
 * Like the other benchmarks, it reports timings, it does not enforce them: the numbers are
 * observations on one machine. It measures time only — not memory, peak or otherwise. Four questions:
 *
 * 1. **Records** — one pass that scans and builds every record's object, against `Xml.parseToObject`
 *    of the whole file (the same objects, held at once).
 * 2. **Positions** — what byte positions cost: the parser given bytes, which tracks `startByte` /
 *    `endByte` for every event, against the same parser given text decoded in advance, and text
 *    decoded chunk by chunk as it is fed. Every handler does the same work (counts the event), so the
 *    difference is the parser's.
 * 3. **Chunks** — the same file read as independent byte ranges and joined with
 *    `Xml.verifyPartition`, as a caller publishing them must: total time, and how far past its range
 *    each read went to place the next boundary.
 * 4. **Source chunking** — whether the size of the source's chunks moves any of it.
 *
 * Before anything is timed, a correctness gate checks that the records `Xml.readRecords` reads —
 * their order, byte ranges and objects — are the whole file's, read in one pass and in every range
 * split measured here, and that each split's results join; a mismatch throws and the benchmark exits
 * without numbers.
 *
 * Timed cases run interleaved: each is warmed up, then every round runs all of them once, in an order
 * that rotates between rounds, and each reports its median. That keeps a fixed order's JIT and GC
 * effects off any one case.
 *
 * Run: pnpm benchmark:xml-records   (XML_BENCHMARK_MB sets the size in MiB, default 64;
 *                                    XML_BENCHMARK_RUNS the rounds, default 5)
 */
import process from "node:process";

import { Xml } from "../src/modules/xml/index.ts";
import type { XmlReadRecordsResult } from "../src/modules/xml/index.ts";
import { SaxParser } from "../src/modules/xml/sax.ts";

const TARGET_MB = Number(process.env.XML_BENCHMARK_MB ?? 64);
const RUNS = Number(process.env.XML_BENCHMARK_RUNS ?? 5);
const WARMUPS = 2;
const SOURCE_CHUNK = 64 * 1024;
const RANGE_SIZES_MB = [1, 8];
const SOURCE_SIZES = [1024, 7 * 1024, 64 * 1024, 1024 * 1024];
const TAGS = ["SalesRecord"];

function generate(targetBytes: number): Uint8Array {
  const parts: string[] = [
    '<?xml version="1.0" encoding="UTF-8"?>\n',
    '<!DOCTYPE DSR [ <!ENTITY label "Ünïversal Records &amp; Co"> ]>\n',
    '<DSR xmlns="http://ddex.net/xml/dsr" ReportId="R-2026-11" Territory="WW">\n',
    "  <Header>\n    <Sender>Spotify AB</Sender>\n    <Label>&label;</Label>\n    <Period>2026-10</Period>\n  </Header>\n"
  ];
  const encoder = new TextEncoder();
  let size = encoder.encode(parts.join("")).length;
  for (let i = 0; size < targetBytes; i++) {
    const record =
      `  <SalesRecord Id="${i}" Type="Stream">\n` +
      `    <ISRC>SE${String(i).padStart(10, "0")}</ISRC>\n` +
      `    <Title Lang="sv">Låt nummer ${i} — “live”</Title>\n` +
      `    <Artists>\n      <Artist Role="Main">Artist ${i % 997}</Artist>\n      <Artist Role="Featured">Gäst ${i % 31}</Artist>\n    </Artists>\n` +
      `    <Label>Ünïversal Records &amp; Co</Label>\n` +
      `    <Usage Territory="${["SE", "DE", "US", "JP"][i % 4]}" Count="${(i * 7919) % 100000}" Revenue="${((i * 31) % 10000) / 100}"/>\n` +
      `    <Note><![CDATA[explicit; remastered]]></Note>\n` +
      `  </SalesRecord>\n`;
    parts.push(record);
    size += encoder.encode(record).length;
  }
  parts.push("</DSR>\n");
  return encoder.encode(parts.join(""));
}

function chunks(bytes: Uint8Array, size: number): Uint8Array[] {
  const out: Uint8Array[] = [];
  for (let at = 0; at < bytes.length; at += size) {
    out.push(bytes.subarray(at, at + size));
  }
  return out;
}

/** Views of `bytes` from `from` on, made only as the reader asks for them. */
function* lazyChunks(bytes: Uint8Array, from: number, size: number): Generator<Uint8Array> {
  for (let at = from; at < bytes.length; at += size) {
    yield bytes.subarray(at, at + size);
  }
}

const bytes = generate(TARGET_MB * 1024 * 1024);
const text = new TextDecoder().decode(bytes);
const sourceChunks = chunks(bytes, SOURCE_CHUNK);
// The text a caller decoding the same chunks would hand the parser: flat strings, one per chunk.
const textChunks: string[] = [];
{
  const decoder = new TextDecoder();
  for (const chunk of sourceChunks) {
    textChunks.push(decoder.decode(chunk, { stream: true }));
  }
  const tail = decoder.decode();
  if (tail) {
    textChunks.push(tail);
  }
}
const { context } = await Xml.readRoot(sourceChunks);

// ---------------------------------------------------------------------------------------------
// Correctness gate
// ---------------------------------------------------------------------------------------------

interface RecordSummary {
  readonly startByte: number;
  readonly endByte: number;
  readonly json: string;
}

function fail(message: string): never {
  throw new Error(`correctness gate: ${message}`);
}

function compareRecords(label: string, actual: RecordSummary[], expected: RecordSummary[]): void {
  if (actual.length !== expected.length) {
    fail(`${label}: ${actual.length} records, expected ${expected.length}`);
  }
  for (let i = 0; i < expected.length; i++) {
    const a = actual[i];
    const e = expected[i];
    if (a.startByte !== e.startByte || a.endByte !== e.endByte) {
      fail(
        `${label}: record ${i} at [${a.startByte}, ${a.endByte}), expected [${e.startByte}, ${e.endByte})`
      );
    }
    if (a.json !== e.json) {
      fail(`${label}: record ${i} object ${a.json}, expected ${e.json}`);
    }
  }
}

/** The whole file read at once: objects from `Xml.parseToObject`, byte ranges from a byte-fed parser. */
function readWhole(): RecordSummary[] {
  const tree = Xml.parseToObject(text, { internalEntities: true }) as {
    DSR: { SalesRecord: unknown[] };
  };
  const objects = tree.DSR.SalesRecord;
  const ranges: Array<[number, number]> = [];
  const parser = new SaxParser({ position: false, internalEntities: true });
  let depth = 0;
  parser.on("opentag", tag => {
    depth++;
    if (depth === 2 && tag.name === "SalesRecord") {
      ranges.push([parser.startByte!, -1]);
    }
  });
  parser.on("closetag", tag => {
    if (depth === 2 && tag.name === "SalesRecord") {
      ranges[ranges.length - 1][1] = parser.endByte!;
    }
    depth--;
  });
  for (const chunk of sourceChunks) {
    parser.write(chunk);
  }
  parser.close();
  if (ranges.length !== objects.length) {
    fail(`whole read: ${ranges.length} byte ranges but ${objects.length} objects`);
  }
  return objects.map((object, i) => ({
    startByte: ranges[i][0],
    endByte: ranges[i][1],
    json: JSON.stringify({ SalesRecord: object })
  }));
}

function summarise(record: { startByte: number; endByte: number; object: unknown }): RecordSummary {
  return {
    startByte: record.startByte,
    endByte: record.endByte,
    json: JSON.stringify(record.object)
  };
}

/** Every range of `rangeSize`, read and staged, then joined as a caller publishing them must. */
async function readRanges(rangeSize: number): Promise<{ records: RecordSummary[]; over: number }> {
  const records: RecordSummary[] = [];
  const results: XmlReadRecordsResult[] = [];
  let over = 0;
  for (let offset = 0; offset < bytes.length; offset += rangeSize) {
    const result = await Xml.readRecords(lazyChunks(bytes, offset, SOURCE_CHUNK), {
      tags: TAGS,
      offset,
      size: rangeSize,
      context,
      onRecord: record => {
        records.push(summarise(record));
      }
    });
    results.push(result);
    over = Math.max(over, result.endByte - Math.min(bytes.length, offset + rangeSize));
  }
  // The batch is the file's only once its structures join: what a publisher checks, so the gate does.
  Xml.verifyPartition(results);
  return { records, over };
}

const expected = readWhole();
{
  const oneRead: RecordSummary[] = [];
  await Xml.readRecords(sourceChunks, {
    tags: TAGS,
    onRecord: record => oneRead.push(summarise(record))
  });
  compareRecords("Xml.readRecords (one pass)", oneRead, expected);
}
const furthestPast = new Map<number, number>();
for (const rangeMb of RANGE_SIZES_MB) {
  const { records, over } = await readRanges(rangeMb * 1024 * 1024);
  compareRecords(`ranges of ${rangeMb} MB`, records, expected);
  furthestPast.set(rangeMb, over);
}
for (const sourceSize of SOURCE_SIZES) {
  const read: RecordSummary[] = [];
  await Xml.readRecords(chunks(bytes, sourceSize), {
    tags: TAGS,
    onRecord: record => read.push(summarise(record))
  });
  compareRecords(`source chunks of ${sourceSize} B`, read, expected);
}

// ---------------------------------------------------------------------------------------------
// Timing
// ---------------------------------------------------------------------------------------------

function scan(input: Iterable<string | Uint8Array>): number {
  const parser = new SaxParser({ position: false, internalEntities: true });
  let n = 0;
  parser.on("opentag", () => {
    n++;
  });
  parser.on("closetag", () => {
    n++;
  });
  for (const chunk of input) {
    parser.write(chunk);
  }
  parser.close();
  return n;
}

function decodeAndScan(): number {
  const parser = new SaxParser({ position: false, internalEntities: true });
  let n = 0;
  parser.on("opentag", () => {
    n++;
  });
  parser.on("closetag", () => {
    n++;
  });
  const decoder = new TextDecoder();
  for (const chunk of sourceChunks) {
    parser.write(decoder.decode(chunk, { stream: true }));
  }
  const tail = decoder.decode();
  if (tail) {
    parser.write(tail);
  }
  parser.close();
  return n;
}

const cases = new Map<string, () => unknown>();
cases.set("records", () => Xml.readRecords(sourceChunks, { tags: TAGS, onRecord: () => {} }));
cases.set("parseToObject", () => Xml.parseToObject(text, { internalEntities: true }));
cases.set("scan text", () => scan(textChunks));
cases.set("scan decode", () => decodeAndScan());
cases.set("scan bytes", () => scan(sourceChunks));
for (const rangeMb of RANGE_SIZES_MB) {
  cases.set(`ranges ${rangeMb}`, async () => {
    const size = rangeMb * 1024 * 1024;
    const results: XmlReadRecordsResult[] = [];
    for (let offset = 0; offset < bytes.length; offset += size) {
      results.push(
        await Xml.readRecords(lazyChunks(bytes, offset, SOURCE_CHUNK), {
          tags: TAGS,
          offset,
          size,
          context,
          onRecord: () => {}
        })
      );
    }
    // Timed with the reads: a batch is not published before it is joined.
    Xml.verifyPartition(results);
  });
}
for (const sourceSize of SOURCE_SIZES) {
  const parts = chunks(bytes, sourceSize);
  cases.set(`source ${sourceSize}`, () =>
    Xml.readRecords(parts, { tags: TAGS, onRecord: () => {} })
  );
}

const names = [...cases.keys()];
for (const name of names) {
  for (let i = 0; i < WARMUPS; i++) {
    await cases.get(name)!();
  }
}
const timings = new Map<string, number[]>(names.map(name => [name, []]));
for (let round = 0; round < RUNS; round++) {
  for (let i = 0; i < names.length; i++) {
    const name = names[(i + round) % names.length];
    const started = performance.now();
    await cases.get(name)!();
    timings.get(name)!.push(performance.now() - started);
  }
}
function median(name: string): number {
  const sorted = [...timings.get(name)!].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)];
}

const mib = (bytes.length / 1024 / 1024).toFixed(1);
const ms = (value: number): string => `${value.toFixed(0)} ms`;
console.log(
  `file: ${mib} MiB, source chunks of ${SOURCE_CHUNK / 1024} KB, median of ${RUNS} interleaved rounds after ${WARMUPS} warmups`
);
console.log(`correctness gate: ${expected.length} records match in every read below\n`);

console.log("1. records");
console.log(`   Xml.readRecords                    ${ms(median("records"))}`);
console.log(`   Xml.parseToObject (whole file)     ${ms(median("parseToObject"))}\n`);

const textScan = median("scan text");
const decodeScan = median("scan decode");
const byteScan = median("scan bytes");
const percent = (value: number, base: number): string =>
  `${(((value - base) / base) * 100).toFixed(1)}%`;
console.log("2. positions");
console.log(`   text, no positions        ${ms(textScan)}`);
console.log(`   decode + text             ${ms(decodeScan)}`);
// Against text decoded as it is fed, which a caller holding bytes has to do anyway, the difference is
// what positions cost; against text decoded in advance, it is that and the decoding too.
console.log(
  `   bytes, startByte/endByte  ${ms(byteScan)}  (${percent(byteScan, decodeScan)} over decode + text, ${percent(byteScan, textScan)} over text decoded in advance)\n`
);

for (const rangeMb of RANGE_SIZES_MB) {
  console.log(
    `3. chunks of ${rangeMb} MB: ${ms(median(`ranges ${rangeMb}`))}; furthest past a range end: ${furthestPast.get(rangeMb)} bytes`
  );
}

console.log("\n4. source chunking");
for (const sourceSize of SOURCE_SIZES) {
  console.log(
    `   ${String(sourceSize / 1024).padStart(5)} KB chunks  ${ms(median(`source ${sourceSize}`))}`
  );
}
