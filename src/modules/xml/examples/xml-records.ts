/**
 * Example: reading a large XML file's records, whole or in independent byte ranges.
 *
 * Covers:
 * - Xml.readRoot: the root element and the prolog's context, reading no further than the root's `>`
 * - Xml.readRecords: each record as an object, in one pass over the bytes
 * - Byte ranges read independently (as separate workers would), staged, and published only when all
 *   succeed and Xml.verifyPartition joins them into one whole document
 * - XmlChunkAmbiguityError: a range that cannot be placed — fall back to reading the file whole
 * - XmlTruncatedError: a file cut short publishes nothing, though some of its ranges read without error
 *
 * Run: pnpm example --filter xml-records
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Xml, XmlChunkAmbiguityError, XmlTruncatedError } from "../index";
import type { XmlReadRecordsResult, XmlRecord } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/xml-examples"
);
fs.mkdirSync(outDir, { recursive: true });
const file = process.argv[2] ?? path.join(outDir, "sales-report.xml");

// A sales report with a DOCTYPE entity and a few thousand records.
let xml =
  '<?xml version="1.0" encoding="UTF-8"?>\n' +
  '<!DOCTYPE SalesReport [ <!ENTITY label "Ünïversal &amp; Co"> ]>\n' +
  '<SalesReport Period="2026-10">\n';
for (let i = 0; i < 5000; i++) {
  xml += `  <Sale Id="${i}">\n    <Title>Låt ${i}</Title>\n    <Label>&label;</Label>\n    <Count>${(i * 37) % 1000}</Count>\n  </Sale>\n`;
}
xml += "</SalesReport>\n";
fs.writeFileSync(file, xml);

const bytes = new Uint8Array(fs.readFileSync(file));

/** The file's bytes from `start` on, in 64 KB chunks, as a file stream would give them. */
function* from(start: number): Generator<Uint8Array> {
  for (let at = start; at < bytes.length; at += 65536) {
    yield bytes.subarray(at, at + 65536);
  }
}

// 1. The root element and the context later ranges need.
const root = await Xml.readRoot(from(0));
console.log(`root <${root.name}>`, { ...root.attributes }, JSON.stringify(root.context));

// 2. The whole file, one record at a time.
let total = 0;
await Xml.readRecords(from(0), {
  tags: ["Sale"],
  onRecord: () => {
    total++;
  }
});
console.log(`whole file: ${total} records`);

// 3. The same file as independent ranges of 32 KB — each could run in its own worker.
//
// A range that reads without an error says nothing about the others, nor about the elements open
// around it: the file may be cut, or not one document. So each range's records are staged, and the
// batch is published only when every range of this one version of the file has succeeded and
// Xml.verifyPartition has joined their structures — otherwise it is discarded whole.
const size = 32 * 1024;

/** Read `source` (one fixed version of the file) as ranges; publish every record, or none. */
async function readInRanges(
  source: Uint8Array,
  publish: (records: readonly XmlRecord[]) => void
): Promise<void> {
  const context = (await Xml.readRoot(source)).context;
  const staged: XmlRecord[][] = [];
  const ranges: Array<Promise<XmlReadRecordsResult>> = [];
  for (let offset = 0; offset < source.length; offset += size) {
    const own: XmlRecord[] = (staged[staged.length] = []);
    ranges.push(
      Xml.readRecords(source.subarray(offset), {
        tags: ["Sale"],
        offset,
        size,
        context,
        onRecord: record => {
          own.push(record);
        }
      })
    );
  }
  const outcomes = await Promise.allSettled(ranges);
  try {
    const failure = outcomes.find(outcome => outcome.status === "rejected");
    if (failure !== undefined) {
      throw (failure as PromiseRejectedResult).reason;
    }
    // The results in file order — as a worker would send them back, as JSON.
    Xml.verifyPartition(
      outcomes.map(outcome => (outcome as PromiseFulfilledResult<XmlReadRecordsResult>).value)
    );
  } catch (err) {
    // Discard every staged record, those of the ranges that succeeded included.
    staged.length = 0;
    throw err;
  }
  publish(staged.flat());
}

let published: readonly XmlRecord[] = [];
try {
  await readInRanges(bytes, records => {
    published = records;
  });
} catch (err) {
  if (!(err instanceof XmlChunkAmbiguityError)) {
    throw err;
  }
  // A range's start could not be placed: read the file in one pass instead (or in larger ranges).
  console.log(`a range is ambiguous at byte ${err.atByte}; reading the file whole`);
  const whole: XmlRecord[] = [];
  await Xml.readRecords(bytes, { tags: ["Sale"], onRecord: record => void whole.push(record) });
  published = whole;
}
console.log(`ranges of ${size} bytes: ${published.length} records`);
console.log("first:", JSON.stringify(published[0].object));
if (published.length !== total) {
  throw new Error("the ranges should give the whole file's records");
}

// 4. A file cut short: some range fails, or the join finds the root still open — so nothing of it
//    is published, not even the ranges that read without an error before the cut.
let publishedFromCut = false;
const cut = await readInRanges(bytes.subarray(0, bytes.length - 1000), () => {
  publishedFromCut = true;
}).catch((err: unknown) => err);
console.log(
  `cut file: ${cut instanceof XmlTruncatedError ? "truncated" : String(cut)}, published: ${publishedFromCut}`
);
if (!(cut instanceof XmlTruncatedError) || publishedFromCut) {
  throw new Error("a cut file should publish nothing");
}
