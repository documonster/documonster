/**
 * How long the streaming readers and the Office decryption hold the event loop, on inputs shaped to find the longest
 * piece of work each does in one go: a part stored rather than deflated (nothing in the source waits, so only the
 * reader can yield), one very large element, a large table part.
 *
 * Like the other benchmarks, it reports timings, it does not enforce them: the numbers are observations on one machine.
 * For each case it reports the total time and the longest stretch the event loop could not run a callback — measured
 * by one that reschedules itself with `setImmediate` and records the gap between its runs. The readers give the event
 * loop a turn once they have worked for `YIELD_BUDGET_MS`, so the longest stretch should be that budget plus one step.
 *
 * Before anything is timed, each case's result is checked against the same input read whole (`Workbook.read`,
 * `Document.read`, the original bytes); a mismatch throws and the benchmark exits without numbers. Timed cases run
 * interleaved — every round runs each case once, in an order that rotates between rounds — and each reports its median.
 *
 * Run: pnpm benchmark:event-loop   (EVENT_LOOP_BENCHMARK_RUNS sets the rounds, default 5)
 */
import process from "node:process";

import { decryptPackage, encryptPackage } from "../src/modules/archive/office-crypto/agile.ts";
import { packZipSync, unpackZipSync } from "../src/modules/archive/zip-one-shot.ts";
import { rowGetValues } from "../src/modules/excel/core/row.ts";
import { Stream as ExcelStream, Workbook, Worksheet } from "../src/modules/excel/index.ts";
import { Document, Stream } from "../src/modules/word/index.ts";
import { YIELD_BUDGET_MS } from "../src/utils/yield.ts";

const RUNS = Number(process.env.EVENT_LOOP_BENCHMARK_RUNS ?? 5);
if (!Number.isInteger(RUNS) || RUNS < 1) {
  throw new Error(
    `EVENT_LOOP_BENCHMARK_RUNS must be a positive integer, got ${process.env.EVENT_LOOP_BENCHMARK_RUNS}`
  );
}
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

interface Case {
  readonly name: string;
  readonly run: () => Promise<unknown>;
  readonly check: (result: unknown) => Promise<void> | void;
}

/** A package's parts repacked at `level`, with `first` leading and `replace` substituting parts. */
function repack(
  bytes: Uint8Array,
  level: number,
  { first = [], replace = {} }: { first?: readonly string[]; replace?: Record<string, string> } = {}
): Uint8Array {
  const parts = unpackZipSync(bytes);
  const out = new Map<string, Uint8Array>();
  for (const name of [...first, ...parts.keys()]) {
    const part = parts.get(name);
    if (part && !out.has(name)) {
      out.set(name, name in replace ? new TextEncoder().encode(replace[name]) : part.data);
    }
  }
  return packZipSync(out, { level });
}

function same(label: string, actual: unknown, expected: unknown): void {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`${label}: the streamed result differs from the whole read`);
  }
}

async function streamedRows(bytes: Uint8Array): Promise<unknown[][]> {
  const rows: unknown[][] = [];
  for await (const sheet of new ExcelStream.WorkbookReader(bytes, { styles: "cache" })) {
    for await (const row of sheet) {
      rows.push(rowGetValues(row));
    }
  }
  return rows;
}

async function wholeRows(bytes: Uint8Array): Promise<unknown[][]> {
  const sheet = Workbook.getWorksheets(await Workbook.read(bytes))[0]!;
  // Each row as `rowGetValues` gives a streamed one, so the two are compared index for index.
  return Worksheet.getValues(sheet, { base: 0 }).filter(row => row !== undefined);
}

async function streamedBody(bytes: Uint8Array): Promise<unknown[]> {
  const items: unknown[] = [];
  for await (const item of new Stream.DocumentReader(bytes)) {
    items.push(item);
  }
  return items;
}

async function excelCases(): Promise<Case[]> {
  const articles = Workbook.create();
  const sheet = Workbook.addWorksheet(articles, "Sheet1");
  for (let i = 0; i < 10_000; i++) {
    Worksheet.addRow(sheet, [`ART-${i}`, `Article ${i} for product line ${i % 97}`, i / 7]);
  }
  const prerequisitesFirst = [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/sharedStrings.xml",
    "xl/styles.xml"
  ];
  const written = await Workbook.toBuffer(articles);
  const strings = Workbook.create();
  const stringSheet = Workbook.addWorksheet(strings, "S");
  for (let i = 0; i < 100_000; i++) {
    Worksheet.addRow(stringSheet, [`unique string number ${i} with some text`]);
  }
  const inputs: [string, Uint8Array][] = [
    ["xlsx, 10k rows, deflated", repack(written, 6, { first: prerequisitesFirst })],
    ["xlsx, 10k rows, stored", repack(written, 0, { first: prerequisitesFirst })],
    ["xlsb, 100k distinct strings", await Workbook.toBuffer(strings, { format: "xlsb" })]
  ];
  return Promise.all(
    inputs.map(async ([name, bytes]) => {
      const expected = await wholeRows(bytes);
      return {
        name: `Excel: ${name}`,
        run: () => streamedRows(bytes),
        check: (rows: unknown) => same(name, rows, expected)
      };
    })
  );
}

async function wordCases(): Promise<Case[]> {
  const doc = Document.create();
  const written = await Document.toBuffer(doc);
  const many = (count: number, item: (i: number) => string): string =>
    Array.from({ length: count }, (_, i) => item(i)).join("");
  const paragraph = (i: number): string =>
    `<w:p><w:r><w:t>Paragraph ${i} with some text</w:t></w:r></w:p>`;
  const body = (xml: string): string =>
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${xml}</w:body></w:document>`;
  const styles = new TextDecoder().decode(unpackZipSync(written).get("word/styles.xml")!.data);
  const paragraphs = Document.create();
  for (let i = 0; i < 20_000; i++) {
    Document.addParagraph(
      paragraphs,
      `Paragraph ${i} with a reasonable amount of body text in it.`
    );
  }
  const inputs: [string, Uint8Array][] = [
    ["20k paragraphs, stored", repack(await Document.toBuffer(paragraphs), 0)],
    [
      "one table of 20k rows",
      repack(written, 0, {
        replace: {
          "word/document.xml": body(
            `<w:tbl>${many(20_000, i => `<w:tr><w:tc>${paragraph(i)}</w:tc></w:tr>`)}</w:tbl>`
          )
        }
      })
    ],
    [
      "one table cell of 20k paragraphs",
      repack(written, 0, {
        replace: {
          "word/document.xml": body(
            `<w:tbl><w:tr><w:tc>${many(20_000, paragraph)}</w:tc></w:tr></w:tbl><w:p/>`
          )
        }
      })
    ],
    [
      "one paragraph of 20k runs",
      repack(written, 0, {
        replace: {
          "word/document.xml": body(
            `<w:p>${many(20_000, i => `<w:r><w:t>run ${i} </w:t></w:r>`)}</w:p>`
          )
        }
      })
    ],
    [
      "a style part of 5k styles",
      repack(written, 0, {
        replace: {
          "word/styles.xml": styles.replace(
            "</w:styles>",
            `${many(5_000, i => `<w:style w:type="paragraph" w:styleId="S${i}"><w:name w:val="Style ${i}"/><w:rPr><w:b/></w:rPr></w:style>`)}</w:styles>`
          )
        }
      })
    ],
    [
      "one hyperlink of 20k runs",
      repack(written, 0, {
        replace: {
          "word/document.xml": body(
            `<w:p><w:hyperlink w:anchor="a">${many(20_000, i => `<w:r><w:t>run ${i} </w:t></w:r>`)}</w:hyperlink></w:p>`
          )
        }
      })
    ],
    [
      "an inline content control of 20k runs",
      repack(written, 0, {
        replace: {
          "word/document.xml": body(
            `<w:p><w:sdt><w:sdtContent>${many(20_000, i => `<w:r><w:t>run ${i} </w:t></w:r>`)}</w:sdtContent></w:sdt></w:p>`
          )
        }
      })
    ],
    [
      "a Flat OPC document of 20k paragraphs",
      new TextEncoder().encode(await Document.toFlatOpc(paragraphs))
    ],
    [
      "a Flat OPC document with a 30 MiB binary part",
      new TextEncoder().encode(
        (await Document.toFlatOpc(doc)).replace(
          "</pkg:package>",
          `<pkg:part pkg:name="/word/media/large.bin" pkg:contentType="application/octet-stream"><pkg:binaryData>` +
            `${Buffer.from(new Uint8Array(30 * 1024 * 1024).map((_, i) => (i * 131) & 255))
              .toString("base64")
              .replace(/(.{76})/g, "$1\n")}</pkg:binaryData></pkg:part></pkg:package>`
        )
      )
    ]
  ];
  return Promise.all(
    inputs.map(async ([name, bytes]) => {
      const expected = (await Document.read(bytes)).body;
      return {
        name: `Word: ${name}`,
        run: () => streamedBody(bytes),
        check: (items: unknown) => same(name, items, expected)
      };
    })
  );
}

async function cryptoCases(): Promise<Case[]> {
  const zip = new Uint8Array(32 * 1024 * 1024).map((_, i) => (i * 7) & 255);
  const encrypted = await encryptPackage(zip, "pw");
  return [
    {
      name: "Office decryption, 32 MiB package",
      run: () => decryptPackage(encrypted, "pw"),
      check: result => {
        const out = result as Uint8Array;
        if (out.length !== zip.length || out.some((byte, i) => byte !== zip[i])) {
          throw new Error("decryption: the result differs from the package encrypted");
        }
      }
    }
  ];
}

/** Run `work`, reporting how long it took and the longest stretch no callback could run. */
async function measure(
  work: () => Promise<unknown>
): Promise<{ total: number; longest: number; result: unknown }> {
  let last = performance.now();
  let longest = 0;
  let running = true;
  const tick = (): void => {
    const now = performance.now();
    longest = Math.max(longest, now - last);
    last = now;
    if (running) {
      setImmediate(tick);
    }
  };
  setImmediate(tick);
  try {
    const start = performance.now();
    const result = await work();
    const total = performance.now() - start;
    // The stretch the work ended in is counted only when the callback next runs.
    await new Promise(resolve => setImmediate(resolve));
    return { total, longest, result };
  } finally {
    // A failed case must not leave the callback rescheduling itself for good.
    running = false;
  }
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)]!;
}

async function main(): Promise<void> {
  const cases = [...(await excelCases()), ...(await wordCases()), ...(await cryptoCases())];
  for (const c of cases) {
    await c.check(await c.run());
  }
  for (const c of cases) {
    await c.run();
  }
  const totals = new Map<string, number[]>(cases.map(c => [c.name, []]));
  const longest = new Map<string, number[]>(cases.map(c => [c.name, []]));
  for (let round = 0; round < RUNS; round++) {
    for (let k = 0; k < cases.length; k++) {
      const c = cases[(k + round) % cases.length]!;
      const measured = await measure(c.run);
      totals.get(c.name)!.push(measured.total);
      longest.get(c.name)!.push(measured.longest);
    }
  }
  console.log(`Yield budget ${YIELD_BUDGET_MS} ms; medians of ${RUNS} rounds\n`);
  console.log(`${"case".padEnd(54)}${"total".padStart(10)}${"longest block".padStart(16)}`);
  for (const c of cases) {
    console.log(
      `${c.name.padEnd(54)}${`${median(totals.get(c.name)!).toFixed(0)} ms`.padStart(10)}` +
        `${`${median(longest.get(c.name)!).toFixed(1)} ms`.padStart(16)}`
    );
  }
}

await main();
