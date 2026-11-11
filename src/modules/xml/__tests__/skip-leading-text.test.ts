/**
 * `skipLeadingText`: what comes before a document's first `<` — the lines `unzip -c` writes before a
 * file's content — read as no part of the document, with positions kept the file's.
 */
import { Xml, XmlChunkAmbiguityError, XmlParseError } from "@xml/index";
import type { XmlReadRecordsResult } from "@xml/index";
import { SaxParser } from "@xml/sax";
import { describe, expect, it } from "vitest";

const encoder = new TextEncoder();
const PREFIX = "Archive:  spotify_20240101.zip\n  inflating: report_é.xml\n";

function document(records: number): string {
  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE DSR [<!ENTITY l "Ünï">]>\n<DSR>\n';
  for (let i = 0; i < records; i++) {
    xml += `  <SalesRecord Id="${i}"><Title>Låt ${i} &l;</Title></SalesRecord>\n`;
  }
  return xml + "</DSR>\n";
}

describe("skipLeadingText", () => {
  it("is off by default: the text is refused at its first byte", async () => {
    const bytes = encoder.encode(PREFIX + document(1));
    await expect(Xml.readRoot(bytes)).rejects.toMatchObject({ byte: 0 });
    await expect(Xml.readRecords(bytes, { tags: ["SalesRecord"], onRecord() {} })).rejects.toThrow(
      XmlParseError
    );
  });

  it("skips to the first `<`, in any chunking, and reports how much, in bytes or characters", () => {
    const text = PREFIX + '<?xml version="1.0"?><r/>';
    for (const input of [text, encoder.encode(text)]) {
      for (const size of [1, 7, 4096]) {
        const parser = new SaxParser({ skipLeadingText: true });
        const texts: string[] = [];
        parser.on("text", t => void texts.push(t));
        for (let at = 0; at < input.length; at += size) {
          parser.write(input.slice(at, at + size));
        }
        parser.close();
        expect(texts).toEqual([]);
        expect(parser.skippedLeadingText).toBe(
          typeof input === "string" ? PREFIX.length : encoder.encode(PREFIX).length
        );
      }
    }
  });

  it("still refuses what follows that is not a whole document", () => {
    // A `<` in the leading text begins the document there.
    expect(() =>
      new SaxParser({ skipLeadingText: true })
        .write('junk <oops> more\n<?xml version="1.0"?><r/>')
        .close()
    ).toThrow(XmlParseError);
    // Text after the root is still text outside it.
    expect(() =>
      new SaxParser({ skipLeadingText: true }).write(`${PREFIX}<r/>trailing`).close()
    ).toThrow(XmlParseError);
  });

  it("reads records at the file's positions, and joins partitions as the whole read", async () => {
    const xml = PREFIX + document(200);
    const bytes = encoder.encode(xml);
    const skipped = encoder.encode(PREFIX).length;
    const tags = ["SalesRecord"];
    const root = await Xml.readRoot(bytes, { skipLeadingText: true });
    expect(root).toMatchObject({ name: "DSR", skippedLeadingBytes: skipped });
    const whole: string[] = [];
    const result = await Xml.readRecords(bytes, {
      tags,
      skipLeadingText: true,
      onRecord: r => void whole.push(`${r.startByte} ${JSON.stringify(r.object)}`)
    });
    expect(result.skippedLeadingBytes).toBe(skipped);
    expect(whole).toHaveLength(200);
    const first = xml.indexOf("<SalesRecord");
    expect(whole[0].startsWith(`${encoder.encode(xml.slice(0, first)).length} `)).toBe(true);

    let agreed = 0;
    for (const size of [37, 500, 4096]) {
      const got: string[] = [];
      const results: XmlReadRecordsResult[] = [];
      try {
        for (let offset = 0; offset < bytes.length; offset += size) {
          results.push(
            await Xml.readRecords(bytes.subarray(offset), {
              offset,
              size,
              context: root.context,
              tags,
              skipLeadingText: true,
              onRecord: r => void got.push(`${r.startByte} ${JSON.stringify(r.object)}`)
            })
          );
        }
        Xml.verifyPartition(JSON.parse(JSON.stringify(results)));
      } catch (error) {
        if (!(error instanceof XmlChunkAmbiguityError)) {
          throw error;
        }
        continue;
      }
      expect(got).toEqual(whole);
      agreed++;
    }
    expect(agreed).toBeGreaterThan(0);
  });
});
