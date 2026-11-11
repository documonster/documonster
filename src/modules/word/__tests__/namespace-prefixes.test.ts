/**
 * Word parts are read by the namespaces their names are in, not the prefixes they are spelled with:
 * a part that binds Word's namespace to another prefix, or makes it the default one, reads as Word
 * wrote it; one that binds `w` to another namespace is not read as Word's; and a Strict document's
 * namespaces read as their Transitional counterparts. `Document.read` and `Stream.DocumentReader`
 * agree on each.
 */
import { Archive } from "@archive/index";
import { Xml } from "@xml/index";
import { describe, expect, it } from "vitest";

import { Document, DocxParseError, Io, Stream } from "../index";
import type { BodyContent } from "../types";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** The text of the first run each reader reads from a package whose body is `xml`. */
async function bothReaders(xml: string): Promise<[unknown, unknown]> {
  const bytes = await Archive.packZip([["word/document.xml", xml]]);
  const textOf = (body: readonly BodyContent[]): string | undefined =>
    /"text":"([^"]*)"/.exec(JSON.stringify(body))?.[1];
  const read = async (body: () => Promise<readonly BodyContent[]>): Promise<unknown> => {
    try {
      return textOf(await body());
    } catch (err) {
      return err;
    }
  };
  return [
    await read(async () => (await Document.read(bytes)).body),
    await read(async () => {
      const out: BodyContent[] = [];
      for await (const item of new Stream.DocumentReader(bytes)) {
        out.push(item);
      }
      return out;
    })
  ];
}

const body = (prefix: string, uri: string): string => {
  const p = prefix === "" ? "" : `${prefix}:`;
  const declaration = prefix === "" ? `xmlns="${uri}"` : `xmlns:${prefix}="${uri}"`;
  return `<${p}document ${declaration}><${p}body><${p}p><${p}r><${p}t>ok</${p}t></${p}r></${p}p></${p}body></${p}document>`;
};

describe("Word reads a part by the namespaces of its names", () => {
  it.each([
    ["Word's namespace under another prefix", body("z", W)],
    ["Word's namespace as the default one", body("", W)],
    ["a Strict document", body("w", "http://purl.oclc.org/ooxml/wordprocessingml/main")]
  ])("%s", async (_label, xml) => {
    expect(await bothReaders(xml)).toEqual(["ok", "ok"]);
  });

  it("does not read another namespace bound to w as Word's", async () => {
    const [buffered, streamed] = await bothReaders(body("w", "urn:other"));
    expect(buffered).toBeInstanceOf(DocxParseError);
    expect(streamed).toBeInstanceOf(DocxParseError);
  });

  it("reads a part under other prefixes with the security policy, as one under w", async () => {
    const deep = (p: string): string =>
      `<${p}:document xmlns:${p}="${W}"><${p}:body>` +
      `<${p}:customXml>`.repeat(300) +
      `<${p}:p><${p}:r><${p}:t>ok</${p}:t></${p}:r></${p}:p>` +
      `</${p}:customXml>`.repeat(300) +
      `</${p}:body></${p}:document>`;
    for (const prefix of ["w", "z"]) {
      const bytes = await Archive.packZip([["word/document.xml", deep(prefix)]]);
      const read = await Document.read(bytes, { securityPolicy: { maxXmlDepth: 512 } });
      expect(JSON.stringify(read.body)).toContain('"ok"');
    }
  });

  it("refuses a character XML forbids under other prefixes, as under w", async () => {
    for (const prefix of ["w", "z"]) {
      const xml = body(prefix, W).replace(">ok<", ">A\u0001B<");
      const [buffered, streamed] = await bothReaders(xml);
      expect(buffered).toBeInstanceOf(Error);
      expect(streamed).toBeInstanceOf(Error);
    }
  });

  it("writes back a document read under other prefixes as one that declares what it uses", async () => {
    const xml =
      `<z:document xmlns:z="${W}" xmlns:q="${W14}" xmlns:k="${MC}" k:Ignorable="q">` +
      `<z:body><z:p q:paraId="1A2B3C4D"><z:r><z:t>ok</z:t></z:r></z:p></z:body></z:document>`;
    const read = await Document.read(await Archive.packZip([["word/document.xml", xml]]));
    const written = await Io.toBuffer(read);
    const part = new TextDecoder().decode(
      (await Archive.unpackZip(written)).get("word/document.xml")!.data
    );
    // Well-formed with namespaces: every prefix it uses is declared.
    expect(() => Xml.parse(part, { xmlns: true })).not.toThrow();
    const again = await Document.read(written);
    expect(JSON.stringify(again.body)).toContain('"ok"');
  });
});
