/**
 * A container holding more blocks than a call can take arguments is read, not refused.
 *
 * The reader appended a container's blocks with `push(...blocks)`, which passes each as an argument: past about a
 * hundred thousand the call stack overflowed and the whole read failed with `Maximum call stack size exceeded` — a
 * document Word opens, refused as damaged. The length is the document's, so it has no bound to stay under.
 */
import { Archive } from "@archive/index";
import { describe, expect, it } from "vitest";

import { Document } from "../index";

const BLOCKS = 200_000;
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function withBody(body: string): Promise<Uint8Array> {
  const parts = await Archive.unpackZip(await Document.toBuffer(Document.create()));
  const entries = new Map<string, Uint8Array>();
  for (const [name, part] of parts) {
    entries.set(name, part.data);
  }
  entries.set(
    "word/document.xml",
    new TextEncoder().encode(
      `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`
    )
  );
  return Archive.packZip(entries, { level: 0 });
}

const paragraphs = (): string => {
  let xml = "";
  for (let i = 0; i < BLOCKS; i++) {
    xml += `<w:p><w:r><w:t>${i}</w:t></w:r></w:p>`;
  }
  return xml;
};

describe("a container of very many blocks", () => {
  it("reads custom XML markup around them", async () => {
    const doc = await Document.read(
      await withBody(`<w:customXml w:element="x">${paragraphs()}</w:customXml>`)
    );
    expect(doc.body).toHaveLength(BLOCKS);
  });

  it("reads a content control around them in a table cell", async () => {
    const doc = await Document.read(
      await withBody(
        `<w:tbl><w:tr><w:tc><w:sdt><w:sdtContent>${paragraphs()}</w:sdtContent></w:sdt></w:tc></w:tr></w:tbl><w:p/>`
      )
    );
    const table = doc.body[0];
    expect(table?.type === "table" && table.rows[0]!.cells[0]!.content).toHaveLength(BLOCKS);
  });
});
