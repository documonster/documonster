/**
 * The two readers of a package — `Document.read` and `Stream.DocumentReader` — and `Package.edit`
 * hold it to the same limits, at the point a part is inflated, and recognise the main part by its
 * namespace; what is kept verbatim is written back bound.
 */
import { createZip } from "@archive/create-archive";
import { unzipZip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { Xml } from "../../xml/index";
import { utf8Decoder, utf8Encoder } from "../core/internal-utils";
import { DocxLimitExceededError } from "../errors";
import { Document, Package, Security, Stream } from "../index";
import type { DocxSecurityPolicy } from "../index";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const P = "<w:p><w:r><w:t>hi</w:t></w:r></w:p>";

async function packageWith(
  documentXml: string,
  extra: Record<string, Uint8Array> = {},
  options: { level?: number; base?: Uint8Array } = {}
): Promise<Uint8Array> {
  const base = options.base ?? (await Document.toBuffer(Document.create()));
  const out = createZip({ level: options.level ?? 6 });
  for await (const entry of unzipZip(base).entries()) {
    out.add(
      entry.path,
      entry.path === "word/document.xml" ? utf8Encoder.encode(documentXml) : await entry.bytes()
    );
  }
  for (const [path, data] of Object.entries(extra)) {
    out.add(path, data);
  }
  return out.toBuffer();
}

async function streamAll(
  bytes: Uint8Array,
  options?: ConstructorParameters<typeof Stream.DocumentReader>[1]
): Promise<unknown[]> {
  const out: unknown[] = [];
  for await (const item of new Stream.DocumentReader(bytes, options)) {
    out.push(item);
  }
  return out;
}

async function uncompressedTotal(bytes: Uint8Array): Promise<number> {
  let total = 0;
  for await (const entry of unzipZip(bytes).entries()) {
    total += (await entry.bytes()).length;
  }
  return total;
}

const doc = (body: string, root = `xmlns:w="${W}"`): string =>
  `<w:document ${root}><w:body>${body}</w:body></w:document>`;

describe("package limits, the same in every reader", () => {
  it("counts a part neither reader inflates toward maxPackageSize", async () => {
    const plain = await packageWith(doc(P));
    const total = await uncompressedTotal(plain);
    const bytes = await packageWith(doc(P), {
      "word/media/unused.bin": new Uint8Array(20000)
    });
    const securityPolicy = { maxPackageSize: total + 10000 };
    await expect(Document.read(bytes, { securityPolicy })).rejects.toMatchObject({
      limit: "packageSize"
    });
    await expect(streamAll(bytes, { securityPolicy })).rejects.toMatchObject({
      limit: "packageSize"
    });
    // At the exact total both accept it.
    const exact = { maxPackageSize: total + 20000 };
    await expect(Document.read(bytes, { securityPolicy: exact })).resolves.toBeDefined();
    await expect(streamAll(bytes, { securityPolicy: exact })).resolves.toHaveLength(1);
  });

  it("refuses a part over maxPartSize before inflating it", async () => {
    const bytes = await packageWith(doc(P), {
      "word/media/big.bin": new Uint8Array(1 << 20)
    });
    const securityPolicy = { maxPartSize: 1024 * 64 };
    const read = Document.read(bytes, { securityPolicy });
    await expect(read).rejects.toBeInstanceOf(DocxLimitExceededError);
    // The declared size is refused: what is reported is the entry's declaration, not bytes inflated.
    await expect(read).rejects.toMatchObject({
      limit: "partSize",
      actual: 1 << 20
    });
    await expect(streamAll(bytes, { securityPolicy })).rejects.toMatchObject({
      limit: "partSize"
    });
    await expect(Package.listParts(bytes, { securityPolicy })).resolves.toContain(
      "word/media/big.bin"
    );
    await expect(
      Package.readPart(bytes, "word/media/big.bin", { securityPolicy })
    ).rejects.toMatchObject({ limit: "partSize" });
    await expect(Package.edit(bytes, [], { securityPolicy })).rejects.toMatchObject({
      limit: "partSize"
    });
  });

  it("holds the watermark's section scan to maxElementBytes", async () => {
    const d = Document.create();
    Document.addParagraph(d, "x");
    Document.setWatermark(d, { type: "text", text: "DRAFT" });
    const base = await Document.toBuffer(d);
    const original = utf8Decoder.decode(await unzipZip(base).bytes("word/document.xml"));
    const huge = `<w:pgSz w:w="1" w:h="1"/>`.repeat(20000);
    const bytes = await packageWith(
      original.replace("<w:sectPr>", `<w:sectPr>${huge}`),
      {},
      {
        base
      }
    );
    let progressed = false;
    await expect(
      streamAll(bytes, {
        maxElementBytes: 100,
        onProgress: () => {
          progressed = true;
        }
      })
    ).rejects.toMatchObject({ limit: "elementBytes" });
    // Refused by the scan that precedes the body, not by the body.
    expect(progressed).toBe(false);
  });

  it("honours checkCrc32: false when streaming, as Document.read does", async () => {
    const bytes = new Uint8Array(await packageWith(doc(P), {}, { level: 0 }));
    const at = new TextDecoder("latin1").decode(bytes).indexOf("hi</w:t>");
    bytes[at] = "H".charCodeAt(0);
    await expect(streamAll(bytes)).rejects.toThrow(/CRC32/);
    await expect(Document.read(bytes)).rejects.toThrow(/CRC32/);
    const streamed = await streamAll(bytes, { checkCrc32: false });
    expect(JSON.stringify(streamed)).toContain('"Hi"');
    expect(JSON.stringify((await Document.read(bytes, { checkCrc32: false })).body)).toContain(
      '"Hi"'
    );
  });
});

describe("the main part is recognised by its namespace", () => {
  const refused = async (xml: string): Promise<void> => {
    const bytes = await packageWith(xml);
    await expect(Document.read(bytes)).rejects.toThrow(/Missing w:body/);
    await expect(streamAll(bytes)).rejects.toThrow(/Missing w:body/);
  };

  it("refuses a document in another namespace", () =>
    refused(`<document xmlns="urn:WRONG"><body><p><r><t>WRONG</t></r></p></body></document>`));

  it("refuses an unbound w: prefix", () =>
    refused(`<w:document><w:body>${P}</w:body></w:document>`));

  it("refuses a body that is not the document's child", () =>
    refused(
      `<w:document xmlns:w="${W}" xmlns:x="urn:x"><x:wrapper><w:body>${P}</w:body></x:wrapper></w:document>`
    ));

  it("reads Word's namespace as the default one", async () => {
    const bytes = await packageWith(
      `<document xmlns="${W}"><body><p><r><t>hi</t></r></p></body></document>`
    );
    expect((await Document.read(bytes)).body).toHaveLength(1);
    expect(await streamAll(bytes)).toHaveLength(1);
  });
});

describe("markup kept verbatim is written back bound", () => {
  const cases = [
    `<w:p><u:thing a="1"/>${P.slice(5)}`,
    `<w:p><w:r><u:thing/><w:t>hi</w:t></w:r></w:p>`
  ];
  for (const body of cases) {
    it(`keeps the root's declaration for ${body}`, async () => {
      const bytes = await packageWith(doc(body, `xmlns:w="${W}" xmlns:u="urn:u"`));
      const written = await Document.toBuffer(await Document.read(bytes));
      const xml = utf8Decoder.decode(await unzipZip(written).bytes("word/document.xml"));
      expect(() => Xml.parse(xml, { xmlns: true })).not.toThrow();
      expect(xml).toContain(`xmlns:u="urn:u"`);
    });
  }

  it("keeps a foreign default namespace", async () => {
    const bytes = await packageWith(
      doc(`<w:p><thing/>${P.slice(5)}`, `xmlns:w="${W}" xmlns="urn:dflt"`)
    );
    const written = await Document.toBuffer(await Document.read(bytes));
    const xml = utf8Decoder.decode(await unzipZip(written).bytes("word/document.xml"));
    expect(xml).toContain(`<thing xmlns="urn:dflt"/>`);
  });

  it("binds what the streaming reader keeps too", async () => {
    const bytes = await packageWith(
      doc(`<w:p><w:r><u:thing/><w:t>hi</w:t></w:r></w:p>`, `xmlns:w="${W}" xmlns:u="urn:u"`)
    );
    expect(JSON.stringify(await streamAll(bytes))).toContain('xmlns:u=\\"urn:u\\"');
  });
});

describe("Package.edit's content types", () => {
  const typesWith = (defaults: string): string =>
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}<Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;

  it("does not take a Default without a ContentType as a declaration", async () => {
    const bytes = await packageWith(doc(P));
    await expect(
      Package.edit(bytes, [
        {
          type: "replacePartText",
          path: "[Content_Types].xml",
          text: typesWith(`<Default Extension="png"/>`)
        },
        {
          type: "replacePart",
          path: "word/media/a.png",
          data: new Uint8Array(3)
        }
      ])
    ).rejects.toThrow(/has no content type/);
  });

  it("reads [Content_Types].xml under maxXmlDepth", async () => {
    const bytes = await packageWith(doc(P));
    await expect(
      Package.edit(bytes, [{ type: "replacePart", path: "word/x.bin", data: new Uint8Array(3) }], {
        securityPolicy: { maxXmlDepth: 1 }
      })
    ).rejects.toMatchObject({ limit: "xmlDepth" });
  });
});

describe("a security policy is read once", () => {
  const flipping = (): DocxSecurityPolicy => {
    let reads = 0;
    return {
      get maxPartSize() {
        return reads++ === 0 ? 10 : Number.NaN;
      }
    } as DocxSecurityPolicy;
  };

  it("uses the value it checked", () => {
    expect(Security.resolvePolicy(flipping()).maxPartSize).toBe(10);
  });

  it("holds every reader to the value it checked", async () => {
    const bytes = await packageWith(doc(P));
    await expect(Document.read(bytes, { securityPolicy: flipping() })).rejects.toMatchObject({
      limit: "partSize"
    });
    await expect(streamAll(bytes, { securityPolicy: flipping() })).rejects.toMatchObject({
      limit: "partSize"
    });
  });
});

describe("a policy is read once by the calls that write", () => {
  /** A policy whose `maxPartSize` answers 32 once and `NaN` after that. */
  function flipping(): { reads: () => number; policy: DocxSecurityPolicy } {
    let reads = 0;
    const policy = {
      get maxPartSize() {
        return ++reads === 1 ? 32 : Number.NaN;
      }
    } as DocxSecurityPolicy;
    return { reads: () => reads, policy };
  }

  it("Package.patch reads and writes under the value it checked", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    const bytes = await Document.toBuffer(doc);
    const { reads, policy } = flipping();
    await expect(Package.patch(bytes, [], { securityPolicy: policy })).rejects.toMatchObject({
      limit: "partSize"
    });
    expect(reads()).toBe(1);
  });

  it("Document.toBuffer reads the policy once", async () => {
    const doc = Document.create();
    const { reads, policy } = flipping();
    await Document.toBuffer(doc, { securityPolicy: policy });
    expect(reads()).toBe(1);
  });
});

describe("maxElementBytes counts UTF-8 bytes exactly", () => {
  it("passes an element of exactly the limit and refuses one byte less", async () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    // Three bytes a character: the count crosses from units to bytes partway through the element.
    const text = "€".repeat(40);
    const xml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`;
    const zip = createZip();
    zip.add("word/document.xml", utf8Encoder.encode(xml));
    const bytes = await zip.toBuffer();
    const exact = utf8Encoder.encode(`w:pw:rw:t${text}`).length;
    const read = async (maxElementBytes: number): Promise<number> => {
      let count = 0;
      for await (const _ of new Stream.DocumentReader(bytes, { maxElementBytes })) {
        count++;
      }
      return count;
    };
    expect(await read(exact)).toBe(1);
    await expect(read(exact - 1)).rejects.toMatchObject({ limit: "elementBytes", actual: exact });
  });
});
