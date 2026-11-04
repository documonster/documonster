/**
 * The siblings of the entry points r2 checked: decryption limits and
 * passwords, the streaming writer's and reader's options, and `options`
 * objects everywhere — each refused where it is passed, not at the end.
 */

import { describe, it, expect } from "vitest";

import { WordCrypto } from "../crypto";
import { DocxArgumentError } from "../errors";
import { WordHtml } from "../html";
import { Build, Convert, Document, Io, Query, Security, Streaming, Validation } from "../index";
import type { DocxDocument } from "../index";
import { WordMarkdown } from "../markdown";

async function refused(label: string, run: () => unknown, field?: string): Promise<void> {
  let error: unknown;
  try {
    await run();
  } catch (e) {
    error = e;
  }
  expect(error, label).toBeInstanceOf(DocxArgumentError);
  if (field !== undefined) {
    expect((error as DocxArgumentError).field, label).toBe(field);
  }
}

function doc(): DocxDocument {
  const d = Document.create();
  Document.addParagraph(d, "x");
  return d;
}

describe("decryption", () => {
  it("checks the size cap and the password before it decrypts", async () => {
    const plain = await Io.toBuffer(doc());
    const enc = await WordCrypto.encrypt(plain, "pw", { spinCount: 1 });
    await refused(
      "cap NaN",
      () => WordCrypto.decrypt(enc, "pw", { maxDecryptedSize: Number.NaN }),
      "options.maxDecryptedSize"
    );
    await refused(
      "cap x",
      () => WordCrypto.decrypt(enc, "pw", { maxDecryptedSize: "x" as never }),
      "options.maxDecryptedSize"
    );
    await refused(
      "cap -1",
      () => WordCrypto.decrypt(enc, "pw", { maxDecryptedSize: -1 }),
      "options.maxDecryptedSize"
    );
    await refused("options", () => WordCrypto.decrypt(enc, "pw", "x" as never), "options");
    await refused("password 42", () => WordCrypto.decrypt(enc, 42 as never), "password");
    await refused("password null", () => WordCrypto.decrypt(enc, null as never), "password");
    await refused("verify", () => WordCrypto.verifyPassword(enc, 42 as never), "password");
    await refused("Io.read", () => Io.read(enc, { password: 42 as never }), "options.password");
    expect(
      await WordCrypto.decrypt(enc, "pw", { maxDecryptedSize: Number.POSITIVE_INFINITY })
    ).toEqual(plain);
  });
});

describe("streaming writer options", () => {
  it("are checked by createWriter, as Document.set* checks them", async () => {
    const header = { content: { children: [Build.textParagraph("h")] } };
    await refused("pageSize NaN", () =>
      Streaming.createWriter({ sectionProperties: { pageSize: { width: Number.NaN, height: 1 } } })
    );
    await refused("pageSize -1", () =>
      Streaming.createWriter({ sectionProperties: { pageSize: { width: -1, height: 1 } } })
    );
    await refused(
      "headers",
      () => Streaming.createWriter({ headers: "x" as never }),
      "options.headers"
    );
    await refused("header content", () =>
      Streaming.createWriter({ headers: new Map([["default", { content: "x" }]]) as never })
    );
    await refused(
      "styles",
      () => Streaming.createWriter({ styles: "x" as never }),
      "options.styles"
    );
    await refused(
      "created",
      () => Streaming.createWriter({ coreProperties: { created: new Date(Number.NaN) } }),
      "options.coreProperties.created"
    );
    await refused(
      "settings",
      () => Streaming.createWriter({ settings: { zoom: -5 } }),
      "options.settings.zoom"
    );
    await refused(
      "app",
      () => Streaming.createWriter({ appProperties: { pages: Number.NaN } }),
      "options.appProperties.pages"
    );
    await refused("options", () => Streaming.createWriter("x" as never), "options");
    expect(Streaming.createWriter({ headers: new Map([["default", header]]) })).toBeDefined();
  });

  it("checks addText's properties and onProgress, naming the method", async () => {
    const w = Streaming.createWriter();
    await refused("onProgress", () => w.onProgress(null as never), "cb");
    await refused("addText props", () => w.addText("x", "y" as never), "properties");
    let error: unknown;
    try {
      w.addText("x", { alignment: "bogus" as never });
    } catch (e) {
      error = e;
    }
    expect((error as Error).message).toMatch(/^Streaming\.StreamingDocxWriter#addText: /);
  });

  it("checks a whole batch before writing any of it, naming the element", async () => {
    const w = Streaming.createWriter();
    const good = { type: "paragraph", children: [] } as const;
    let error: unknown;
    try {
      w.addMany([good, { type: "bogus" } as never]);
    } catch (e) {
      error = e;
    }
    expect((error as { field?: string }).field).toBe("elements[1].type");
    expect((error as Error).message).toMatch(/^Streaming\.StreamingDocxWriter#addMany: /);
    await refused("addMany", () => w.addMany(42 as never), "elements");
    // Nothing of the refused batch was written: the result is an untouched writer's.
    const doc = await Io.read(await w.finalize());
    const untouched = await Io.read(await Streaming.createWriter().finalize());
    expect(doc.body).toEqual(untouched.body);
  });
});

describe("streaming reader", () => {
  it("checks its source and options up front, and reads Flat OPC as Io.read does", async () => {
    await refused("null", () => Streaming.createReader(null as never), "source");
    await refused(
      "options",
      () => Streaming.createReader(new Uint8Array(4), "x" as never),
      "options"
    );
    const d = doc();
    const flat = new TextEncoder().encode(await Io.toFlatOpc(d));
    const types: string[] = [];
    for await (const block of Streaming.createReader(flat)) {
      types.push(block.type);
    }
    expect(types).toEqual(["paragraph"]);
  });
});

describe("options objects", () => {
  it("are refused when they are not objects", async () => {
    const d = doc();
    const bytes = await Io.toBuffer(d);
    await refused("Io.read", () => Io.read(bytes, "x" as never), "options");
    await refused("Validation.validate", () => Validation.validate(d, "x" as never), "options");
    await refused("WordHtml.fromDocx", () => WordHtml.fromDocx(d, "x" as never), "options");
    await refused("WordMarkdown.fromDocx", () => WordMarkdown.fromDocx(d, "x" as never), "options");
    await refused("Convert.toSemantic", () => Convert.toSemantic(d, "x" as never), "options");
    await refused("Query.listTables", () => Query.listTables(d, "x" as never), "options");
  });

  it("refuses an unknown key in a security policy, where a typo weakens a limit", async () => {
    await refused(
      "resolvePolicy",
      () => Security.resolvePolicy({ maxEntries: 1 } as never),
      "policy.maxEntries"
    );
    const bytes = await Io.toBuffer(doc());
    await refused("Io.read", () => Io.read(bytes, { securityPolicy: { maxEntries: 1 } as never }));
    // Other options objects keep taking keys they do not know: callers spread
    // objects read from files into them.
    await expect(Io.read(bytes, { bogus: 1 } as never)).resolves.toBeDefined();
  });
});
