/**
 * Errors the public calls throw are the module's own, name the public call and
 * the parameter at fault — not a bare `Error` naming an internal function.
 */

import { describe, expect, it } from "vitest";

import * as CryptoEntry from "../crypto";
import { DocxArgumentError, DocxError, DocxParseError, Diff, Document, Transform } from "../index";
import type { BodyContent, DocxDocument, Paragraph } from "../types";

const { WordCrypto } = CryptoEntry;

function para(text: string): Paragraph {
  return { type: "paragraph", children: [{ content: [{ type: "text", text }] }] };
}

function docOf(body: BodyContent[]): DocxDocument {
  return { body };
}

describe("Diff.compare", () => {
  it("checks both documents the same way", () => {
    const doc = Document.create();
    for (const [args, field] of [
      [[42, doc], "oldDoc"],
      [[doc, 42], "newDoc"],
      [[doc, {}], "newDoc"]
    ] as const) {
      const error = (() => {
        try {
          Diff.compare(...(args as unknown as [DocxDocument, DocxDocument]));
        } catch (e) {
          return e;
        }
      })();
      expect(error, field).toBeInstanceOf(DocxArgumentError);
      expect((error as DocxArgumentError).field, field).toBe(field);
      expect((error as Error).message, field).toMatch(/^Diff\.compare: /);
    }
  });

  it("refuses a pair too large to compare with a DocxError naming the call", () => {
    const big = docOf(Array.from({ length: 6000 }, (_, i) => para(String(i))));
    expect(() => Diff.compare(big, big)).toThrow(DocxError);
    expect(() => Diff.compare(big, big)).toThrow(/^Diff\.compare: LCS table for 6000x6000/);
  });
});

describe("Transform.map", () => {
  it("refuses a transform that puts a non-paragraph where only paragraphs may be", () => {
    const doc = docOf([{ type: "textBox", content: [para("in box")] }]);
    const table: BodyContent = { type: "table", rows: [] };
    const call = (): unknown =>
      Transform.map(doc, {
        transformBodyContent: block => (block.type === "paragraph" ? table : block)
      });
    expect(call).toThrow(DocxArgumentError);
    expect(call).toThrow(expect.objectContaining({ field: "transformer.transformBodyContent" }));
    expect(call).toThrow(/^Transform\.map: textBox\.content must remain Paragraph\[\]/);
  });
});

describe("WordCrypto signature helpers check their arguments", () => {
  it("parseSignature", () => {
    for (const [args, field] of [
      [[42, "sig1.xml"], "xml"],
      [["<Signature/>", undefined], "fileName"]
    ] as const) {
      expect(() => WordCrypto.parseSignature(...(args as unknown as [string, string]))).toThrow(
        expect.objectContaining({ name: "DocxArgumentError", field })
      );
      expect(() => WordCrypto.parseSignature(...(args as unknown as [string, string]))).toThrow(
        /^WordCrypto\.parseSignature: /
      );
    }
    expect(WordCrypto.parseSignature("<x/>", "sig1.xml").fileName).toBe("sig1.xml");
  });

  it("isWellFormedSignature", () => {
    expect(() => WordCrypto.isWellFormedSignature(null as never)).toThrow(
      expect.objectContaining({ name: "DocxArgumentError", field: "info" })
    );
    expect(() => WordCrypto.isWellFormedSignature({} as never)).toThrow(
      /^WordCrypto\.isWellFormedSignature: info\.rawXml /
    );
    const info = WordCrypto.parseSignature("<x/>", "sig1.xml");
    expect(WordCrypto.isWellFormedSignature(info)).toBe(false);
  });
});

describe("documonster/word/crypto names the errors it throws", () => {
  it("exports DocxError and DocxParseError, the same classes as the main entry", async () => {
    expect(CryptoEntry.DocxError).toBe(DocxError);
    expect(CryptoEntry.DocxParseError).toBe(DocxParseError);
    expect(() => WordCrypto.Cfb.read(new Uint8Array(16))).toThrow(CryptoEntry.DocxParseError);
  });
});
