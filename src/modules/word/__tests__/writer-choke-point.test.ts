/**
 * The writer is the one place every model goes through, so it refuses what it
 * cannot write — an unknown node kind, a list that is not an array, a number
 * with no XML form — with `DocxArgumentError` naming where the node is,
 * instead of dropping it or failing with another module's error. The entry
 * points run the same top-level check on what they are given.
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError } from "../errors";
import { Build, Document, Io } from "../index";
import type { BodyContent, DocxDocument, Paragraph } from "../types";

async function writeError(doc: DocxDocument): Promise<DocxArgumentError> {
  const error = await Io.toBuffer(doc).catch((e: unknown) => e);
  expect(error).toBeInstanceOf(DocxArgumentError);
  return error as DocxArgumentError;
}

function withBody(...body: unknown[]): DocxDocument {
  return { ...Document.create(), body: body as BodyContent[] };
}

describe("the writer refuses what it cannot write", () => {
  it("names the path of an unknown node kind", async () => {
    const bogusRun = { type: "paragraph", children: [{ content: [{ type: "bogus" }] }] };
    expect((await writeError(withBody(Build.textParagraph("x"), bogusRun))).field).toBe(
      "body[1].children[0].content[0].type"
    );
    expect((await writeError(withBody({ type: "bogus" }))).field).toBe("body[0].type");
    const bogusChild = { type: "paragraph", children: [{ type: "bogus" }] };
    expect((await writeError(withBody(bogusChild))).field).toBe("body[0].children[0].type");
  });

  it("refuses a list that is not an array", async () => {
    expect((await writeError(withBody({ type: "paragraph", children: "x" }))).field).toBe(
      "body[0].children"
    );
    expect((await writeError(withBody({ type: "table", rows: "x" }))).field).toBe("body[0].rows");
    const badCell = { type: "table", rows: [{ cells: [{ content: "x" }] }] };
    expect((await writeError(withBody(badCell))).field).toBe("body[0].rows[0].cells[0].content");
  });

  it("turns a number with no XML form into its own error, with the path", async () => {
    const para: Paragraph = Build.paragraph([
      { content: [{ type: "text", text: "x" }], properties: { size: Number.NaN } }
    ]);
    const error = await writeError(withBody(Build.textParagraph("ok"), para));
    expect(error.field.startsWith("body[1]")).toBe(true);
  });

  it("refuses an unknown node in a header, a note and a comment", async () => {
    const header = Document.create();
    (header as { headers?: Map<string, unknown> }).headers = new Map([
      ["default", { content: { children: [{ type: "bogus" }] } }]
    ]);
    expect((await writeError(header)).field).toBe('headers["default"].children[0].type');
  });
});

describe("the entry points run the same top-level check", () => {
  const bogus = { type: "bogus" } as never;

  it.each([
    [
      "setHeader",
      (d: DocxDocument) => Document.setHeader(d, "default", { children: [bogus] }),
      "content.children[0].type"
    ],
    [
      "setFooter",
      (d: DocxDocument) => Document.setFooter(d, "default", { children: [bogus] }),
      "content.children[0].type"
    ],
    ["addFootnote", (d: DocxDocument) => Document.addFootnote(d, [bogus]), "content[0].type"],
    ["addEndnote", (d: DocxDocument) => Document.addEndnote(d, [bogus]), "content[0].type"],
    ["addComment", (d: DocxDocument) => Document.addComment(d, "A", [bogus]), "content[0].type"],
    ["addTextBox", (d: DocxDocument) => Document.addTextBox(d, [bogus]), "content[0].type"]
  ] as const)("%s", (_, call, field) => {
    const d = Document.create();
    let error: unknown;
    try {
      call(d);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(DocxArgumentError);
    expect((error as DocxArgumentError).field).toBe(field);
  });
});
