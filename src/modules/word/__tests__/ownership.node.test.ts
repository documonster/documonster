/**
 * A node built in another realm (an iframe, `node:vm`) is plain document data like any other: the copy a shared
 * document makes before a tree edit must copy it, not share it with the document it was derived from.
 */
import vm from "node:vm";

import { describe, expect, it } from "vitest";

import { Document, Query, Security } from "../index";
import type { Paragraph } from "../index";

describe("a shared document's tree edit, with nodes from another realm", () => {
  it("does not reach the derived document", async () => {
    const paragraph = vm.runInNewContext(
      '({ type: "paragraph", children: [{ content: [{ type: "text", text: "hello" }] }] })'
    ) as Paragraph;
    const a = Document.create();
    Document.addContent(a, paragraph);
    const b = await Security.protect(a, { edit: "readOnly" });

    expect(Document.replaceText(a, "hello", "changed")).toBe(1);

    expect(Query.getText(a)).toBe("changed");
    expect(Query.getText(b)).toBe("hello");
    expect(Document.getContent(a, 0)).not.toBe(Document.getContent(b, 0));
  });
});
