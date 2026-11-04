/**
 * A written story part never holds half a range. Editing can leave a
 * bookmark's or a range permission's end before its start (moving a block),
 * or put its start and end in different documents (splitting one): a range
 * that cannot be paired in a part is removed from that part, start and end
 * together, and every other range is written as it was.
 */

import { describe, expect, it } from "vitest";

import { Build, Document, Io, Streaming, Transform } from "../index";
import type { DocxDocument, PermEnd, PermStart } from "../index";
import { packageIssues, packageParts, partText } from "./roundtrip-integrity.helpers";

/** `[start|end, id]` of each marker of `kind` in `xml`, in document order. */
function markers(xml: string, kind: "bookmark" | "perm"): string[] {
  return [...xml.matchAll(new RegExp(`<w:${kind}(Start|End)\\b[^>]*?\\bw:id="([^"]*)"`, "g"))].map(
    m => `${m[1]}:${m[2]}`
  );
}

async function written(doc: DocxDocument): Promise<{ issues: string[]; body: string }> {
  const parts = await packageParts(await Io.toBuffer(doc));
  return { issues: packageIssues(parts), body: partText(parts, "word/document.xml") };
}

function permStart(id: number): PermStart {
  return { type: "permStart", id };
}

function permEnd(id: number): PermEnd {
  return { type: "permEnd", id };
}

/** P0 opens bookmark 1 and permission 7, P1 closes both, P2 holds a range of its own. */
function crossParagraphDocument(): DocxDocument {
  const d = Document.create();
  Document.addContent(
    d,
    Build.paragraph([
      Build.bookmarkStart(1, "across"),
      { ...Build.text("open"), rangeMarkersBefore: [permStart(7)] }
    ])
  );
  Document.addContent(d, {
    ...Build.paragraph([Build.text("close"), Build.bookmarkEnd(1)]),
    rangeMarkersAfter: [permEnd(7)]
  });
  Document.addContent(
    d,
    Build.paragraph([Build.bookmarkStart(2, "own"), Build.text("own"), Build.bookmarkEnd(2)])
  );
  return d;
}

describe("range markers that cannot be paired", () => {
  it("removes a range whose end was moved before its start, and keeps the others", async () => {
    const d = crossParagraphDocument();
    Document.insertContent(d, 0, Document.removeContent(d, 1)!);
    const { issues, body } = await written(d);
    expect(issues).toEqual([]);
    expect(markers(body, "bookmark")).toEqual(["Start:2", "End:2"]);
    expect(markers(body, "perm")).toEqual([]);
    expect(body).not.toContain('w:name="across"');
  });

  it("keeps a range written in order", async () => {
    const { issues, body } = await written(crossParagraphDocument());
    expect(issues).toEqual([]);
    expect(markers(body, "bookmark")).toEqual(["Start:1", "End:1", "Start:2", "End:2"]);
    expect(markers(body, "perm")).toEqual(["Start:7", "End:7"]);
  });

  it("removes a start whose end is in another part, and the end there", async () => {
    const d = Document.create();
    Document.addContent(d, Build.paragraph([Build.bookmarkStart(5, "body"), Build.text("b")]));
    Document.setHeader(d, "default", {
      children: [Build.paragraph([Build.text("h"), Build.bookmarkEnd(5)])]
    });
    const parts = await packageParts(await Io.toBuffer(d));
    expect(packageIssues(parts)).toEqual([]);
    expect(markers(partText(parts, "word/document.xml"), "bookmark")).toEqual([]);
    expect(markers(partText(parts, "word/header1.xml"), "bookmark")).toEqual([]);
  });

  it("each part of a split document holds whole ranges only", async () => {
    const d = Document.create();
    Document.addHeading(d, "H1", 1);
    Document.addContent(d, Build.paragraph([Build.bookmarkStart(1, "x"), Build.text("one")]));
    Document.addHeading(d, "H2", 1);
    Document.addContent(d, Build.paragraph([Build.text("two"), Build.bookmarkEnd(1)]));
    const pieces = Transform.split(d, { by: "heading", headingLevel: 1 });
    expect(pieces).toHaveLength(2);
    for (const piece of pieces) {
      const { issues, body } = await written(piece);
      expect(issues).toEqual([]);
      expect(markers(body, "bookmark")).toEqual([]);
    }
  });

  it("a streamed body closes a range left open at its end", async () => {
    const stream = Streaming.createWriter();
    stream.add(Build.paragraph([Build.bookmarkStart(3, "tail"), Build.text("open")]));
    stream.add(Build.paragraph([Build.text("later")]));
    const parts = await packageParts(await stream.finalize());
    expect(packageIssues(parts)).toEqual([]);
    expect(markers(partText(parts, "word/document.xml"), "bookmark")).toEqual(["Start:3", "End:3"]);
  });
});
