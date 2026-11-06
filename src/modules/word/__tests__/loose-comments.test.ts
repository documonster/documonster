/**
 * A comment nothing anchors is schema-valid, and the model's to keep: the
 * writers write every comment in `DocxDocument.comments`, anchored or not, so
 * a read → write of a package keeps a loose comment as it was.
 */

import { zip } from "@archive/create-archive";
import { describe, expect, it } from "vitest";

import { Build, Document, Io, Stream } from "../index";
import { packageParts, partText } from "./roundtrip-integrity.helpers";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** A hand-written package: comment 4 anchored in the body, comment 5 anchored nowhere. */
async function handWrittenDocx(): Promise<Uint8Array> {
  const enc = new TextEncoder();
  const archive = zip();
  archive.add(
    "[Content_Types].xml",
    enc.encode(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
        `</Types>`
    )
  );
  archive.add(
    "_rels/.rels",
    enc.encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
        `</Relationships>`
    )
  );
  archive.add(
    "word/_rels/document.xml.rels",
    enc.encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>` +
        `</Relationships>`
    )
  );
  archive.add(
    "word/document.xml",
    enc.encode(
      `<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeStart w:id="4"/><w:r><w:t>x</w:t></w:r>` +
        `<w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r></w:p></w:body></w:document>`
    )
  );
  archive.add(
    "word/comments.xml",
    enc.encode(
      `<w:comments xmlns:w="${W}"><w:comment w:id="4" w:author="A"><w:p><w:r><w:t>anchored</w:t></w:r></w:p></w:comment>` +
        `<w:comment w:id="5" w:author="B"><w:p><w:r><w:t>loose</w:t></w:r></w:p></w:comment></w:comments>`
    )
  );
  return archive.toBuffer();
}

describe("comments nothing anchors", () => {
  it("survive read → write of a package that carries one", async () => {
    const read = await Io.read(await handWrittenDocx());
    expect(read.comments?.map(c => c.id)).toEqual([4, 5]);
    const once = await Io.toBuffer(read);
    const comments = partText(await packageParts(once), "word/comments.xml");
    expect(comments).toContain('w:id="5"');
    expect(comments).toContain(">loose<");
    expect((await Io.read(once)).comments?.map(c => c.id)).toEqual([4, 5]);
  });

  it("are written byte for byte on every later round trip", async () => {
    const d = Document.create();
    const anchored = Document.addComment(d, "A", [
      { ...Build.textParagraph("anchored"), paraId: "0000000A" }
    ]);
    Document.addComment(d, "B", [{ ...Build.textParagraph("loose"), paraId: "0000000B" }]);
    Document.addContent(d, Build.paragraph([Build.text("x"), Build.commentReference(anchored)]));
    const written = await Io.toBuffer({
      ...d,
      comments: d.comments!.map(c => ({ ...c, done: false }))
    });
    const again = await Io.toBuffer(await Io.read(written));
    const [a, b] = await Promise.all([packageParts(written), packageParts(again)]);
    for (const path of ["word/comments.xml", "word/commentsExtended.xml"]) {
      expect(partText(b, path)).toBe(partText(a, path));
    }
    expect(partText(b, "word/comments.xml")).toContain(">loose<");
    expect(partText(b, "word/commentsExtended.xml")).toContain('w15:paraId="0000000B"');
  });

  it("are written by the streaming writer", async () => {
    const stream = new Stream.DocumentWriter({
      comments: [{ id: 2, author: "B", content: [Build.textParagraph("loose")] }]
    });
    stream.add(Build.textParagraph("body"));
    const comments = partText(await packageParts(await stream.finalize()), "word/comments.xml");
    expect(comments).toContain(">loose<");
  });
});
