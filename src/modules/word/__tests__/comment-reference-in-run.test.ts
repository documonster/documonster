/**
 * `w:commentReference` and `w:annotationRef` are different marks (ECMA-376
 * §17.13.4.5 and §17.13.4.1): the first anchors a comment in the story and
 * carries its id; the second sits inside the comment's own body and has no
 * attributes. A reference that shares its run with other content must be
 * read as the comment's anchor — the paragraph child `CommentReference`, with
 * its id, its place (the run is split around it) and the run's formatting —
 * and every operation that moves or drops content must treat it as one.
 */
import { unzip } from "@archive/read-archive";
import { createZip } from "@archive/zip/zip-bytes";
import { describe, it, expect } from "vitest";

import { Build, Document, Io, Query, Template, Transform, Validation } from "../index";
import type { DocxDocument, Paragraph, Run } from "../types";

const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const enc = new TextEncoder();

/** A package with `body` as its document body and comment 3 (`text`) defined. */
async function docx(body: string, commentText = "note"): Promise<Uint8Array> {
  return createZip([
    {
      name: "[Content_Types].xml",
      data: enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`
      )
    },
    {
      name: "_rels/.rels",
      data: enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
      )
    },
    {
      name: "word/_rels/document.xml.rels",
      data: enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`
      )
    },
    {
      name: "word/document.xml",
      data: enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`
      )
    },
    {
      name: "word/comments.xml",
      data: enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W}><w:comment w:id="3" w:author="A"><w:p><w:r><w:annotationRef/></w:r><w:r><w:t>${commentText}</w:t></w:r></w:p></w:comment></w:comments>`
      )
    }
  ]);
}

/** A run holding text, the reference to comment 3, and more text. */
const MIXED_RUN = `<w:r><w:rPr><w:b/></w:rPr><w:t>see</w:t><w:commentReference w:id="3"/><w:t>more</w:t></w:r>`;
const COMMENTED = `<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>text</w:t></w:r><w:commentRangeEnd w:id="3"/>${MIXED_RUN}</w:p>`;

async function parts(bytes: Uint8Array): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for await (const entry of unzip(bytes).entries()) {
    if (entry.path.endsWith(".xml")) {
      out.set(entry.path, new TextDecoder().decode(await entry.bytes()));
    }
  }
  return out;
}

/** Comment ids, and the ids of the references to them, in a written package. */
async function anchors(doc: DocxDocument): Promise<{ comments: string[]; references: string[] }> {
  const p = await parts(await Io.toBuffer(doc));
  const ids = (xml: string | undefined, tag: string) =>
    [...(xml ?? "").matchAll(new RegExp(`<w:${tag} [^>]*?w:id="(\\d+)"`, "g"))].map(m => m[1]!);
  return {
    comments: ids(p.get("word/comments.xml"), "comment"),
    references: ids(p.get("word/document.xml"), "commentReference")
  };
}

describe("a comment reference sharing a run with other content", () => {
  it("is read as the comment's reference, splitting the run around it", async () => {
    const doc = await Io.read(await docx(COMMENTED));
    const children = (doc.body[0] as Paragraph).children;
    expect(children.slice(-3)).toEqual([
      { properties: { bold: true }, content: [{ type: "text", text: "see" }] },
      { type: "commentReference", id: 3, properties: { bold: true } },
      { properties: { bold: true }, content: [{ type: "text", text: "more" }] }
    ]);
  });

  it("is read with the mark's formatting when it is alone in its run", async () => {
    const doc = await Io.read(
      await docx(
        `<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>t</w:t></w:r><w:commentRangeEnd w:id="3"/>` +
          `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="3"/></w:r></w:p>`
      )
    );
    expect((doc.body[0] as Paragraph).children.at(-1)).toEqual({
      type: "commentReference",
      id: 3,
      properties: { style: "CommentReference" }
    });
    const xml = (await parts(await Io.toBuffer(doc))).get("word/document.xml")!;
    expect(xml).toContain(
      '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="3"/></w:r>'
    );
  });

  it("is written back as three runs, the middle one the reference", async () => {
    const doc = await Io.read(await docx(COMMENTED));
    const xml = (await parts(await Io.toBuffer(doc))).get("word/document.xml")!;
    expect(xml).toMatch(
      /<w:r><w:rPr><w:b\/><\/w:rPr><w:t[^>]*>see<\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:commentReference w:id="3"\/><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t[^>]*>more<\/w:t><\/w:r>/
    );
    expect(xml).not.toContain("annotationRef");
    const again = await Io.read(await Io.toBuffer(doc));
    expect((again.body[0] as Paragraph).children).toEqual((doc.body[0] as Paragraph).children);
  });

  it("splits a run inside a tracked insertion or deletion the same way", async () => {
    const doc = await Io.read(
      await docx(
        `<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>t</w:t></w:r><w:commentRangeEnd w:id="3"/>` +
          `<w:ins w:id="9" w:author="A"><w:r><w:t>a</w:t><w:commentReference w:id="3"/><w:t>b</w:t></w:r></w:ins>` +
          `<w:del w:id="10" w:author="A"><w:r><w:delText>c</w:delText><w:commentReference w:id="3"/></w:r></w:del></w:p>`
      )
    );
    const kinds = (doc.body[0] as Paragraph).children.map(c =>
      "type" in c ? (c.type === "commentReference" ? `ref${c.id}` : c.type) : "run"
    );
    expect(kinds).toEqual([
      "commentRangeStart",
      "run",
      "commentRangeEnd",
      "insertedRun",
      "ref3",
      "insertedRun",
      "deletedRun",
      "ref3"
    ]);
  });

  it("leaves the comment body's own mark an annotationRef", async () => {
    const doc = await Io.read(await docx(COMMENTED));
    const body = doc.comments![0]!.content[0]!;
    expect((body.children[0] as Run).content).toEqual([{ type: "annotationReference", id: 0 }]);
    const comments = (await parts(await Io.toBuffer(doc))).get("word/comments.xml")!;
    expect(comments).toContain("<w:annotationRef/>");
    expect(comments).not.toContain("commentReference");
  });

  it("survives the template removing the directive paragraph that holds it", async () => {
    const doc = await Io.read(
      await docx(
        `<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>{{#each items}}</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t>{{name}}</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t>{{/each}}</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:t> </w:t><w:commentReference w:id="3"/></w:r></w:p>`
      )
    );
    const filled = Template.fill(doc, { items: [{ name: "a" }] });
    expect(await anchors(filled)).toEqual({ comments: ["3"], references: ["3"] });
  });

  it("survives the template filling the run that holds it", async () => {
    const doc = await Io.read(
      await docx(
        `<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>Hi {{name}}</w:t><w:commentReference w:id="3"/></w:r><w:commentRangeEnd w:id="3"/></w:p>`
      )
    );
    const filled = Template.fill(doc, { name: "Ada" });
    expect(Query.getText(filled)).toBe("Hi Ada");
    expect(await anchors(filled)).toEqual({ comments: ["3"], references: ["3"] });
  });

  it("is kept on one copy when a template loop repeats it", async () => {
    const doc = await Io.read(
      await docx(
        `<w:p><w:r><w:t>{{#each items}}</w:t></w:r></w:p>${COMMENTED.replace("text", "{{name}}")}<w:p><w:r><w:t>{{/each}}</w:t></w:r></w:p>`
      )
    );
    const filled = Template.fill(doc, { items: [{ name: "a" }, { name: "b" }, { name: "c" }] });
    expect(await anchors(filled)).toEqual({ comments: ["3"], references: ["3"] });
  });

  it("follows its comment when Transform.merge renumbers it", async () => {
    const a = await Io.read(await docx(COMMENTED, "first"));
    const b = await Io.read(await docx(COMMENTED, "second"));
    const merged = Transform.merge([a, b]);
    const { comments, references } = await anchors(merged);
    expect(comments).toHaveLength(2);
    expect(new Set(comments).size).toBe(2);
    expect(references.sort()).toEqual([...comments].sort());
  });

  it("is kept as the anchor when rejecting the insertion that holds it", async () => {
    const doc = await Io.read(
      await docx(
        `<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>kept</w:t></w:r><w:commentRangeEnd w:id="3"/>` +
          `<w:ins w:id="9" w:author="A"><w:r><w:t>added</w:t><w:commentReference w:id="3"/></w:r></w:ins></w:p>`
      )
    );
    Query.rejectAllRevisions(doc);
    expect(Query.getText(doc)).toBe("kept");
    expect(await anchors(doc)).toEqual({ comments: ["3"], references: ["3"] });
  });

  it("is checked against the defined comments by validation", async () => {
    const doc = await Io.read(
      await docx(COMMENTED.replace('commentReference w:id="3"', 'commentReference w:id="8"'))
    );
    const issues = Validation.validate(doc).issues.map(i => i.rule);
    expect(issues).toContain("ref-comment-missing");
  });

  it("can be built with its formatting", async () => {
    const doc = Document.create();
    const id = Document.addComment(doc, "A", "c");
    Document.addContent(
      doc,
      Build.paragraph([
        Build.commentRangeStart(id),
        Build.text("x"),
        Build.commentRangeEnd(id),
        { ...Build.commentReference(id), properties: { style: "CommentReference" } }
      ])
    );
    expect(await anchors(doc)).toEqual({ comments: ["0"], references: ["0"] });
    const xml = (await parts(await Io.toBuffer(doc))).get("word/document.xml")!;
    expect(xml).toContain('<w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference');
  });
});
