/**
 * Placeholders are filled wherever the template can hold them: rich
 * placeholders resolve against the loop item they are repeated for, comment
 * bodies are filled, and removing directive paragraphs, untaken branches or
 * empty loops never separates a comment from its reference.
 */
import { unzip } from "@archive/read-archive";
import { describe, it, expect } from "vitest";

import { Build, Document, Io, Template } from "../../index";
import type {
  DocxDocument,
  InlineImageContent,
  Paragraph,
  ParagraphChild,
  TemplateImage
} from "../../index";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

function picture(fileName: string): TemplateImage {
  return {
    image: { data: PNG, fileName, mediaType: "png" },
    width: 9525,
    height: 9525
  };
}

function loopTemplate(placeholder: string): DocxDocument {
  const doc = Document.create();
  Document.addParagraph(doc, "{{#each items}}");
  Document.addParagraph(doc, placeholder);
  Document.addParagraph(doc, "{{/each}}");
  return doc;
}

/** The inline images in `blocks`, in order, by name. */
function imageNames(blocks: readonly unknown[]): string[] {
  const out: string[] = [];
  const visit = (node: unknown): void => {
    if (Array.isArray(node)) {
      node.forEach(visit);
    } else if (node && typeof node === "object") {
      if ((node as { type?: string }).type === "image") {
        out.push((node as InlineImageContent).name ?? "");
      }
      Object.values(node).forEach(visit);
    }
  };
  visit(blocks);
  return out;
}

function text(para: Paragraph): string {
  return para.children
    .flatMap(c => ("content" in c ? c.content : []))
    .map(c => (c.type === "text" ? c.text : ""))
    .join("");
}

const items = [
  { pic: picture("a.png"), rich: [Build.text("one", { bold: true })] },
  { pic: picture("b.png"), rich: [Build.text("two", { italic: true })] }
];

describe("rich placeholders inside loops", () => {
  it("resolve {{%.pic}} against each item", () => {
    const filled = Template.fill(loopTemplate("{{%.pic}}"), { items }, { richPlaceholders: true });
    expect(imageNames(filled.body)).toEqual(["a.png", "b.png"]);
    expect(filled.images?.map(i => i.fileName)).toEqual(["a.png", "b.png"]);
  });

  it("resolve {{&.rich}} against each item, keeping its formatting", () => {
    const filled = Template.fill(loopTemplate("{{&.rich}}"), { items }, { richPlaceholders: true });
    const paras = filled.body.filter((b): b is Paragraph => b.type === "paragraph");
    expect(paras.map(text)).toEqual(["one", "two"]);
    const runs = paras.map(p => p.children[0] as { properties?: object });
    expect(runs.map(r => r.properties)).toEqual([{ bold: true }, { italic: true }]);
  });

  it("resolve in a repeated table row", () => {
    const doc = Document.create();
    Document.addTable(doc, {
      type: "table",
      rows: [
        {
          cells: [
            { content: [Build.textParagraph("{{#each items}}{{@index}}")] },
            { content: [Build.textParagraph("{{%.pic}}")] },
            { content: [Build.textParagraph("{{/each}}")] }
          ]
        }
      ]
    });
    const filled = Template.fill(doc, { items }, { richPlaceholders: true });
    expect(imageNames(filled.body)).toEqual(["a.png", "b.png"]);
  });

  it("resolve in the deprecated fillTemplateEnhanced too, instead of failing", () => {
    const filled = Template.fillTemplateEnhanced(loopTemplate("{{%.pic}}"), { items });
    expect(imageNames(filled.body)).toEqual(["a.png", "b.png"]);
  });

  it("are not re-read as template text once inserted", () => {
    const literal = [{ rich: [Build.text("{{not a placeholder}}")] }];
    const filled = Template.fill(
      loopTemplate("{{&.rich}}"),
      { items: literal },
      { richPlaceholders: true }
    );
    expect(text(filled.body.find(b => b.type === "paragraph") as Paragraph)).toBe(
      "{{not a placeholder}}"
    );
  });

  it("refuse a sub-document that inserts itself forever", () => {
    const sub: Paragraph[] = [Build.textParagraph("{{>sub}}")];
    const doc = Document.create();
    Document.addParagraph(doc, "{{>sub}}");
    expect(() => Template.fill(doc, { sub }, { richPlaceholders: true })).toThrow(/nest/);
  });
});

describe("comment bodies", () => {
  function withComment(): DocxDocument {
    const doc = Document.create();
    const id = Document.addComment(doc, "A", "Reviewed by {{reviewer}}");
    Document.addContent(
      doc,
      Build.paragraph([
        Build.commentRangeStart(id),
        Build.text("text"),
        Build.commentRangeEnd(id),
        Build.commentReference(id)
      ])
    );
    return doc;
  }

  it("are filled by Template.fill", () => {
    const filled = Template.fill(withComment(), { reviewer: "Ada" });
    expect(text(filled.comments![0]!.content[0]!)).toBe("Reviewed by Ada");
  });

  it("are left alone by the deprecated fillTemplate, as before", () => {
    const filled = Template.fillTemplate(withComment(), {});
    expect(text(filled.comments![0]!.content[0]!)).toBe("Reviewed by {{reviewer}}");
  });

  it("do not reach the template", () => {
    const tpl = withComment();
    Template.fill(tpl, { reviewer: "Ada" });
    expect(text(tpl.comments![0]!.content[0]!)).toBe("Reviewed by {{reviewer}}");
  });
});

/** Ids of the comments, the references to them and their range markers in a written package. */
async function commentAnchors(doc: DocxDocument) {
  const parts = new Map<string, string>();
  for await (const entry of unzip(await Io.toBuffer(doc)).entries()) {
    if (entry.path.endsWith(".xml")) {
      parts.set(entry.path, new TextDecoder().decode(await entry.bytes()));
    }
  }
  const ids = (xml: string | undefined, tag: string) =>
    [...(xml ?? "").matchAll(new RegExp(`<w:${tag} [^>]*?w:id="(\\d+)"`, "g"))].map(m => m[1]);
  const body = parts.get("word/document.xml");
  return {
    comments: ids(parts.get("word/comments.xml"), "comment"),
    references: ids(body, "commentReference"),
    starts: ids(body, "commentRangeStart"),
    ends: ids(body, "commentRangeEnd")
  };
}

describe("comments around removed template markup", () => {
  function commented(paragraphs: (ids: { id: number }) => ParagraphChild[][]): DocxDocument {
    const doc = Document.create();
    const id = Document.addComment(doc, "A", "c");
    for (const children of paragraphs({ id })) {
      Document.addContent(doc, Build.paragraph(children));
    }
    Document.addParagraph(doc, "after");
    return doc;
  }

  it("keep their reference when the range spans the loop's directive paragraphs", async () => {
    const doc = commented(({ id }) => [
      [Build.commentRangeStart(id), Build.text("{{#each items}}")],
      [Build.text("{{name}}")],
      [Build.text("{{/each}}"), Build.commentRangeEnd(id), Build.commentReference(id)]
    ]);
    const filled = Template.fill(doc, { items: [{ name: "a" }, { name: "b" }] });
    expect(await commentAnchors(filled)).toEqual({
      comments: ["0"],
      references: ["0"],
      starts: ["0"],
      ends: ["0"]
    });
  });

  it("keep their reference when the whole comment is inside an empty loop", async () => {
    const doc = commented(({ id }) => [
      [Build.text("{{#each none}}")],
      [
        Build.commentRangeStart(id),
        Build.text("z"),
        Build.commentRangeEnd(id),
        Build.commentReference(id)
      ],
      [Build.text("{{/each}}")]
    ]);
    const filled = Template.fill(doc, { none: [] });
    expect(await commentAnchors(filled)).toEqual({
      comments: ["0"],
      references: ["0"],
      starts: ["0"],
      ends: ["0"]
    });
  });

  it("keep their reference when it sits in an untaken branch", async () => {
    const doc = commented(({ id }) => [
      [Build.commentRangeStart(id), Build.text("kept")],
      [Build.text("{{#if flag}}")],
      [Build.text("hidden"), Build.commentRangeEnd(id), Build.commentReference(id)],
      [Build.text("{{/if}}")]
    ]);
    const filled = Template.fill(doc, { flag: false });
    expect(await commentAnchors(filled)).toEqual({
      comments: ["0"],
      references: ["0"],
      starts: ["0"],
      ends: ["0"]
    });
  });

  it("keep their reference when a rich placeholder replaces its paragraph", async () => {
    const doc = commented(({ id }) => [
      [
        Build.commentRangeStart(id),
        Build.text("{{%pic}}"),
        Build.commentRangeEnd(id),
        Build.commentReference(id)
      ]
    ]);
    const filled = Template.fill(doc, { pic: picture("p.png") }, { richPlaceholders: true });
    expect((await commentAnchors(filled)).references).toEqual(["0"]);
  });

  it("keep their reference when a row loop has no items", async () => {
    const doc = Document.create();
    const id = Document.addComment(doc, "A", "c");
    Document.addTable(doc, {
      type: "table",
      rows: [
        { cells: [{ content: [Build.textParagraph("H")] }] },
        {
          cells: [
            {
              content: [
                Build.paragraph([
                  Build.commentRangeStart(id),
                  Build.text("{{#each none}}{{name}}{{/each}}"),
                  Build.commentRangeEnd(id),
                  Build.commentReference(id)
                ])
              ]
            }
          ]
        }
      ]
    });
    const filled = Template.fill(doc, { none: [] });
    expect(await commentAnchors(filled)).toEqual({
      comments: ["0"],
      references: ["0"],
      starts: ["0"],
      ends: ["0"]
    });
  });
});

describe("listing tags", () => {
  it("lists comment-body tags with listTags, not with the deprecated listTemplateTags", () => {
    const doc = Document.create();
    Document.addComment(doc, "A", "{{reviewer}}");
    Document.addParagraph(doc, "{{name}}");
    expect(Template.listTags(doc).map(t => [t.expression, t.location])).toEqual([
      ["name", "body paragraph 0"],
      ["reviewer", "comment:0 paragraph 0"]
    ]);
    expect(Template.listTemplateTags(doc).map(t => t.expression)).toEqual(["name"]);
  });
});
