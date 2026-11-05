/**
 * A relationship id is local to the part whose `.rels` declares it, while the
 * model's image table is shared by the whole document. Two headers may each
 * call their own logo `rId1`; the body may use `rId5` for one picture while a
 * header uses `rId5` for another. Every picture must survive a round trip and
 * every reference must still resolve to the picture it named.
 */
import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { WordHtml } from "../html";
import { Build, Document, Io, Layout, Stream, Template, Transform } from "../index";
import type { DocxDocument } from "../index";
import { WordMarkdown } from "../markdown";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const REL_IMAGE = `${R}/image`;
const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** A distinct "PNG" per letter: the signature plus one identifying byte. */
function png(letter: string): Uint8Array {
  return new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, letter.charCodeAt(0)]);
}

const A = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
const PIC = `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;

function pic(rId: string, id: number): string {
  return (
    `<a:graphic ${A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic ${PIC}><pic:nvPicPr><pic:cNvPr id="${id}" name="P${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill><pic:spPr/></pic:pic>` +
    `</a:graphicData></a:graphic>`
  );
}

/**
 * A picture referencing `ref`. `ref` is a local relationship id, optionally
 * prefixed with the form it takes: `anchor:` (floating), `vml:` (a VML
 * shape's image fill) or `shape:` (a DrawingML shape filled with a picture —
 * kept as raw markup by the reader). Plain ids are inline pictures.
 */
function picture(ref: string, id: number): string {
  const [form, rId] = ref.includes(":") ? ref.split(":") : ["inline", ref];
  switch (form) {
    case "anchor":
      return (
        `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" ` +
        `behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
        `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
        `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
        `<wp:extent cx="9525" cy="9525"/><wp:wrapNone/><wp:docPr id="${id}" name="P${id}"/>` +
        pic(rId!, id) +
        `</wp:anchor></w:drawing></w:r>`
      );
    case "anchorShape":
      return (
        `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" ` +
        `behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
        `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
        `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
        `<wp:extent cx="9525" cy="9525"/><wp:wrapNone/><wp:docPr id="${id}" name="S${id}"/>` +
        `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
        `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/>` +
        `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>` +
        `<a:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>` +
        `<a:effectLst><a:outerShdw blurRad="1"/></a:effectLst></wps:spPr>` +
        `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
      );
    case "vml":
      return (
        `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="s${id}" style="width:9pt;height:9pt">` +
        `<v:imagedata r:id="${rId}" o:title="" xmlns:o="urn:schemas-microsoft-com:office:office"/></v:shape></w:pict></w:r>`
      );
    case "shape":
      return (
        `<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="${id}" name="S${id}"/>` +
        `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
        `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/>` +
        `<wps:spPr><a:prstGeom prst="rect"/><a:blipFill><a:blip r:embed="${rId}"/></a:blipFill></wps:spPr>` +
        `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
      );
    default:
      return `<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="${id}" name="P${id}"/>${pic(rId!, id)}</wp:inline></w:drawing></w:r>`;
  }
}

function para(inner: string): string {
  return `<w:p>${inner}</w:p>`;
}

function rels(entries: [id: string, type: string, target: string][]): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    entries
      .map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`)
      .join("") +
    `</Relationships>`
  );
}

interface PartSpec {
  /** Path inside `word/`, e.g. `header1.xml`. */
  readonly name: string;
  /** Relationship type from the document to the part. */
  readonly type: string;
  readonly root: string;
  readonly contentType: string;
  /** The part's own image relationships: local id → media file. */
  readonly images: Record<string, string>;
  /** Pictures in the part, by local id. */
  readonly pictures: readonly string[];
}

const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}"`;

function header(name: string, images: Record<string, string>, pictures: string[]): PartSpec {
  return {
    name,
    type: `${R}/header`,
    root: "w:hdr",
    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    images,
    pictures
  };
}

function footer(name: string, images: Record<string, string>, pictures: string[]): PartSpec {
  return {
    name,
    type: `${R}/footer`,
    root: "w:ftr",
    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    images,
    pictures
  };
}

/**
 * A minimal package: a body with pictures, its own image relationships, and
 * any number of header/footer/note/comment parts each with their own.
 */
async function buildPackage(options: {
  body: { images: Record<string, string>; pictures: string[] };
  parts?: readonly PartSpec[];
  notes?: {
    kind: "footnotes" | "endnotes" | "comments";
    images: Record<string, string>;
    pictures: string[];
  }[];
  numbering?: { images: Record<string, string>; bullets: string[] };
  glossary?: { images: Record<string, string>; pictures: string[] };
  media: Record<string, Uint8Array>;
}): Promise<Uint8Array> {
  const archive = zip();
  const overrides: string[] = [];
  const docRels: [string, string, string][] = [];
  let n = 100;
  let pid = 1;
  const sectRefs: string[] = [];
  for (const part of options.parts ?? []) {
    const relId = `rIdPart${n++}`;
    docRels.push([relId, part.type, part.name]);
    const kind = part.root === "w:hdr" ? "headerReference" : "footerReference";
    const refType =
      sectRefs.filter(r => r.startsWith(`<w:${kind}`)).length === 0 ? "default" : "first";
    sectRefs.push(`<w:${kind} w:type="${refType}" r:id="${relId}"/>`);
    archive.add(
      `word/${part.name}`,
      encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${part.root} ${NS}>` +
          part.pictures.map(id => para(picture(id, pid++))).join("") +
          `</${part.root}>`
      )
    );
    archive.add(
      `word/_rels/${part.name}.rels`,
      encoder.encode(
        rels(Object.entries(part.images).map(([id, file]) => [id, REL_IMAGE, `media/${file}`]))
      )
    );
    overrides.push(`<Override PartName="/word/${part.name}" ContentType="${part.contentType}"/>`);
  }
  const bodyRefs: string[] = [];
  for (const note of options.notes ?? []) {
    const single = note.kind === "comments" ? "comment" : note.kind.slice(0, -1);
    const relId = `rIdPart${n++}`;
    docRels.push([relId, `${R}/${note.kind}`, `${note.kind}.xml`]);
    const content = note.pictures.map(id => para(picture(id, pid++))).join("");
    const entries =
      note.kind === "comments"
        ? `<w:comment w:id="0" w:author="a">${content}</w:comment>`
        : `<w:${single} w:id="1">${content}</w:${single}>`;
    archive.add(
      `word/${note.kind}.xml`,
      encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${note.kind} ${NS}>${entries}</w:${note.kind}>`
      )
    );
    archive.add(
      `word/_rels/${note.kind}.xml.rels`,
      encoder.encode(
        rels(Object.entries(note.images).map(([id, file]) => [id, REL_IMAGE, `media/${file}`]))
      )
    );
    overrides.push(
      `<Override PartName="/word/${note.kind}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${note.kind}+xml"/>`
    );
    bodyRefs.push(
      note.kind === "comments"
        ? `<w:r><w:commentReference w:id="0"/></w:r>`
        : `<w:r><w:${single}Reference w:id="1"/></w:r>`
    );
  }
  if (options.numbering) {
    const relId = `rIdPart${n++}`;
    docRels.push([relId, `${R}/numbering`, "numbering.xml"]);
    const bullets = options.numbering.bullets
      .map(
        (rId, i) =>
          `<w:numPicBullet w:numPicBulletId="${i}"><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" ` +
          `id="_x0000_i10${i}" style="width:9pt;height:9pt"><v:imagedata r:id="${rId}" o:title="" ` +
          `xmlns:o="urn:schemas-microsoft-com:office:office"/></v:shape></w:pict></w:numPicBullet>`
      )
      .join("");
    archive.add(
      "word/numbering.xml",
      encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>${bullets}` +
          `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
          `<w:lvlText w:val="\uF0B7"/><w:lvlPicBulletId w:val="0"/></w:lvl></w:abstractNum>` +
          `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
      )
    );
    archive.add(
      "word/_rels/numbering.xml.rels",
      encoder.encode(
        rels(
          Object.entries(options.numbering.images).map(([id, file]) => [
            id,
            REL_IMAGE,
            `media/${file}`
          ])
        )
      )
    );
    overrides.push(
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`
    );
  }
  if (options.glossary) {
    docRels.push([`rIdPart${n++}`, `${R}/glossaryDocument`, "glossary/document.xml"]);
    archive.add(
      "word/glossary/document.xml",
      encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument ${NS}><w:docParts><w:docPart>` +
          `<w:docPartPr><w:name w:val="G"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category></w:docPartPr>` +
          `<w:docPartBody>${options.glossary.pictures.map(id => para(picture(id, pid++))).join("")}</w:docPartBody>` +
          `</w:docPart></w:docParts></w:glossaryDocument>`
      )
    );
    archive.add(
      "word/glossary/_rels/document.xml.rels",
      encoder.encode(
        rels(
          Object.entries(options.glossary.images).map(([id, file]) => [
            id,
            REL_IMAGE,
            `media/${file}`
          ])
        )
      )
    );
    overrides.push(
      `<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>`
    );
  }
  for (const [id, file] of Object.entries(options.body.images)) {
    docRels.push([id, REL_IMAGE, `media/${file}`]);
  }
  archive.add(
    "word/document.xml",
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
        options.body.pictures.map(id => para(picture(id, pid++))).join("") +
        para(bodyRefs.join("")) +
        `<w:sectPr>${sectRefs.join("")}</w:sectPr></w:body></w:document>`
    )
  );
  archive.add("word/_rels/document.xml.rels", encoder.encode(rels(docRels)));
  for (const [file, data] of Object.entries(options.media)) {
    archive.add(file.includes("/") ? `word/${file}` : `word/media/${file}`, data);
  }
  archive.add(
    "_rels/.rels",
    encoder.encode(
      rels([
        [
          "rId1",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
          "word/document.xml"
        ]
      ])
    )
  );
  archive.add(
    "[Content_Types].xml",
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        overrides.join("") +
        `</Types>`
    )
  );
  return archive.toBuffer();
}

async function entries(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for await (const entry of unzip(bytes).entries()) {
    out.set(entry.path.replace(/^\/+/, ""), await entry.bytes());
  }
  return out;
}

/**
 * For every part of a written package that references pictures, the letters
 * of the pictures it shows, in document order — resolved the way Word would,
 * through that part's own `.rels`. A reference that does not resolve is "?".
 */
async function picturesByPart(bytes: Uint8Array): Promise<Map<string, string[]>> {
  const files = await entries(bytes);
  const out = new Map<string, string[]>();
  for (const [path, data] of files) {
    if (!/^word\/(?:glossary\/)?[^/]+\.xml$/.test(path)) {
      continue;
    }
    const xml = decoder.decode(data);
    const ids = [
      ...xml.matchAll(/<(?:a:blip\b[^>]*\br:embed|v:imagedata\b[^>]*\br:id)="([^"]+)"/g)
    ].map(m => m[1]!);
    if (ids.length === 0) {
      continue;
    }
    const name = path.slice("word/".length);
    const dir = path.slice(0, path.lastIndexOf("/") + 1);
    const relsXml = files.get(`${dir}_rels/${path.slice(dir.length)}.rels`);
    const targets = new Map<string, string>();
    for (const m of relsXml ? decoder.decode(relsXml).matchAll(/<Relationship\b[^>]*>/g) : []) {
      const id = /\bId="([^"]+)"/.exec(m[0])?.[1];
      const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1];
      if (id && target) {
        targets.set(id, target);
      }
    }
    out.set(
      name,
      ids.map(id => {
        const target = targets.get(id);
        const media = target ? files.get(`${dir}${target}`) : undefined;
        return media ? String.fromCharCode(media[media.length - 1]!) : "?";
      })
    );
  }
  return out;
}

async function roundTrip(bytes: Uint8Array): Promise<Uint8Array> {
  return Io.toBuffer(await Io.read(bytes));
}

describe("image references stay bound to their own part's relationships", () => {
  it("two headers that each call a different logo rId1", async () => {
    const source = await buildPackage({
      body: { images: {}, pictures: [] },
      parts: [
        header("header1.xml", { rId1: "a.png" }, ["rId1"]),
        header("header2.xml", { rId1: "b.png" }, ["rId1"])
      ],
      media: { "a.png": png("a"), "b.png": png("b") }
    });
    expect(await picturesByPart(source)).toEqual(
      new Map([
        ["header1.xml", ["a"]],
        ["header2.xml", ["b"]]
      ])
    );
    const written = await picturesByPart(await roundTrip(source));
    expect([...written.values()].map(v => v.join("")).sort()).toEqual(["a", "b"]);
    // And again: a second round trip is stable.
    const twice = await picturesByPart(await roundTrip(await roundTrip(source)));
    expect([...twice.values()].map(v => v.join("")).sort()).toEqual(["a", "b"]);
  });

  it("a header and the body that use the same rId for different pictures", async () => {
    const source = await buildPackage({
      body: { images: { rId5: "body.png" }, pictures: ["rId5"] },
      parts: [header("header1.xml", { rId5: "head.png" }, ["rId5"])],
      media: { "body.png": png("B"), "head.png": png("H") }
    });
    const written = await picturesByPart(await roundTrip(source));
    expect(written.get("document.xml")).toEqual(["B"]);
    expect(
      [...written.entries()].filter(([k]) => k.startsWith("header")).map(([, v]) => v)
    ).toEqual([["H"]]);
  });

  it("a header and the body sharing one file under different ids", async () => {
    const source = await buildPackage({
      body: { images: { rId7: "same.png" }, pictures: ["rId7"] },
      parts: [header("header1.xml", { rId1: "same.png" }, ["rId1"])],
      media: { "same.png": png("S") }
    });
    const read = await Io.read(source);
    expect(read.images).toHaveLength(1);
    const written = await picturesByPart(await Io.toBuffer(read));
    expect(written.get("document.xml")).toEqual(["S"]);
    expect([...written.values()].flat()).toEqual(["S", "S"]);
  });

  it("layout, HTML and Markdown resolve a picture named by an alias", async () => {
    const source = await buildPackage({
      body: { images: { rId7: "same.png" }, pictures: ["rId7"] },
      parts: [header("header1.xml", { rId1: "same.png" }, ["rId1"])],
      media: { "same.png": png("S") }
    });
    const read = await Io.read(source);
    const headerContent = [...read.headers!.values()][0]!.content;
    const alias = read.images![0]!.aliasRIds?.[0];
    expect(alias).toBe("rId1");

    const page = Layout.compute(read).pages[0]!;
    const headerImages = (page.header ?? []).flatMap(item =>
      item.type === "paragraph"
        ? item.lines.flatMap(l => l.runs).filter(r => r.type === "image")
        : []
    );
    expect(headerImages).toHaveLength(1);

    // A header's content rendered on its own, through the converters.
    const asBody: DocxDocument = { ...read, body: headerContent.children };
    expect(WordHtml.fromDocx(asBody, { imageMode: "filename" }).html).toContain('src="same.png"');
    expect(WordMarkdown.fromDocx(asBody)).toContain("same.png");
  });

  it("footers, footnotes, endnotes and comments with colliding ids", async () => {
    const source = await buildPackage({
      body: { images: { rId1: "body.png" }, pictures: ["rId1"] },
      parts: [
        footer("footer1.xml", { rId1: "foot.png" }, ["rId1"]),
        header("header1.xml", { rId1: "head.png" }, ["rId1"])
      ],
      notes: [
        { kind: "footnotes", images: { rId1: "fn.png" }, pictures: ["rId1"] },
        { kind: "endnotes", images: { rId1: "en.png" }, pictures: ["rId1"] },
        { kind: "comments", images: { rId1: "cm.png" }, pictures: ["rId1"] }
      ],
      media: {
        "body.png": png("B"),
        "foot.png": png("F"),
        "head.png": png("H"),
        "fn.png": png("N"),
        "en.png": png("E"),
        "cm.png": png("C")
      }
    });
    expect([...(await picturesByPart(source)).values()].flat().sort()).toEqual([
      "B",
      "C",
      "E",
      "F",
      "H",
      "N"
    ]);
    const written = await picturesByPart(await roundTrip(source));
    expect(written.get("document.xml")).toEqual(["B"]);
    expect(written.get("footnotes.xml")).toEqual(["N"]);
    expect(written.get("endnotes.xml")).toEqual(["E"]);
    expect(written.get("comments.xml")).toEqual(["C"]);
    expect([...written.values()].flat().sort()).toEqual(["B", "C", "E", "F", "H", "N"]);
  });

  it("floating pictures, VML image fills and picture-filled shapes in headers", async () => {
    for (const form of ["anchor", "vml", "shape"]) {
      const source = await buildPackage({
        body: { images: { rId1: "body.png" }, pictures: ["rId1"] },
        parts: [
          header("header1.xml", { rId1: "a.png" }, [`${form}:rId1`]),
          header("header2.xml", { rId1: "b.png" }, [`${form}:rId1`])
        ],
        media: { "body.png": png("B"), "a.png": png("a"), "b.png": png("b") }
      });
      expect([...(await picturesByPart(source)).values()].flat().sort(), form).toEqual([
        "B",
        "a",
        "b"
      ]);
      const written = await picturesByPart(await roundTrip(source));
      expect(written.get("document.xml"), form).toEqual(["B"]);
      expect([...written.values()].flat().sort(), form).toEqual(["B", "a", "b"]);
    }
  });

  it("numbering picture bullets", async () => {
    const source = await buildPackage({
      body: { images: { rId1: "body.png" }, pictures: ["rId1"] },
      numbering: { images: { rId1: "bullet.png" }, bullets: ["rId1"] },
      media: { "body.png": png("B"), "bullet.png": png("u") }
    });
    expect((await picturesByPart(source)).get("numbering.xml")).toEqual(["u"]);
    const written = await picturesByPart(await roundTrip(source));
    expect(written.get("document.xml")).toEqual(["B"]);
    expect(written.get("numbering.xml")).toEqual(["u"]);
  });

  it("a glossary with its own rId1", async () => {
    const source = await buildPackage({
      body: { images: { rId1: "body.png" }, pictures: ["rId1"] },
      glossary: { images: { rId1: "g.png" }, pictures: ["rId1"] },
      media: { "body.png": png("B"), "glossary/media/g.png": png("g") }
    });
    expect((await picturesByPart(source)).get("glossary/document.xml")).toEqual(["g"]);
    const written = await picturesByPart(await roundTrip(source));
    expect(written.get("document.xml")).toEqual(["B"]);
    expect(written.get("glossary/document.xml")).toEqual(["g"]);
  });

  it("merging two documents that both carry such headers keeps every picture", async () => {
    // Both documents call their body picture rId5 / image1.png, and their two
    // header logos rId1 — every one a different picture.
    const make = (x: string, y: string, z: string) =>
      buildPackage({
        body: { images: { rId5: "image1.png" }, pictures: ["rId5"] },
        parts: [
          header("header1.xml", { rId1: `${x}.png` }, ["rId1"]),
          header("header2.xml", { rId1: `${y}.png` }, ["rId1"])
        ],
        media: { "image1.png": png(z), [`${x}.png`]: png(x), [`${y}.png`]: png(y) }
      });
    const first = await Io.read(await make("a", "b", "c"));
    const second = await Io.read(await make("d", "e", "f"));
    const merged = Transform.merge([first, second]);
    const written = await picturesByPart(await Io.toBuffer(merged));
    expect(written.get("document.xml")).toEqual(["c", "f"]);
    // Both documents' headers are kept, each with its own logo.
    const headers = [...written.entries()].filter(([k]) => k.startsWith("header"));
    expect(headers.map(([, v]) => v.join("")).sort()).toEqual(["a", "b", "d", "e"]);

    // Merging a document with itself adds nothing new: every key already
    // names the same picture.
    const self = Transform.merge([first, first]);
    expect(self.images).toHaveLength(first.images!.length);
    expect((await picturesByPart(await Io.toBuffer(self))).get("document.xml")).toEqual(["c", "c"]);
  });

  it("merging renames image references in notes, comments and verbatim markup", async () => {
    const make = (n: string, c: string, s: string) =>
      buildPackage({
        body: { images: { rId1: `${s}.png` }, pictures: ["shape:rId1"] },
        notes: [
          { kind: "footnotes", images: { rId1: `${n}.png` }, pictures: ["rId1"] },
          { kind: "comments", images: { rId1: `${c}.png` }, pictures: ["rId1"] }
        ],
        media: { [`${n}.png`]: png(n), [`${c}.png`]: png(c), [`${s}.png`]: png(s) }
      });
    const merged = Transform.merge([
      await Io.read(await make("n", "c", "s")),
      await Io.read(await make("N", "C", "S"))
    ]);
    const written = await picturesByPart(await Io.toBuffer(merged));
    expect(written.get("document.xml")).toEqual(["s", "S"]);
    expect(written.get("footnotes.xml")?.sort()).toEqual(["N", "n"]);
    expect(written.get("comments.xml")?.sort()).toEqual(["C", "c"]);
  });

  it("a minimal split keeps the pictures its body shows", async () => {
    const doc = Document.create();
    Document.addImage(doc, png("p"), "png", 9525, 9525);
    Document.addSectionBreak(doc, {});
    Document.addImage(doc, png("q"), "png", 9525, 9525);
    const parts = Transform.split(doc, { preserveSharedParts: false });
    expect(parts).toHaveLength(2);
    const shown = await Promise.all(
      parts.map(async p => (await picturesByPart(await Io.toBuffer(p))).get("document.xml"))
    );
    expect(shown).toEqual([["p"], ["q"]]);
    expect(parts.map(p => p.images?.length)).toEqual([1, 1]);
  });

  it("template images keep their own picture beside the template's", async () => {
    const doc = Document.create();
    Document.addImage(doc, png("t"), "png", 9525, 9525);
    // A caller's pictures: one named like the template's own image1.png, one
    // carrying an rId of its own that the template already uses for another
    // picture, and two different pictures under one file name.
    const templateKey = doc.images![0]!.rId!;
    Document.addParagraph(doc, "{{%a}}");
    Document.addParagraph(doc, "{{%b}}");
    Document.addParagraph(doc, "{{%c}}");
    Document.addParagraph(doc, "{{%d}}");
    const image = (letter: string, fileName: string, rId?: string) => ({
      image: { data: png(letter), mediaType: "png" as const, fileName, ...(rId ? { rId } : {}) },
      width: 9525,
      height: 9525
    });
    const filled = Template.fill(
      doc,
      {
        a: image("a", "image1.png"),
        b: image("b", "b.png", templateKey),
        c: image("c", "same.png"),
        d: image("d", "same.png")
      },
      { richPlaceholders: true }
    );
    const written = await picturesByPart(await Io.toBuffer(filled));
    expect(written.get("document.xml")).toEqual(["t", "a", "b", "c", "d"]);

    // Supplying the template's own picture under its own key reuses it.
    const own = Document.create();
    Document.addImage(own, png("t"), "png", 9525, 9525);
    Document.addParagraph(own, "{{%a}}");
    const reuse = Template.fill(
      own,
      { a: { image: own.images![0]!, width: 9525, height: 9525 } },
      { richPlaceholders: true }
    );
    expect(reuse.images).toHaveLength(1);
    expect((await picturesByPart(await Io.toBuffer(reuse))).get("document.xml")).toEqual([
      "t",
      "t"
    ]);
  });

  it("pictures imported from two HTML or Markdown fragments do not collide", async () => {
    const dataUrl = (letter: string) =>
      `data:image/png;base64,${btoa(String.fromCharCode(...png(letter)))}`;
    const first = WordHtml.toDocxBody(`<p><img src="${dataUrl("h")}" width="1" height="1"></p>`);
    const second = WordHtml.toDocxBody(`<p><img src="${dataUrl("i")}" width="1" height="1"></p>`);
    const html: DocxDocument = {
      ...Document.create(),
      body: [...first.body, ...second.body],
      images: [...first.images, ...second.images]
    };
    expect((await picturesByPart(await Io.toBuffer(html))).get("document.xml")).toEqual(["h", "i"]);

    const md = async (letter: string) =>
      WordMarkdown.toDocxBody("![x](x.png)", {
        resolveImage: () => ({ data: png(letter), mediaType: "png" })
      });
    const m1 = await md("m");
    const m2 = await md("n");
    const markdown: DocxDocument = {
      ...Document.create(),
      body: [...m1.body, ...m2.body],
      images: [...m1.images, ...m2.images]
    };
    expect((await picturesByPart(await Io.toBuffer(markdown))).get("document.xml")).toEqual([
      "m",
      "n"
    ]);
  });

  it("a body shape filled with a picture keeps its fill and effects", async () => {
    const make = (letter: string) =>
      buildPackage({
        body: { images: { rId1: `${letter}.png` }, pictures: ["anchorShape:rId1"] },
        parts: [header("header1.xml", { rId1: "h.png" }, ["anchorShape:rId1"])],
        media: { [`${letter}.png`]: png(letter), "h.png": png("h") }
      });
    const read = await Io.read(await make("f"));
    const shape = read.body.find(b => b.type === "drawingShape");
    expect(shape?.type).toBe("drawingShape");
    const written = await Io.toBuffer(read);
    expect((await picturesByPart(written)).get("document.xml")).toEqual(["f"]);
    const xml = decoder.decode((await entries(written)).get("word/document.xml")!);
    expect(xml).toContain('prst="ellipse"');
    expect(xml).toContain("<a:stretch>");
    expect(xml).toContain("<a:outerShdw");

    // The fill's image key takes part in merging like any other reference.
    const merged = Transform.merge([read, await Io.read(await make("g"))]);
    expect((await picturesByPart(await Io.toBuffer(merged))).get("document.xml")).toEqual([
      "f",
      "g"
    ]);
  });

  it("every drawing in the package has its own docPr id", async () => {
    // Source ids are unique; the writer renumbers modelled drawings from 1, so
    // markup it keeps verbatim must be renumbered with them.
    const source = await buildPackage({
      body: {
        images: { rId1: "a.png" },
        pictures: ["rId1", "shape:rId1", "rId1", "anchorShape:rId1"]
      },
      parts: [
        header("header1.xml", { rId1: "a.png" }, ["anchor:rId1", "rId1"]),
        footer("footer1.xml", { rId1: "a.png" }, ["shape:rId1"])
      ],
      notes: [
        { kind: "footnotes", images: { rId1: "a.png" }, pictures: ["anchor:rId1", "rId1"] },
        { kind: "comments", images: { rId1: "a.png" }, pictures: ["shape:rId1"] }
      ],
      media: { "a.png": png("a") }
    });
    const files = await entries(await roundTrip(source));
    const ids: string[] = [];
    for (const [path, data] of files) {
      if (/^word\/[^/]+\.xml$/.test(path)) {
        ids.push(
          ...[...decoder.decode(data).matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)].map(m => m[1]!)
        );
      }
    }
    expect(ids).toHaveLength(10);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("the streaming writer gives each part its own image relationships", async () => {
    const a = { data: png("a"), mediaType: "png" as const, fileName: "a.png", rId: "__img_1" };
    const b = { data: png("b"), mediaType: "png" as const, fileName: "b.png", rId: "__img_2" };
    const show = (rId: string) => [
      Build.paragraph([{ content: [{ type: "image" as const, rId, width: 9525, height: 9525 }] }])
    ];
    const stream = new Stream.DocumentWriter({
      images: [a, b],
      headers: new Map([
        ["h1", { content: { children: show("__img_2") } }],
        ["h2", { content: { children: show("__img_1") } }]
      ]),
      footnotes: [{ id: 1, content: show("__img_1") }],
      endnotes: [{ id: 1, content: show("__img_2") }],
      comments: [{ id: 0, author: "x", content: show("__img_1") }]
    });
    stream.add(show("__img_1")[0]!);
    const written = await picturesByPart(await stream.finalize());
    expect(written.get("header1.xml")).toEqual(["b"]);
    expect(written.get("header2.xml")).toEqual(["a"]);
    expect(written.get("footnotes.xml")).toEqual(["a"]);
    expect(written.get("endnotes.xml")).toEqual(["b"]);
    expect(written.get("comments.xml")).toEqual(["a"]);
  });

  it("builder images placed in a header resolve under a canonical relationship id", async () => {
    const doc = Document.create();
    const { rId } = Document.addImage(doc, png("x"), "png", 9525, 9525);
    Document.setHeader(doc, "default", {
      children: [Build.paragraph([{ content: [{ type: "image", rId, width: 1, height: 1 }] }])]
    });
    const files = await entries(await Io.toBuffer(doc));
    const written = await picturesByPart(await Io.toBuffer(doc));
    expect(written.get("document.xml")).toEqual(["x"]);
    expect([...written.values()].flat()).toEqual(["x", "x"]);
    for (const [path, data] of files) {
      if (path.endsWith(".rels")) {
        for (const m of decoder.decode(data).matchAll(/\bId="([^"]+)"/g)) {
          expect(m[1]).toMatch(/^rId\d+$/);
        }
      }
    }
  });
});
