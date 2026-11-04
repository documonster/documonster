/**
 * The `wp:anchor` / `wp:inline` frame of a modelled drawing — distances,
 * z-order and locking flags, positions, wrapping and its polygon, effect
 * extent, `wp14` sizes and ids, `docPr` title / hidden / hyperlink, frame
 * locks — survives a round trip as Word wrote it, for pictures, shapes and
 * charts alike. A modelled field that is edited replaces what it was read
 * from.
 */
import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { parsePartXml } from "@word/reader/part-xml";
import type { XmlElement } from "@xml/types";
import { describe, expect, it } from "vitest";

import { Build, Document, Io, Streaming } from "../index";
import type { BodyContent, DocxDocument, FloatingImage } from "../types";

const decoder = new TextDecoder();
const encoder = new TextEncoder();
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 7]);

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';

const HYPERLINK = "https://example.com/logo";

function pictureGraphic(rId: string): string {
  return (
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Logo"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`
  );
}

/** A picture wrapped tightly round a polygon, centred on the margin, with a hyperlink. */
function tightPicture(rId: string, linkId: string): string {
  return (
    `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" ` +
    `behindDoc="0" locked="1" layoutInCell="1" hidden="0" allowOverlap="0" wp14:anchorId="1A2B3C4D" wp14:editId="5E6F7A8B">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>12700</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="952500" cy="952500"/><wp:effectExtent l="1" t="2" r="3" b="4"/>` +
    `<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="1"><wp:start x="0" y="0"/>` +
    `<wp:lineTo x="0" y="21000"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/>` +
    `<wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>` +
    `<wp:docPr id="5" name="Logo" descr="Company logo" title="Logo title">` +
    `<a:hlinkClick r:id="${linkId}" tooltip="Visit"/></wp:docPr>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1" noSelect="1"/></wp:cNvGraphicFramePr>` +
    pictureGraphic(rId) +
    `<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH>` +
    `<wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV>` +
    `</wp:anchor>`
  );
}

/** A shape behind the text, positioned against the page, sized against the margin. */
const BEHIND_SHAPE =
  `<wp:anchor distT="45720" distB="45720" distL="0" distR="0" simplePos="0" relativeHeight="251660288" ` +
  `behindDoc="1" locked="0" layoutInCell="0" allowOverlap="1" wp14:anchorId="0F0F0F0F">` +
  `<wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="page"><wp:posOffset>-457200</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="margin"><wp:align>bottom</wp:align></wp:positionV>` +
  `<wp:extent cx="7772400" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
  `<wp:docPr id="9" name="Band" hidden="1"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="7772400" cy="914400"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>` +
  `</a:graphicData></a:graphic>` +
  `<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>100000</wp14:pctWidth></wp14:sizeRelH>` +
  `<wp14:sizeRelV relativeFrom="page"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV>` +
  `</wp:anchor>`;

/** An inline picture with spacing, an effect extent and frame locks. */
function inlinePicture(rId: string): string {
  return (
    `<wp:inline distT="12700" distB="0" distL="0" distR="0" wp14:anchorId="11112222" wp14:editId="33334444">` +
    `<wp:extent cx="952500" cy="952500"/><wp:effectExtent l="19050" t="19050" r="19050" b="19050"/>` +
    `<wp:docPr id="3" name="Inline" descr="An inline picture" title="Inline title"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    pictureGraphic(rId) +
    `</wp:inline>`
  );
}

async function entries(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for await (const entry of unzip(bytes).entries()) {
    out.set(entry.path.replace(/^\/+/, ""), await entry.bytes());
  }
  return out;
}

async function pack(files: Map<string, Uint8Array>): Promise<Uint8Array> {
  const archive = zip();
  for (const [path, data] of files) {
    archive.add(path, data);
  }
  return archive.toBuffer();
}

function para(drawing: string): string {
  return `<w:p><w:r><w:drawing>${drawing}</w:drawing></w:r></w:p>`;
}

/**
 * A package written by the library — so its parts and relationships are
 * valid — with its body replaced by `drawings(imageRId, linkRId)`, and its
 * footnote holding the inline picture too.
 */
async function fixture(drawings: (image: string, link: string) => string): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addImage(doc, PNG, "png", 952500, 952500);
  Document.addFootnote(doc, "note");
  const files = await entries(await Io.toBuffer(doc));
  const rels = decoder.decode(files.get("word/_rels/document.xml.rels")!);
  const image = [...targets(files, "word/document.xml")].find(([, t]) =>
    t.startsWith("media/")
  )![0];
  files.set(
    "word/_rels/document.xml.rels",
    encoder.encode(
      rels.replace(
        "</Relationships>",
        `<Relationship Id="rIdLink9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${HYPERLINK}" TargetMode="External"/></Relationships>`
      )
    )
  );
  files.set(
    "word/document.xml",
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
        drawings(image, "rIdLink9") +
        `<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>`
    )
  );
  return pack(files);
}

function frames(root: XmlElement, out: XmlElement[] = []): XmlElement[] {
  for (const child of root.children) {
    if (child.type !== "element") {
      continue;
    }
    if (child.name === "wp:anchor" || child.name === "wp:inline") {
      out.push(child);
    } else {
      frames(child, out);
    }
  }
  return out;
}

/** Relationship id → target of a part's `.rels`. */
function targets(files: Map<string, Uint8Array>, part: string): Map<string, string> {
  const out = new Map<string, string>();
  const rels = files.get(part.replace(/([^/]+)$/, "_rels/$1.rels"));
  for (const m of rels ? decoder.decode(rels).matchAll(/<Relationship\b[^>]*>/g) : []) {
    const id = /\bId="([^"]+)"/.exec(m[0])?.[1];
    const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1];
    if (id && target) {
      out.set(id, target);
    }
  }
  return out;
}

/**
 * Canonical form of a frame: names, sorted attributes, children in order,
 * namespace declarations dropped. `a:graphic` is the drawing itself, not its
 * frame, so it is reduced to its URI; the drawing id is the writer's to
 * assign; a relationship id is replaced by what it points at.
 */
function canonical(el: XmlElement, rels: Map<string, string>): string {
  if (el.name === "a:graphic") {
    const data = el.children.find((c): c is XmlElement => c.type === "element");
    return `<a:graphic uri="${data?.attributes["uri"]}"/>`;
  }
  const attrs = Object.entries(el.attributes)
    .filter(([k]) => !k.startsWith("xmlns") && !(el.name === "wp:docPr" && k === "id"))
    .map(([k, v]) => [k, k === "r:id" ? `→${rels.get(v) ?? "?"}` : v] as const)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([k, v]) => ` ${k}="${v}"`)
    .join("");
  const kids = el.children
    .map(c => (c.type === "element" ? canonical(c, rels) : c.type === "text" ? c.value : ""))
    .join("");
  return `<${el.name}${attrs}>${kids}</${el.name}>`;
}

async function canonicalFrames(bytes: Uint8Array, part = "word/document.xml"): Promise<string[]> {
  const files = await entries(bytes);
  const rels = targets(files, part);
  return frames(parsePartXml(decoder.decode(files.get(part)!), undefined).root).map(f =>
    canonical(f, rels)
  );
}

describe("drawing frames survive a round trip", () => {
  it("a tightly wrapped picture, a behind-text shape and an inline picture", async () => {
    const source = await fixture(
      (image, link) =>
        para(tightPicture(image, link)) + para(BEHIND_SHAPE) + para(inlinePicture(image))
    );
    const before = await canonicalFrames(source);
    expect(before).toHaveLength(3);

    const read = await Io.read(source);
    const kinds = read.body.map(b => b.type);
    expect(kinds).toContain("floatingImage");
    expect(kinds).toContain("drawingShape");

    const once = await Io.toBuffer(read);
    expect(await canonicalFrames(once)).toEqual(before);
    expect(await canonicalFrames(await Io.toBuffer(await Io.read(once)))).toEqual(before);
  });

  it("a picture's hyperlink is registered in the part it is written into, by either writer", async () => {
    const source = await fixture((image, link) => para(tightPicture(image, link)));
    const [before] = await canonicalFrames(source);
    expect(before).toContain(`<a:hlinkClick r:id="→${HYPERLINK}" tooltip="Visit">`);
    const read = await Io.read(source);

    // Two copies in one document: each frame points at a relationship of the part.
    const doubled: DocxDocument = { ...read, body: [...read.body, ...read.body] };
    const twice = await canonicalFrames(await Io.toBuffer(doubled));
    expect(twice).toEqual([before, before]);

    const stream = Streaming.createWriter({ images: read.images ? [...read.images] : [] });
    for (const block of read.body) {
      stream.add(block);
    }
    const [streamed] = await canonicalFrames(await stream.finalize());
    expect(streamed).toEqual(before);
  });

  it("an inline chart keeps its frame", async () => {
    const doc = Document.create();
    Document.addContent(
      doc,
      Build.chart({
        type: "column",
        series: [{ name: "S", categories: ["a", "b"], values: [1, 2] }],
        width: 4572000,
        height: 2743200
      })
    );
    const files = await entries(await Io.toBuffer(doc));
    const xml = decoder.decode(files.get("word/document.xml")!);
    const graphic = /<a:graphic\b.*?<\/a:graphic>/s.exec(xml)![0];
    const frame =
      `<wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="77778888">` +
      `<wp:extent cx="4572000" cy="2743200"/><wp:effectExtent l="0" t="0" r="19050" b="0"/>` +
      `<wp:docPr id="4" name="Chart 4" title="Sales"/><wp:cNvGraphicFramePr/>` +
      graphic +
      `</wp:inline>`;
    files.set(
      "word/document.xml",
      encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
          para(frame) +
          `<w:sectPr/></w:body></w:document>`
      )
    );
    const source = await pack(files);
    const before = await canonicalFrames(source);
    const read = await Io.read(source);
    expect(read.body.some(b => b.type === "chart")).toBe(true);
    expect(await canonicalFrames(await Io.toBuffer(read))).toEqual(before);
  });

  it("an inline picture in a footnote keeps its frame", async () => {
    const source = await fixture(() => "<w:p/>");
    const files = await entries(source);
    const notes = decoder.decode(files.get("word/footnotes.xml")!);
    const image = [...targets(files, "word/document.xml")].find(([, t]) => t.startsWith("media/"))!;
    // The footnote part has relationships of its own; give it the picture.
    files.set(
      "word/_rels/footnotes.xml.rels",
      encoder.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${image[1]}"/></Relationships>`
      )
    );
    files.set(
      "word/footnotes.xml",
      encoder.encode(
        notes
          .replace(/<w:footnotes\b[^>]*>/, `<w:footnotes ${NS}>`)
          .replace(/(<w:footnote w:id="1"[^>]*>)/, `$1${para(inlinePicture("rId1"))}`)
      )
    );
    const patched = await pack(files);
    const before = await canonicalFrames(patched, "word/footnotes.xml");
    expect(before).toHaveLength(1);
    const written = await Io.toBuffer(await Io.read(patched));
    expect(await canonicalFrames(written, "word/footnotes.xml")).toEqual(before);
    // The frame's prefixes are declared where it is written.
    const root = /<w:footnotes\b[^>]*>/.exec(
      decoder.decode((await entries(written)).get("word/footnotes.xml")!)
    )![0];
    for (const prefix of ["wp", "wp14", "wps", "mc"]) {
      expect(root).toContain(`xmlns:${prefix}=`);
    }
  });

  it("an edited field replaces the frame markup it was read from", async () => {
    const source = await fixture((image, link) => para(tightPicture(image, link)));
    const read = await Io.read(source);
    const index = read.body.findIndex(b => b.type === "floatingImage");
    const image = read.body[index] as FloatingImage;
    const edited: FloatingImage = {
      ...image,
      wrap: { style: "square", side: "left" },
      horizontalPosition: { relativeTo: "page", offset: 914400 },
      behindDoc: true,
      altText: "New text"
    };
    const doc: DocxDocument = {
      ...read,
      body: read.body.map((b, i) => (i === index ? edited : b)) as BodyContent[]
    };
    const [frame] = await canonicalFrames(await Io.toBuffer(doc));
    expect(frame).toContain('<wp:wrapSquare wrapText="left">');
    expect(frame).not.toContain("wp:wrapPolygon");
    expect(frame).toContain(
      '<wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset>'
    );
    expect(frame).toContain('behindDoc="1"');
    expect(frame).toContain('descr="New text"');
    // What no field covers stays.
    expect(frame).toContain('title="Logo title"');
    expect(frame).toContain(`<a:hlinkClick r:id="→${HYPERLINK}" tooltip="Visit">`);
    expect(frame).toContain('wp14:anchorId="1A2B3C4D"');
    expect(frame).toContain('<wp:effectExtent b="4" l="1" r="3" t="2">');
    expect(frame).toContain("<wp14:sizeRelH");
    expect(frame).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>12700</wp:posOffset>'
    );
  });
});
