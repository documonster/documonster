/**
 * `Build.shape` is the one shape builder: it does everything the deprecated
 * `Build.drawingShape` did, and names its flips as the node it builds does.
 */

import { describe, it, expect } from "vitest";

import { Build, Document, Io, Package } from "../index";
import type { BodyContent } from "../types";

async function documentXml(block: BodyContent): Promise<string> {
  const doc = Document.create();
  Document.addContent(doc, block);
  const bytes = await Io.toBuffer(doc);
  return new TextDecoder().decode((await Package.readPart(bytes, "word/document.xml"))!);
}

describe("Build.shape", () => {
  it("writes what Build.drawingShape wrote for the same shape", async () => {
    const paragraphs = [Build.textParagraph("inside")];
    const legacy = Build.drawingShape({
      shapeType: "roundRect",
      width: 914400,
      height: 457200,
      fillColor: "FF0000",
      outlineColor: "0000FF",
      outlineWidth: 12700,
      textContent: paragraphs,
      altText: "alt",
      name: "S",
      behindDoc: true,
      rotation: 5400000
    });
    const current = Build.shape({
      shapeType: "roundRect",
      width: 914400,
      height: 457200,
      fill: { type: "solid", color: "FF0000" },
      outline: { color: "0000FF", width: 12700 },
      textBody: { paragraphs },
      altText: "alt",
      name: "S",
      behindDoc: true,
      rotation: 5400000,
      horizontalPosition: { relativeTo: "column", offset: 0 },
      verticalPosition: { relativeTo: "paragraph", offset: 0 },
      wrap: { style: "square" }
    });
    expect(await documentXml(current)).toBe(await documentXml(legacy));
  });

  it("takes flipHorizontal / flipVertical, as the shape node names them", () => {
    const shape = Build.shape({
      shapeType: "rect",
      width: 1,
      height: 1,
      flipHorizontal: true,
      flipVertical: true
    });
    expect(shape).toMatchObject({ flipHorizontal: true, flipVertical: true });
    // The deprecated spelling still works.
    expect(Build.shape({ shapeType: "rect", width: 1, height: 1, flipH: true })).toMatchObject({
      flipHorizontal: true
    });
  });
});
