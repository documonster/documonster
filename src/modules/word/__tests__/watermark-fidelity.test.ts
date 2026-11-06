/**
 * Text watermark geometry and size: what is set is what is written, what Word
 * writes is what is read, and both survive a second round trip unchanged.
 */

import { describe, it, expect } from "vitest";

import { Document, Io, Package } from "../index";
import type { TextWatermark } from "../types";

const decoder = new TextDecoder();

async function watermarkHeader(bytes: Uint8Array): Promise<{ path: string; xml: string }> {
  for (const path of await Package.listParts(bytes)) {
    if (!/^word\/header\d*\.xml$/.test(path)) {
      continue;
    }
    const xml = decoder.decode((await Package.readPart(bytes, path))!);
    if (xml.includes("PowerPlusWaterMarkObject")) {
      return { path, xml };
    }
  }
  throw new Error("no watermark header");
}

async function roundTrip(wm: TextWatermark): Promise<TextWatermark> {
  const doc = Document.create();
  Document.addParagraph(doc, "x");
  Document.setWatermark(doc, wm);
  const read = await Io.read(await Io.toBuffer(doc));
  return read.watermark as TextWatermark;
}

/**
 * The "DRAFT" (horizontal) watermark exactly as Word ships it in its
 * Watermarks building blocks (`Document Elements/Watermarks.dotx`): the size
 * is the shape box, the text path says `font-size:1pt` and `fitshape`
 * stretches the text over the box, there is no rotation, and the colour is a
 * VML colour name.
 */
const WORD_DRAFT_SHAPE =
  '<v:shape id="PowerPlusWaterMarkObject357870517" o:spid="_x0000_s1033" type="#_x0000_t136" ' +
  'style="position:absolute;margin-left:0;margin-top:0;width:468pt;height:280.8pt;z-index:-251642880;' +
  "mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;" +
  'mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f">' +
  '<v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/>' +
  '<w10:wrap anchorx="margin" anchory="margin"/></v:shape>';

async function wordAuthoredDocument(): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addParagraph(doc, "x");
  Document.setWatermark(doc, { type: "text", text: "DRAFT" });
  const bytes = await Io.toBuffer(doc);
  const { path, xml } = await watermarkHeader(bytes);
  const replaced = xml.replace(
    /<v:shape id="PowerPlusWaterMarkObject[\s\S]*?<\/v:shape>/,
    WORD_DRAFT_SHAPE
  );
  expect(replaced).not.toBe(xml);
  return Package.edit(bytes, [{ type: "replacePartText", path, text: replaced }]);
}

describe("text watermark fidelity", () => {
  it("sizes the shape from fontSize and reads the size back", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    Document.setWatermark(doc, { type: "text", text: "DRAFT", fontSize: 144 });
    const { xml } = await watermarkHeader(await Io.toBuffer(doc));
    // 144 half-points = 72pt: the box is one line of 72pt text tall.
    expect(xml).toMatch(/height:72pt/);
    expect(xml).toMatch(/font-size:72pt/);

    const read = await roundTrip({ type: "text", text: "DRAFT", fontSize: 144 });
    expect(read.fontSize).toBe(144);
    const bigger = await roundTrip({ type: "text", text: "DRAFT", fontSize: 288 });
    expect(bigger.widthPt!).toBeCloseTo(read.widthPt! * 2, 0);
  });

  it("reads a watermark without a size as auto-sized, not as 1pt", async () => {
    const read = await roundTrip({ type: "text", text: "DRAFT" });
    expect(read.fontSize).toBeUndefined();
  });

  it("keeps rotation 0 as 0", async () => {
    expect((await roundTrip({ type: "text", text: "DRAFT", rotation: 0 })).rotation).toBe(0);
    expect((await roundTrip({ type: "text", text: "DRAFT", rotation: -45 })).rotation).toBe(-45);
  });

  it("writes widthPt / heightPt as the shape box", async () => {
    const read = await roundTrip({ type: "text", text: "DRAFT", widthPt: 400, heightPt: 100 });
    expect(read.widthPt).toBe(400);
    expect(read.heightPt).toBe(100);
  });

  it("reads Word's own watermark as Word draws it and writes it back unchanged", async () => {
    const read = await Io.read(await wordAuthoredDocument());
    const wm = read.watermark as TextWatermark;
    expect(wm).toMatchObject({
      type: "text",
      text: "DRAFT",
      font: "Calibri",
      color: "C0C0C0",
      rotation: 0,
      widthPt: 468,
      heightPt: 280.8
    });
    expect(wm.fontSize).toBeUndefined();

    const { xml } = await watermarkHeader(await Io.toBuffer(read));
    expect(xml).toMatch(/width:468pt;height:280\.8pt/);
    expect(xml).not.toMatch(/rotation:/);
    expect(xml).toMatch(/fillcolor="#C0C0C0"/);
    expect(xml).toMatch(/font-size:1pt/);
  });
});
