import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { GenericEchoXform } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

/**
 * `<xdr:pic><xdr:spPr>` — a picture's shape properties, specifically the
 * absolute position/size in its `<a:xfrm>`.
 *
 * This replaces the old hard-coded `StaticXform` template that always emitted
 * `<a:off x="0" y="0"/><a:ext cx="0" cy="0"/>`. That template discarded a
 * picture's real geometry on every write. For a `twoCellAnchor editAs="oneCell"`
 * picture Excel renders from THIS absolute `<a:xfrm>` — the anchor's `<xdr:to>`
 * cell is only a cache — so a plain read/write round-trip zeroed the offset and
 * extent, rendering the picture at (0,0) with zero size (i.e. invisible).
 *
 * The values are captured verbatim on read and re-emitted on write. When they
 * are absent (a picture created programmatically that never had an `<a:xfrm>`,
 * or a shape whose geometry is fully described by the anchor) they fall back to
 * zero, exactly matching the previous static template's behaviour.
 */
export interface PicSpPrModel {
  xfrmOffX?: number;
  xfrmOffY?: number;
  xfrmExtCx?: number;
  xfrmExtCy?: number;
  /** Full loaded `<xdr:spPr>` subtree for lossless round-trip. */
  rawSpPr?: EchoNode;
}

class PicSpPrXform extends BaseXform<PicSpPrModel> {
  private readonly echo = new GenericEchoXform("xdr:spPr");

  get tag(): string {
    return "xdr:spPr";
  }

  render(xmlStream: XmlSink, model?: PicSpPrModel): void {
    if (model?.rawSpPr) {
      this.echo.render(xmlStream, model.rawSpPr);
      return;
    }
    xmlStream.openNode(this.tag);
    xmlStream.openNode("a:xfrm");
    xmlStream.leafNode("a:off", { x: model?.xfrmOffX ?? 0, y: model?.xfrmOffY ?? 0 });
    xmlStream.leafNode("a:ext", { cx: model?.xfrmExtCx ?? 0, cy: model?.xfrmExtCy ?? 0 });
    xmlStream.closeNode(); // a:xfrm
    xmlStream.openNode("a:prstGeom", { prst: "rect" });
    xmlStream.leafNode("a:avLst");
    xmlStream.closeNode(); // a:prstGeom
    xmlStream.closeNode(); // xdr:spPr
  }

  parseOpen(node: ParseOpenTag): boolean {
    switch (node.name) {
      case this.tag:
        this.reset();
        this.echo.parseOpen(node);
        return true;
      case "a:off":
        this.echo.parseOpen(node);
        this.mergeModel({
          xfrmOffX: parseInt(node.attributes.x, 10),
          xfrmOffY: parseInt(node.attributes.y, 10)
        });
        return true;
      case "a:ext":
        this.echo.parseOpen(node);
        this.mergeModel({
          xfrmExtCx: parseInt(node.attributes.cx, 10),
          xfrmExtCy: parseInt(node.attributes.cy, 10)
        });
        return true;
      default:
        this.echo.parseOpen(node);
        return true;
    }
  }

  parseText(text: string): void {
    this.echo.parseText(text);
  }

  parseClose(name: string): boolean {
    const keepParsing = this.echo.parseClose(name);
    if (!keepParsing) {
      const raw = this.echo.model;
      // Avoid bloating every ordinary picture model with a duplicate XML tree.
      // The structured fields already round-trip the canonical rect form; retain
      // raw XML only when it contains attributes/elements that form cannot
      // represent (rotation, flip, fill, line, effects, custom geometry, …).
      if (raw && !isCanonicalRectSpPr(raw)) {
        this.mergeModel({ rawSpPr: raw });
      }
    }
    return keepParsing;
  }
}

function isCanonicalRectSpPr(root: EchoNode): boolean {
  if (Object.keys(root.attrs).length !== 0 || root.children.length !== 2) {
    return false;
  }
  const [xfrm, geom] = root.children;
  return (
    xfrm.tag === "a:xfrm" &&
    Object.keys(xfrm.attrs).length === 0 &&
    xfrm.children.length === 2 &&
    xfrm.children[0].tag === "a:off" &&
    xfrm.children[1].tag === "a:ext" &&
    geom.tag === "a:prstGeom" &&
    geom.attrs.prst === "rect" &&
    geom.children.length === 1 &&
    geom.children[0].tag === "a:avLst"
  );
}

export { PicSpPrXform };
