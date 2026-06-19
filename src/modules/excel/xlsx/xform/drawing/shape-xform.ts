import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { EMU_PER_POINT } from "@utils/units";

/** Fill specification for a drawing shape. */
export interface ShapeFill {
  /** Solid fill colour as a hex RGB string (e.g. "FF0000"). Omit for no fill. */
  color?: string;
}

/** Line (outline) specification for a drawing shape. */
export interface ShapeLine {
  /** Line colour as a hex RGB string (e.g. "000000"). */
  color?: string;
  /** Line width in points. */
  width?: number;
}

/**
 * Model for a user-visible drawing shape (`<xdr:sp>`).
 *
 * Distinct from the form-control shape rendered by `SpXform` — this one is
 * visible, carries a configurable preset geometry, solid fill, outline and an
 * optional text label, and is NOT wrapped in an `a14` AlternateContent block.
 */
export interface ShapeRenderModel {
  /** Marks this as a user shape so the anchor routes to ShapeXform. */
  kind: "userShape";
  /** Unique drawing id. */
  cNvPrId: number;
  /** Display name (e.g. "Rectangle 1"). */
  name: string;
  /** Preset geometry name (e.g. "rect", "ellipse", "line", "roundRect"). */
  shapeType: string;
  fill?: ShapeFill;
  line?: ShapeLine;
  /** Optional text label centred in the shape. */
  text?: string;
}

/**
 * Normalize a user-supplied colour to the bare 6-digit RGB hex that OOXML's
 * `<a:srgbClr val="...">` requires:
 * - strips a leading `#`
 * - accepts 8-digit ARGB (the form documonster uses for cell fills) and drops the
 *   leading alpha byte, since `srgbClr` carries no alpha channel
 * - upper-cases
 *
 * Anything that isn't a 6- or 8-digit hex string is passed through unchanged so
 * a caller using a less common form is not silently broken.
 */
function normalizeColor(color: string): string {
  const hex = color.startsWith("#") ? color.slice(1) : color;
  if (/^[0-9a-fA-F]{8}$/.test(hex)) {
    // ARGB → RGB: drop the alpha byte (srgbClr has no alpha component).
    return hex.slice(2).toUpperCase();
  }
  if (/^[0-9a-fA-F]{6}$/.test(hex)) {
    return hex.toUpperCase();
  }
  return hex;
}

/**
 * Renders a user-visible drawing shape. Geometry/position is governed by the
 * enclosing anchor (`xfrm` is written as zero, matching how Excel anchors a
 * shape to a cell range), while preset geometry, fill, outline and text are
 * taken from the model. Write-only: shapes are not parsed back on read (the
 * same limitation that already applies to all non-chart drawing content).
 */
class ShapeXform extends BaseXform {
  declare public model: ShapeRenderModel;

  get tag(): string {
    return "xdr:sp";
  }

  render(xmlStream: any, model: ShapeRenderModel): void {
    xmlStream.openNode("xdr:sp", { macro: "", textlink: "" });

    // --- Non-visual shape properties ---
    xmlStream.openNode("xdr:nvSpPr");
    xmlStream.leafNode("xdr:cNvPr", { id: model.cNvPrId, name: model.name });
    xmlStream.leafNode("xdr:cNvSpPr", {});
    xmlStream.closeNode(); // xdr:nvSpPr

    // --- Shape properties ---
    xmlStream.openNode("xdr:spPr");

    // Position/size is driven by the anchor; emit a zero xfrm placeholder.
    xmlStream.openNode("a:xfrm");
    xmlStream.leafNode("a:off", { x: 0, y: 0 });
    xmlStream.leafNode("a:ext", { cx: 0, cy: 0 });
    xmlStream.closeNode(); // a:xfrm

    xmlStream.openNode("a:prstGeom", { prst: model.shapeType });
    xmlStream.leafNode("a:avLst");
    xmlStream.closeNode(); // a:prstGeom

    // Fill: a colour produces a solidFill, otherwise an explicit noFill.
    if (model.fill && model.fill.color) {
      xmlStream.openNode("a:solidFill");
      xmlStream.leafNode("a:srgbClr", { val: normalizeColor(model.fill.color) });
      xmlStream.closeNode(); // a:solidFill
    } else {
      xmlStream.leafNode("a:noFill");
    }

    // Line: an `a:ln` with width (pt → EMU) and/or a solid colour. When width
    // is given without a colour, Excel applies its default outline colour at
    // that width. When neither colour nor width is supplied, emit an explicit
    // noFill line (no visible outline).
    if (model.line && (model.line.color || model.line.width !== undefined)) {
      const lnAttrs: Record<string, number> = {};
      if (model.line.width !== undefined) {
        lnAttrs.w = Math.round(model.line.width * EMU_PER_POINT);
      }
      xmlStream.openNode("a:ln", lnAttrs);
      if (model.line.color) {
        xmlStream.openNode("a:solidFill");
        xmlStream.leafNode("a:srgbClr", { val: normalizeColor(model.line.color) });
        xmlStream.closeNode(); // a:solidFill
      }
      xmlStream.closeNode(); // a:ln
    } else {
      xmlStream.openNode("a:ln");
      xmlStream.leafNode("a:noFill");
      xmlStream.closeNode(); // a:ln
    }

    xmlStream.closeNode(); // xdr:spPr

    // --- Text body ---
    xmlStream.openNode("xdr:txBody");
    xmlStream.leafNode("a:bodyPr", { vertOverflow: "clip", wrap: "square", anchor: "ctr" });
    xmlStream.leafNode("a:lstStyle");
    xmlStream.openNode("a:p");
    xmlStream.openNode("a:pPr", { algn: "ctr" });
    xmlStream.closeNode(); // a:pPr
    if (model.text) {
      xmlStream.openNode("a:r");
      xmlStream.openNode("a:rPr", { lang: "en-US" });
      xmlStream.closeNode(); // a:rPr
      xmlStream.openNode("a:t");
      xmlStream.writeText(model.text);
      xmlStream.closeNode(); // a:t
      xmlStream.closeNode(); // a:r
    } else {
      xmlStream.leafNode("a:endParaRPr", { lang: "en-US" });
    }
    xmlStream.closeNode(); // a:p
    xmlStream.closeNode(); // xdr:txBody

    xmlStream.closeNode(); // xdr:sp
  }
}

export { ShapeXform };
