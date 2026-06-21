import { BaseCellAnchorXform } from "@excel/xlsx/xform/drawing/base-cell-anchor-xform";
import { CellPositionXform } from "@excel/xlsx/xform/drawing/cell-position-xform";
import type { PositionModel } from "@excel/xlsx/xform/drawing/cell-position-xform";
import { GraphicFrameXform } from "@excel/xlsx/xform/drawing/graphic-frame-xform";
import type { GraphicFrameModel } from "@excel/xlsx/xform/drawing/graphic-frame-xform";
import { PicXform } from "@excel/xlsx/xform/drawing/pic-xform";
import type { PicModel } from "@excel/xlsx/xform/drawing/pic-xform";
import { ShapeXform } from "@excel/xlsx/xform/drawing/shape-xform";
import type { ShapeRenderModel } from "@excel/xlsx/xform/drawing/shape-xform";
import { SpXform } from "@excel/xlsx/xform/drawing/sp-xform";
import { StaticXform } from "@excel/xlsx/xform/static-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface TwoCellModel {
  range: {
    editAs?: string;
    tl: PositionModel;
    br: PositionModel;
  };
  picture?: PicModel;
  shape?: ShapeRenderModel;
  /** Graphic frame model (for charts and other embedded objects) */
  graphicFrame?: GraphicFrameModel;
  /** Wrap the anchor in mc:AlternateContent for modern drawing clients */
  alternateContent?: { requires: string };
  medium?: unknown;
}

class TwoCellAnchorXform extends BaseCellAnchorXform {
  // Parse state for the inner `<mc:AlternateContent>` block that
  // chartEx drawings place in the shape slot of a `<xdr:twoCellAnchor>`.
  // See `render()` for the canonical layout. When `_inFallback`
  // flips on we swallow every open/close/text until the matching
  // `</mc:Fallback>` — the Fallback contents are a legacy-Excel
  // placeholder shape that structured consumers don't need to
  // surface; `alternateContent.requires` on the model is enough
  // to re-emit it verbatim on save.
  private _acDepth = 0;
  private _inFallback = false;
  private _fallbackDepth = 0;
  private _choiceRequires: string | undefined = undefined;

  constructor() {
    super();

    this.map = {
      "xdr:from": new CellPositionXform({ tag: "xdr:from" }),
      "xdr:to": new CellPositionXform({ tag: "xdr:to" }),
      "xdr:pic": new PicXform(),
      "xdr:sp": new SpXform(),
      "xdr:userShape": new ShapeXform(),
      "xdr:graphicFrame": new GraphicFrameXform(),
      "xdr:clientData": new StaticXform({ tag: "xdr:clientData" })
    };
  }

  get tag(): string {
    return "xdr:twoCellAnchor";
  }

  prepare(model: TwoCellModel, options: { index: number }): void {
    if (model.picture) {
      this.map["xdr:pic"].prepare(model.picture, options);
    } else if (model.graphicFrame) {
      this.map["xdr:graphicFrame"].prepare(model.graphicFrame, options);
    }
  }

  render(xmlStream: XmlSink, model: TwoCellModel): void {
    // ChartEx anchors (sunburst, treemap, funnel, waterfall, boxplot,
    // histogram, region map, …) MUST be wrapped in
    // `<mc:AlternateContent>`. ChartEx is a Microsoft extension that
    // was never in the base OOXML spec; Excel's strict loader
    // rejects a bare `<xdr:graphicFrame><cx:chart/>` anchor with
    // "Removed Part: /xl/drawings/drawingN.xml (Drawing shape)".
    //
    // CRITICAL: Excel emits `<mc:AlternateContent>` as the SHAPE slot
    // INSIDE the `<xdr:twoCellAnchor>`, between `<xdr:to>` and
    // `<xdr:clientData>` — NOT around the whole anchor. The Choice
    // carries the `<xdr:graphicFrame>` and the Fallback carries the
    // `<xdr:sp>`; they share the single `<xdr:from>`/`<xdr:to>` of
    // the outer anchor. This is how Microsoft Excel itself writes
    // chartEx drawings — see `tmp/ttttt.xlsx` (Excel 2021-authored)
    // for the canonical byte layout. Earlier revisions of this file
    // wrapped the ENTIRE anchor in `<mc:AlternateContent>` and
    // duplicated `<xdr:from>`/`<xdr:to>` in both Choice and
    // Fallback. Excel's strict loader rejected that nesting even
    // though MC substitution rules permit it in theory — the
    // "Removed Part: drawing1.xml" repair dialog users saw on open
    // was caused by this exact mis-nesting.
    //
    // Canonical structure:
    //
    //   <xdr:twoCellAnchor>
    //     <xdr:from>…</xdr:from>
    //     <xdr:to>…</xdr:to>
    //     <mc:AlternateContent>
    //       <mc:Choice Requires="cx1" xmlns:cx1="…/2015/9/8/chartex">
    //         <xdr:graphicFrame>…<cx:chart/>…</xdr:graphicFrame>
    //       </mc:Choice>
    //       <mc:Fallback>
    //         <xdr:sp>…placeholder…</xdr:sp>
    //       </mc:Fallback>
    //     </mc:AlternateContent>
    //     <xdr:clientData/>
    //   </xdr:twoCellAnchor>
    //
    // The Fallback MUST NOT be empty — an empty fallback is spec-legal
    // but Excel's validator treats an empty-choice+empty-fallback
    // pair as "collapses to nothing" and drops the drawing. The
    // placeholder shape we emit mirrors the "This chart isn't
    // available in your version of Excel" message Office uses.
    //
    // The `a14`-requires branch (form-control shapes) uses the LEGACY
    // wrap-the-whole-anchor form with an empty Fallback. That form
    // was fine for 2010-era form controls and round-tripping code
    // still depends on it; it is NOT the chartEx shape.
    const requires = model.alternateContent?.requires;
    const isChartEx = requires === "cx1" || requires === "cx";
    const wrapAnchorInAc = !!requires && !isChartEx;

    if (wrapAnchorInAc) {
      // Legacy wrap-the-whole-anchor path (form controls, etc.)
      xmlStream.openNode("mc:AlternateContent", {
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"
      });
      xmlStream.openNode("mc:Choice", {
        Requires: requires,
        ...(requires === "a14"
          ? { "xmlns:a14": "http://schemas.microsoft.com/office/drawing/2010/main" }
          : {})
      });
    }

    const editAs = model.range.editAs ?? (model.graphicFrame ? undefined : "oneCell");
    // `editAs="twoCell"` is the default per the CT_TwoCellAnchor
    // schema — Microsoft Excel omits the attribute in this case.
    // Emitting it anyway is spec-valid but byte-diverges from
    // Excel's own output for chartEx drawings, so suppress it when
    // the effective value is the default.
    const emitEditAs = editAs && editAs !== "twoCell" ? { editAs } : {};
    xmlStream.openNode(this.tag, emitEditAs);

    this.map["xdr:from"].render(xmlStream, model.range.tl);
    this.map["xdr:to"].render(xmlStream, model.range.br);

    if (isChartEx) {
      // ChartEx: wrap the graphicFrame/shape INSIDE the anchor in
      // mc:AlternateContent, matching Excel's emitted layout.
      this.renderChartExAlternateContent(xmlStream, model);
    } else if (model.picture) {
      this.map["xdr:pic"].render(xmlStream, model.picture);
    } else if (model.graphicFrame) {
      this.map["xdr:graphicFrame"].render(xmlStream, model.graphicFrame);
    } else if (model.shape) {
      // A user-visible shape routes to the dedicated ShapeXform; the legacy
      // form-control shape (no `kind`) stays on the SpXform path.
      if (model.shape.kind === "userShape") {
        this.map["xdr:userShape"].render(xmlStream, model.shape);
      } else {
        this.map["xdr:sp"].render(xmlStream, model.shape);
      }
    }
    this.map["xdr:clientData"].render(xmlStream, {});

    xmlStream.closeNode(); // xdr:twoCellAnchor

    if (wrapAnchorInAc) {
      xmlStream.closeNode(); // mc:Choice
      xmlStream.leafNode("mc:Fallback");
      xmlStream.closeNode(); // mc:AlternateContent
    }
  }

  /**
   * Emit the ChartEx `<mc:AlternateContent>` block that occupies the
   * shape slot inside a `<xdr:twoCellAnchor>` — between `<xdr:to>`
   * and `<xdr:clientData>`.
   *
   * Structure mirrors Microsoft Excel's own output:
   *
   *   <mc:AlternateContent>
   *     <mc:Choice Requires="cx1" xmlns:cx1="…/2015/9/8/chartex">
   *       <xdr:graphicFrame>…<cx:chart/>…</xdr:graphicFrame>
   *     </mc:Choice>
   *     <mc:Fallback>
   *       <xdr:sp>…placeholder shape…</xdr:sp>
   *     </mc:Fallback>
   *   </mc:AlternateContent>
   *
   * Both the Choice graphicFrame and the Fallback shape inherit
   * sizing from the outer `<xdr:from>`/`<xdr:to>` cell range — they
   * do NOT each re-declare the anchor; there is only one anchor.
   */
  private renderChartExAlternateContent(xmlStream: XmlSink, model: TwoCellModel): void {
    xmlStream.openNode("mc:AlternateContent", {
      "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"
    });

    xmlStream.openNode("mc:Choice", {
      "xmlns:cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
      Requires: "cx1"
    });
    if (model.graphicFrame) {
      this.map["xdr:graphicFrame"].render(xmlStream, model.graphicFrame);
    }
    xmlStream.closeNode(); // mc:Choice

    xmlStream.openNode("mc:Fallback");
    this.renderChartExFallbackShape(xmlStream, model);
    xmlStream.closeNode(); // mc:Fallback

    xmlStream.closeNode(); // mc:AlternateContent
  }

  /**
   * Emit the placeholder `<xdr:sp>` that lives inside the
   * `<mc:Fallback>` of a ChartEx anchor. Legacy Excel (2010/2013)
   * and non-Microsoft loaders that don't understand the `cx1`
   * namespace render this in place of the `cx:chart`.
   *
   * Matches Microsoft Excel's convention: cNvPr id=0 / empty name
   * (a deliberate placeholder, not a real drawing object), a white
   * rectangle sized to the anchor's cell range, and the localizable
   * "This chart isn't available" message. The shape's own xfrm
   * carries concrete EMU values rather than zeros because the
   * Fallback expansion is meant to be rendered standalone — a
   * zero-size shape would collapse and Excel 2016+ flags that as
   * "drawing shape" validation failure even though the outer
   * twoCellAnchor provides sizing.
   */
  private renderChartExFallbackShape(xmlStream: XmlSink, model: TwoCellModel): void {
    xmlStream.openNode("xdr:sp", { macro: "", textlink: "" });

    xmlStream.openNode("xdr:nvSpPr");
    // `id=0`, empty name — Microsoft's convention for Fallback
    // placeholders. Using a real numeric id here would conflict with
    // the real graphicFrame id in Choice on loaders that scan both
    // branches during validation.
    xmlStream.leafNode("xdr:cNvPr", { id: 0, name: "" });
    xmlStream.openNode("xdr:cNvSpPr");
    xmlStream.leafNode("a:spLocks", { noTextEdit: 1 });
    xmlStream.closeNode(); // xdr:cNvSpPr
    xmlStream.closeNode(); // xdr:nvSpPr

    xmlStream.openNode("xdr:spPr");
    xmlStream.openNode("a:xfrm");
    // Non-zero placeholder geometry. Excel expects the fallback
    // shape to have a concrete size even though the outer
    // twoCellAnchor already pins its bounds to from/to. Using
    // 6"×4" (5486400 × 3657600 EMU) here — the dimensions Excel
    // uses in its own fallback shapes. The position (3917950,
    // 698500) roughly tracks Excel's offset for the same fallback
    // shape (the exact values don't matter for validation; they
    // only affect legacy-client rendering).
    xmlStream.leafNode("a:off", { x: 3917950, y: 698500 });
    xmlStream.leafNode("a:ext", { cx: 5486400, cy: 3657600 });
    xmlStream.closeNode(); // a:xfrm
    xmlStream.openNode("a:prstGeom", { prst: "rect" });
    xmlStream.leafNode("a:avLst");
    xmlStream.closeNode(); // a:prstGeom
    xmlStream.openNode("a:solidFill");
    xmlStream.leafNode("a:prstClr", { val: "white" });
    xmlStream.closeNode(); // a:solidFill
    xmlStream.openNode("a:ln", { w: 1 });
    xmlStream.openNode("a:solidFill");
    xmlStream.leafNode("a:prstClr", { val: "black" });
    xmlStream.closeNode(); // a:solidFill
    xmlStream.closeNode(); // a:ln
    xmlStream.closeNode(); // xdr:spPr

    xmlStream.openNode("xdr:txBody");
    xmlStream.leafNode("a:bodyPr", { vertOverflow: "clip", horzOverflow: "clip" });
    xmlStream.leafNode("a:lstStyle");
    xmlStream.openNode("a:p");
    xmlStream.openNode("a:r");
    xmlStream.leafNode("a:rPr", { lang: "en-US", sz: 1100 });
    xmlStream.openNode("a:t");
    xmlStream.writeText(
      "This chart isn't available in your version of Excel.\n\n" +
        "Editing this shape or saving this workbook into a different file format will permanently break the chart."
    );
    xmlStream.closeNode(); // a:t
    xmlStream.closeNode(); // a:r
    xmlStream.closeNode(); // a:p
    xmlStream.closeNode(); // xdr:txBody

    xmlStream.closeNode(); // xdr:sp
  }

  parseOpen(node: ParseOpenTag): boolean {
    // Swallow everything inside `<mc:Fallback>` — it is a legacy
    // placeholder shape the writer regenerates verbatim.
    if (this._inFallback) {
      this._fallbackDepth++;
      return true;
    }

    // Intercept the MC substitution elements. The children of
    // `<mc:Choice>` (a `<xdr:graphicFrame>`, typically) should be
    // parsed by the existing xforms as if the wrapper weren't there.
    switch (node.name) {
      case "mc:AlternateContent":
        this._acDepth++;
        return true;
      case "mc:Choice":
        if (this._acDepth > 0) {
          this._choiceRequires = node.attributes?.Requires;
        }
        return true;
      case "mc:Fallback":
        if (this._acDepth > 0) {
          this._inFallback = true;
          this._fallbackDepth = 1;
        }
        return true;
      default:
        break;
    }

    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          range: {
            editAs: node.attributes.editAs
          }
        };
        this._acDepth = 0;
        this._inFallback = false;
        this._fallbackDepth = 0;
        this._choiceRequires = undefined;
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this._inFallback) {
      return;
    }
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    // Close matching against the inner `<mc:Fallback>` skip window
    // first. `_fallbackDepth` tracks nested opens inside the
    // fallback shape (xdr:sp → xdr:spPr → …); only when it drops
    // back to zero do we exit skip mode.
    if (this._inFallback) {
      this._fallbackDepth--;
      if (this._fallbackDepth === 0) {
        this._inFallback = false;
      }
      return true;
    }

    switch (name) {
      case "mc:AlternateContent":
        this._acDepth--;
        if (this._acDepth <= 0) {
          this._acDepth = 0;
          // Tag the anchor model with the Requires value so the
          // writer re-emits the MC wrapper on save.
          if (this._choiceRequires) {
            this.model.alternateContent = { requires: this._choiceRequires };
          }
          this._choiceRequires = undefined;
        }
        return true;
      case "mc:Choice":
      case "mc:Fallback":
        return true;
      default:
        break;
    }

    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model.range.tl = this.map["xdr:from"].model;
        this.model.range.br = this.map["xdr:to"].model;
        this.model.picture = this.map["xdr:pic"].model;
        this.model.graphicFrame = this.map["xdr:graphicFrame"].model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(
    model: TwoCellModel,
    options: Parameters<BaseCellAnchorXform["reconcilePicture"]>[1]
  ): void {
    if (model.picture) {
      model.medium = this.reconcilePicture(model.picture, options);
    }
    // graphicFrame reconciliation handled at DrawingXform level
  }
}

export { TwoCellAnchorXform };
