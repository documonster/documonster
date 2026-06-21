import { colCache } from "@excel/utils/col-cache";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { AbsoluteAnchorXform } from "@excel/xlsx/xform/drawing/absolute-anchor-xform";
import { OneCellAnchorXform } from "@excel/xlsx/xform/drawing/one-cell-anchor-xform";
import { TwoCellAnchorXform } from "@excel/xlsx/xform/drawing/two-cell-anchor-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

function getAnchorType(model: any): string {
  const range = typeof model.range === "string" ? colCache.decode(model.range) : model.range;

  if (range.pos !== undefined) {
    return "xdr:absoluteAnchor";
  }
  return range.br ? "xdr:twoCellAnchor" : "xdr:oneCellAnchor";
}

interface DrawingModel {
  anchors: any[];
}

class DrawingXform extends BaseXform<DrawingModel> {
  declare public map: Record<string, BaseXform>;
  declare public parser?: BaseXform;

  // mc:AlternateContent parse state
  private _inAlternateContent = false;
  private _acDepth = 0;
  private _inChoice = false;
  private _inFallback = false;
  private _fallbackDepth = 0;
  private _choiceRequires: string | undefined = undefined;
  /** Number of anchors before entering mc:AlternateContent — used to tag new anchors */
  private _anchorCountBeforeAC = 0;

  constructor() {
    super();

    this.map = {
      "xdr:twoCellAnchor": new TwoCellAnchorXform(),
      "xdr:oneCellAnchor": new OneCellAnchorXform(),
      "xdr:absoluteAnchor": new AbsoluteAnchorXform()
    };
    this.model = { anchors: [] };
  }

  prepare(model: DrawingModel): void {
    model.anchors.forEach((item, index) => {
      item.anchorType = getAnchorType(item);
      const anchor = this.map[item.anchorType];
      anchor.prepare(item, { index });
    });
  }

  get tag(): string {
    return "xdr:wsDr";
  }

  render(xmlStream: XmlSink, model?: DrawingModel): void {
    const renderModel = model || this.model;
    xmlStream.openXml(StdDocAttributes);
    // `<xdr:wsDr>` must declare every namespace prefix any descendant
    // uses. Inline `xmlns:mc=…` declarations on a deeply nested
    // `<mc:AlternateContent>` are legal XML and strictly sufficient for
    // the parser, but Microsoft Excel's **drawing validator** walks the
    // anchor tree ahead of the parser and rejects the whole drawing
    // when it encounters an element whose prefix was not pre-declared
    // on the wsDr root ("Removed Part: /xl/drawings/drawingN.xml
    // part. (Drawing shape)"). Excel's own output for any drawing that
    // hosts ChartEx content therefore declares all six prefixes —
    // `xdr`, `a`, `r`, `c`, `mc`, `cx` — at the root.
    //
    // Switch to the extended namespace set whenever any anchor carries
    // `alternateContent` metadata (form-control shapes, ChartEx
    // charts, and their `cx1` variants all use the same wrapper). The
    // bare `xdr` + `a` set is correct for "legacy only" drawings and
    // round-trips byte-for-byte with Excel's output for those files.
    const needsMcNamespaces = renderModel!.anchors.some(a => !!a?.alternateContent);
    const rootAttrs = needsMcNamespaces
      ? DrawingXform.DRAWING_ATTRIBUTES_WITH_MC
      : DrawingXform.DRAWING_ATTRIBUTES;
    xmlStream.openNode(this.tag, rootAttrs);

    renderModel!.anchors.forEach(item => {
      const anchor = this.map[item.anchorType];
      anchor.render(xmlStream, item);
    });

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    // Inside mc:Fallback — skip everything
    if (this._inFallback) {
      this._fallbackDepth++;
      return true;
    }

    // Delegate to active parser
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = { anchors: [] };
        break;

      case "mc:AlternateContent":
        this._acDepth = (this._acDepth ?? 0) + 1;
        if (this._acDepth === 1) {
          this._inAlternateContent = true;
          this._anchorCountBeforeAC = this.model!.anchors.length;
        }
        break;

      case "mc:Choice":
        if (this._inAlternateContent) {
          this._inChoice = true;
          this._choiceRequires = node.attributes?.Requires;
        }
        break;

      case "mc:Fallback":
        if (this._inAlternateContent) {
          this._inFallback = true;
          this._fallbackDepth = 1;
        }
        break;

      default:
        // Normal anchor dispatch (works both inside mc:Choice and at top level)
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
    // Inside mc:Fallback — skip until closed
    if (this._inFallback) {
      this._fallbackDepth--;
      if (this._fallbackDepth === 0) {
        this._inFallback = false;
      }
      return true;
    }

    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.anchors.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }

    switch (name) {
      case this.tag:
        return false;

      case "mc:AlternateContent":
        this._acDepth = (this._acDepth ?? 0) - 1;
        if (this._acDepth <= 0) {
          // Tag any anchors added during this mc:AlternateContent with alternateContent info
          if (this._inAlternateContent && this._choiceRequires) {
            const anchors = this.model!.anchors;
            for (let i = this._anchorCountBeforeAC; i < anchors.length; i++) {
              anchors[i].alternateContent = { requires: this._choiceRequires };
            }
          }
          this._inAlternateContent = false;
          this._inChoice = false;
          this._choiceRequires = undefined;
          this._acDepth = 0;
        }
        return true;

      case "mc:Choice":
        this._inChoice = false;
        return true;

      default:
        return true;
    }
  }

  reconcile(model: DrawingModel, options: any): void {
    model.anchors.forEach(anchor => {
      if (anchor.range?.pos !== undefined) {
        this.map["xdr:absoluteAnchor"].reconcile(anchor, options);
      } else if (anchor.br) {
        this.map["xdr:twoCellAnchor"].reconcile(anchor, options);
      } else {
        this.map["xdr:oneCellAnchor"].reconcile(anchor, options);
      }
    });
  }

  static DRAWING_ATTRIBUTES = {
    "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main"
  };

  /**
   * Extended namespace set used when the drawing contains anchors
   * wrapped in `<mc:AlternateContent>` — the ChartEx case. Excel's
   * strict loader requires the `mc`, `r`, `c`, and `cx` prefixes at
   * the root of `<xdr:wsDr>` before it will parse `<mc:Choice>` /
   * `<cx:chart>` descendants. Declaring them only inside the inner
   * `<mc:AlternateContent>` element (as earlier versions did) made
   * Excel reject the drawing and report "Removed Part:
   * /xl/drawings/drawingN.xml (Drawing shape)".
   *
   * Mirrors what Excel itself emits when it writes a drawing that
   * hosts a ChartEx chart — see `__tests__/data/workbook-roundtrip-
   * chartex.xlsx` (fixtures authored by Excel 2016+).
   */
  static DRAWING_ATTRIBUTES_WITH_MC = {
    "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xmlns:cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"
  };
}

export { DrawingXform };
