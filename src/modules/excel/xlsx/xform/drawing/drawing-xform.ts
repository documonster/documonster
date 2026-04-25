import { colCache } from "@excel/utils/col-cache";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { AbsoluteAnchorXform } from "@excel/xlsx/xform/drawing/absolute-anchor-xform";
import { OneCellAnchorXform } from "@excel/xlsx/xform/drawing/one-cell-anchor-xform";
import { TwoCellAnchorXform } from "@excel/xlsx/xform/drawing/two-cell-anchor-xform";
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
  declare public map: { [key: string]: any };
  declare public parser: any;

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

  render(xmlStream: any, model?: DrawingModel): void {
    const renderModel = model || this.model;
    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode(this.tag, DrawingXform.DRAWING_ATTRIBUTES);

    renderModel!.anchors.forEach(item => {
      const anchor = this.map[item.anchorType];
      anchor.render(xmlStream, item);
    });

    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
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
}

export { DrawingXform };
