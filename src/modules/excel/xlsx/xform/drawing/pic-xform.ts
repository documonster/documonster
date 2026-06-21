import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { BlipFillXform } from "@excel/xlsx/xform/drawing/blip-fill-xform";
import { NvPicPrXform } from "@excel/xlsx/xform/drawing/nv-pic-pr-xform";
import { spPrJSON } from "@excel/xlsx/xform/drawing/sp-pr";
import { StaticXform } from "@excel/xlsx/xform/static-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface PicModel {
  index?: number;
  rId?: string;
  /** Alpha modulation for transparency (OOXML percentage, e.g. 15000 = 15%). */
  alphaModFix?: number;
  /** When true, render the picture as an external linked image (`r:link`). */
  external?: boolean;
  /** Relationship id of an SVG companion (asvg:svgBlip extension). */
  svgRId?: string;
  [key: string]: any;
}

class PicXform extends BaseXform {
  declare public map: Record<string, BaseXform>;
  declare public parser?: BaseXform;
  declare public model: any;

  constructor() {
    super();

    this.map = {
      "xdr:nvPicPr": new NvPicPrXform(),
      "xdr:blipFill": new BlipFillXform(),
      "xdr:spPr": new StaticXform(spPrJSON)
    };
  }

  get tag(): string {
    return "xdr:pic";
  }

  prepare(model: PicModel, options: { index: number }): void {
    model.index = options.index + 1;
  }

  render(xmlStream: XmlSink, model: PicModel): void {
    xmlStream.openNode(this.tag);

    this.map["xdr:nvPicPr"].render(xmlStream, model);
    // Pass alphaModFix through to blipFill → blip
    this.map["xdr:blipFill"].render(xmlStream, {
      rId: model.rId,
      alphaModFix: model.alphaModFix,
      external: model.external,
      svgRId: model.svgRId
    });
    this.map["xdr:spPr"].render(xmlStream, model);

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
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

  parseText(): void {}

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.mergeModel(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
}

export { PicXform };
