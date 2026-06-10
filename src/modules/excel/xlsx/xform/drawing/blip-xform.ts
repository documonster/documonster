import { BaseXform } from "@excel/xlsx/xform/base-xform";

interface BlipModel {
  rId: string;
  /** Alpha modulation (opacity) as OOXML percentage (e.g. 15000 = 15%). */
  alphaModFix?: number;
  /**
   * When true, the blip references an external linked image via `r:link`
   * instead of an embedded one via `r:embed`.
   */
  external?: boolean;
  /**
   * Relationship id of an SVG companion. When set, the raster blip carries an
   * `asvg:svgBlip` extension referencing the SVG media so Excel 2016+ renders
   * the vector image while older consumers fall back to the raster blip.
   */
  svgRId?: string;
}

/** OOXML extension URI for the SVG blip (Office 2016 SVG feature). */
const SVG_BLIP_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
/** Namespace for the asvg:svgBlip element. */
const SVG_BLIP_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

class BlipXform extends BaseXform<BlipModel> {
  constructor() {
    super();
    this.model = { rId: "" };
  }

  get tag(): string {
    return "a:blip";
  }

  render(xmlStream: any, model: BlipModel): void {
    // External (linked) images use `r:link`; embedded images use `r:embed`.
    const relAttr = model.external ? "r:link" : "r:embed";
    const hasAlpha = model.alphaModFix !== undefined && model.alphaModFix < 100000;
    const hasSvg = model.svgRId !== undefined;

    // A bare blip can be a leaf node; an alpha modulation or an SVG extension
    // both require child elements, so switch to the open/close form.
    if (!hasAlpha && !hasSvg) {
      xmlStream.leafNode(this.tag, {
        "xmlns:r": REL_NS,
        [relAttr]: model.rId,
        cstate: "print"
      });
      return;
    }

    xmlStream.openNode(this.tag, {
      "xmlns:r": REL_NS,
      [relAttr]: model.rId,
      cstate: "print"
    });
    if (hasAlpha) {
      xmlStream.leafNode("a:alphaModFix", { amt: String(model.alphaModFix) });
    }
    if (hasSvg) {
      xmlStream.openNode("a:extLst");
      xmlStream.openNode("a:ext", { uri: SVG_BLIP_EXT_URI });
      xmlStream.leafNode("asvg:svgBlip", {
        "xmlns:asvg": SVG_BLIP_NS,
        "r:embed": model.svgRId
      });
      xmlStream.closeNode(); // a:ext
      xmlStream.closeNode(); // a:extLst
    }
    xmlStream.closeNode(); // a:blip
  }

  parseOpen(node: any): boolean {
    switch (node.name) {
      case this.tag: {
        // A blip may carry `r:embed` (embedded) or `r:link` (external linked).
        const link = node.attributes["r:link"];
        if (link !== undefined) {
          this.model = { rId: link, external: true };
        } else {
          this.model = { rId: node.attributes["r:embed"] };
        }
        return true;
      }
      case "a:alphaModFix":
        if (node.attributes.amt) {
          this.model!.alphaModFix = parseInt(node.attributes.amt, 10);
        }
        return true;
      case "asvg:svgBlip": {
        // Capture the SVG companion's relationship id for round-trip.
        const embed = node.attributes["r:embed"];
        if (embed !== undefined && this.model) {
          this.model.svgRId = embed;
        }
        return true;
      }
      default:
        return true;
    }
  }

  parseText(): void {}

  parseClose(name: string): boolean {
    switch (name) {
      case this.tag:
        return false;
      default:
        // unprocessed internal nodes (a:extLst / a:ext / alphaModFix)
        return true;
    }
  }
}

export { BlipXform, type BlipModel };
