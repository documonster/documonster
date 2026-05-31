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
}

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
    if (model.alphaModFix !== undefined && model.alphaModFix < 100000) {
      // Render as open/close node with a:alphaModFix child
      xmlStream.openNode(this.tag, {
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        [relAttr]: model.rId,
        cstate: "print"
      });
      xmlStream.leafNode("a:alphaModFix", { amt: String(model.alphaModFix) });
      xmlStream.closeNode();
    } else {
      xmlStream.leafNode(this.tag, {
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        [relAttr]: model.rId,
        cstate: "print"
      });
    }
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
        // unprocessed internal nodes
        return true;
    }
  }
}

export { BlipXform, type BlipModel };
