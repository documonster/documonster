import { BaseXform } from "@excel/xlsx/xform/base-xform";

interface BlipModel {
  rId: string;
  /** Alpha modulation (opacity) as OOXML percentage (e.g. 15000 = 15%). */
  alphaModFix?: number;
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
    if (model.alphaModFix !== undefined && model.alphaModFix < 100000) {
      // Render as open/close node with a:alphaModFix child
      xmlStream.openNode(this.tag, {
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "r:embed": model.rId,
        cstate: "print"
      });
      xmlStream.leafNode("a:alphaModFix", { amt: String(model.alphaModFix) });
      xmlStream.closeNode();
    } else {
      xmlStream.leafNode(this.tag, {
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "r:embed": model.rId,
        cstate: "print"
      });
    }
  }

  parseOpen(node: any): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes["r:embed"]
        };
        return true;
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
