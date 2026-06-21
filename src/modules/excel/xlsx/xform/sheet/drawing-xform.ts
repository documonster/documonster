import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

interface DrawingModel {
  rId: string;
}

class DrawingXform extends BaseXform {
  get tag(): string {
    return "drawing";
  }

  render(xmlStream: XmlSink, model?: DrawingModel): void {
    if (model) {
      xmlStream.leafNode(this.tag, { "r:id": model.rId });
    }
  }

  parseOpen(node: any): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes["r:id"]
        };
        return true;
      default:
        return false;
    }
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { DrawingXform };
