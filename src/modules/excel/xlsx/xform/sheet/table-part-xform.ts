import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

interface TablePartModel {
  rId: string;
}

class TablePartXform extends BaseXform {
  get tag(): string {
    return "tablePart";
  }

  render(xmlStream: XmlSink, model?: TablePartModel): void {
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

export { TablePartXform };
