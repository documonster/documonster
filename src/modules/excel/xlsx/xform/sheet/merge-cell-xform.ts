import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

class MergeCellXform extends BaseXform {
  get tag(): string {
    return "mergeCell";
  }

  render(xmlStream: XmlSink, model: string): void {
    xmlStream.leafNode("mergeCell", { ref: model });
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === "mergeCell") {
      this.model = node.attributes.ref;
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { MergeCellXform };
