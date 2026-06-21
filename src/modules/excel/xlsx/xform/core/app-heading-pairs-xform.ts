import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

class AppHeadingPairsXform extends BaseXform {
  render(xmlStream: XmlSink, model: any[]): void {
    xmlStream.openNode("HeadingPairs");
    xmlStream.openNode("vt:vector", { size: 2, baseType: "variant" });

    xmlStream.openNode("vt:variant");
    xmlStream.leafNode("vt:lpstr", undefined, "Worksheets");
    xmlStream.closeNode();

    xmlStream.openNode("vt:variant");
    xmlStream.leafNode("vt:i4", undefined, model.length);
    xmlStream.closeNode();

    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    // no parsing
    return node.name === "HeadingPairs";
  }

  parseText(): void {}

  parseClose(name: string): boolean {
    return name !== "HeadingPairs";
  }
}

export { AppHeadingPairsXform };
