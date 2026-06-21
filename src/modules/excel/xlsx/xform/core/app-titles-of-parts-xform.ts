import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

class AppTitlesOfPartsXform extends BaseXform {
  render(xmlStream: XmlSink, model: any[]): void {
    xmlStream.openNode("TitlesOfParts");
    xmlStream.openNode("vt:vector", { size: model.length, baseType: "lpstr" });

    model.forEach(sheet => {
      xmlStream.leafNode("vt:lpstr", undefined, sheet.name);
    });

    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    // no parsing
    return node.name === "TitlesOfParts";
  }

  parseText(): void {}

  parseClose(name: string): boolean {
    return name !== "TitlesOfParts";
  }
}

export { AppTitlesOfPartsXform };
