import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

class AppTitlesOfPartsXform extends BaseXform {
  render(xmlStream: XmlSink, model: { name: string }[]): void {
    xmlStream.openNode("TitlesOfParts");
    xmlStream.openNode("vt:vector", { size: model.length, baseType: "lpstr" });

    model.forEach(sheet => {
      xmlStream.leafNode("vt:lpstr", undefined, sheet.name);
    });

    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    // no parsing
    return node.name === "TitlesOfParts";
  }

  parseText(): void {}

  parseClose(name: string): boolean {
    return name !== "TitlesOfParts";
  }
}

export { AppTitlesOfPartsXform };
