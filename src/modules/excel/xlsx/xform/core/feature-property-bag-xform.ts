import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

// FeaturePropertyBag is used to enable checkbox functionality
// This is a static XML file that MS Excel requires for checkboxes to work
class FeaturePropertyBagXform extends BaseXform {
  render(xmlStream: XmlSink): void {
    xmlStream.openXml({ version: "1.0", encoding: "UTF-8", standalone: "yes" });

    xmlStream.openNode("FeaturePropertyBags", {
      xmlns: "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"
    });

    // Checkbox feature
    xmlStream.leafNode("bag", { type: "Checkbox" });

    // XFControls bag
    xmlStream.openNode("bag", { type: "XFControls" });
    xmlStream.leafNode("bagId", { k: "CellControl" }, "0");
    xmlStream.closeNode();

    // XFComplement bag
    xmlStream.openNode("bag", { type: "XFComplement" });
    xmlStream.leafNode("bagId", { k: "XFControls" }, "1");
    xmlStream.closeNode();

    // XFComplements bag
    xmlStream.openNode("bag", { type: "XFComplements", extRef: "XFComplementsMapperExtRef" });
    xmlStream.openNode("a", { k: "MappedFeaturePropertyBags" });
    xmlStream.leafNode("bagId", {}, "2");
    xmlStream.closeNode();
    xmlStream.closeNode();

    xmlStream.closeNode();
  }

  parseOpen(): boolean {
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { FeaturePropertyBagXform };
