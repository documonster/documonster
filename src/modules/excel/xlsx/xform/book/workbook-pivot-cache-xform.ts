import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface PivotCacheModel {
  cacheId: string;
  rId: string;
}

class WorkbookPivotCacheXform extends BaseXform {
  render(xmlStream: XmlSink, model: PivotCacheModel): void {
    xmlStream.leafNode("pivotCache", {
      cacheId: model.cacheId,
      "r:id": model.rId
    });
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === "pivotCache") {
      this.model = {
        cacheId: node.attributes.cacheId,
        rId: node.attributes["r:id"]
      };
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { WorkbookPivotCacheXform };
