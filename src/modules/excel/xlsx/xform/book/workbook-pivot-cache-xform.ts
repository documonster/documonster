import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

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

  parseOpen(node: any): boolean {
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
