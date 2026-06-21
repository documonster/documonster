import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface PageSetupPropertiesModel {
  fitToPage: boolean;
}

class PageSetupPropertiesXform extends BaseXform {
  get tag(): string {
    return "pageSetUpPr";
  }

  render(xmlStream: XmlSink, model?: PageSetupPropertiesModel): boolean {
    if (model && model.fitToPage) {
      xmlStream.leafNode(this.tag, {
        fitToPage: model.fitToPage ? "1" : undefined
      });
      return true;
    }
    return false;
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === this.tag) {
      this.model = {
        fitToPage: node.attributes.fitToPage === "1"
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

export { PageSetupPropertiesXform };
