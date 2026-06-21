import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlAttributes, XmlSink } from "@xml/types";

interface PageBreakModel {
  id: number;
  max: number;
  man: number;
  min?: number;
}

/**
 * Xform for individual page break (brk element)
 * Used by both RowBreaksXform and ColBreaksXform
 */
class PageBreaksXform extends BaseXform<PageBreakModel> {
  get tag(): string {
    return "brk";
  }

  render(xmlStream: XmlSink, model?: PageBreakModel): void {
    // PageBreakModel is a numeric attribute bag; matches XmlAttributes
    // structurally but lacks an index signature.
    xmlStream.leafNode("brk", model as XmlAttributes | undefined);
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === "brk") {
      const { id, max, man, min } = node.attributes;
      this.model = {
        id: +id,
        max: +max,
        man: +man
      };
      if (min !== undefined) {
        this.model.min = +min;
      }
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { PageBreaksXform };
