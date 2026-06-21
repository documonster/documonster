import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

/**
 * Xform for individual page break (brk element)
 * Used by both RowBreaksXform and ColBreaksXform
 */
class PageBreaksXform extends BaseXform {
  declare public model: any;

  get tag(): string {
    return "brk";
  }

  render(xmlStream: XmlSink, model: any): void {
    xmlStream.leafNode("brk", model);
  }

  parseOpen(node: any): boolean {
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
