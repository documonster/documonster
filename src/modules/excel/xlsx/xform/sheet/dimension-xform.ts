import { type RangeData, rangeRange } from "@excel/range";
import { BaseXform } from "@excel/xlsx/xform/base-xform";

class DimensionXform extends BaseXform {
  declare public model: any;

  get tag(): string {
    return "dimension";
  }

  render(xmlStream: any, model: string | RangeData | undefined): void {
    if (model) {
      const ref = typeof model === "string" ? model : rangeRange(model);
      xmlStream.leafNode("dimension", { ref });
    }
  }

  parseOpen(node: any): boolean {
    if (node.name === "dimension") {
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

export { DimensionXform };
