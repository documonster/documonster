import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

function isDefined(attr: any): boolean {
  return typeof attr !== "undefined";
}

interface OutlinePropertiesModel {
  summaryBelow?: boolean;
  summaryRight?: boolean;
}

class OutlinePropertiesXform extends BaseXform {
  get tag(): string {
    return "outlinePr";
  }

  render(xmlStream: XmlSink, model?: OutlinePropertiesModel): boolean {
    if (model && (isDefined(model.summaryBelow) || isDefined(model.summaryRight))) {
      xmlStream.leafNode(this.tag, {
        summaryBelow: isDefined(model.summaryBelow) ? Number(model.summaryBelow) : undefined,
        summaryRight: isDefined(model.summaryRight) ? Number(model.summaryRight) : undefined
      });
      return true;
    }
    return false;
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === this.tag) {
      this.model = {
        summaryBelow: isDefined(node.attributes.summaryBelow)
          ? Boolean(Number(node.attributes.summaryBelow))
          : undefined,
        summaryRight: isDefined(node.attributes.summaryRight)
          ? Boolean(Number(node.attributes.summaryRight))
          : undefined
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

export { OutlinePropertiesXform };
