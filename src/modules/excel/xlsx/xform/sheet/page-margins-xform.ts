import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { parseXsdFloatOrText } from "@excel/xlsx/xform/xsd-values";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface PageMarginsModel {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

class PageMarginsXform extends BaseXform {
  get tag(): string {
    return "pageMargins";
  }

  render(xmlStream: XmlSink, model: PageMarginsModel): void {
    if (model) {
      const attributes = {
        left: model.left,
        right: model.right,
        top: model.top,
        bottom: model.bottom,
        header: model.header,
        footer: model.footer
      };
      if (Object.values(attributes).some((value: unknown) => value !== undefined)) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  }

  parseOpen(node: ParseOpenTag): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          left: parseXsdFloatOrText(node.attributes.left || undefined) ?? 0.7,
          right: parseXsdFloatOrText(node.attributes.right || undefined) ?? 0.7,
          top: parseXsdFloatOrText(node.attributes.top || undefined) ?? 0.75,
          bottom: parseXsdFloatOrText(node.attributes.bottom || undefined) ?? 0.75,
          header: parseXsdFloatOrText(node.attributes.header || undefined) ?? 0.3,
          footer: parseXsdFloatOrText(node.attributes.footer || undefined) ?? 0.3
        };
        return true;
      default:
        return false;
    }
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { PageMarginsXform };
