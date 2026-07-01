import { BaseXform } from "@excel/xlsx/xform/base-xform";
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
          left: node.attributes.left ? parseFloat(node.attributes.left) : 0.7,
          right: node.attributes.right ? parseFloat(node.attributes.right) : 0.7,
          top: node.attributes.top ? parseFloat(node.attributes.top) : 0.75,
          bottom: node.attributes.bottom ? parseFloat(node.attributes.bottom) : 0.75,
          header: node.attributes.header ? parseFloat(node.attributes.header) : 0.3,
          footer: node.attributes.footer ? parseFloat(node.attributes.footer) : 0.3
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
