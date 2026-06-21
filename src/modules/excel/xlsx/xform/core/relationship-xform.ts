import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlAttributes, XmlSink } from "@xml/types";

interface RelationshipModel {
  Id?: string;
  Type?: string;
  Target?: string;
  TargetMode?: string;
}

class RelationshipXform extends BaseXform {
  render(xmlStream: XmlSink, model: RelationshipModel): void {
    // RelationshipModel is a plain attribute bag (all string|undefined); it
    // matches XmlAttributes structurally but lacks an index signature.
    xmlStream.leafNode("Relationship", model as XmlAttributes);
  }

  parseOpen(node: ParseOpenTag): boolean {
    switch (node.name) {
      case "Relationship":
        this.model = node.attributes;
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

export { RelationshipXform };
