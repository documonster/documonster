import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { IntegerXform } from "@excel/xlsx/xform/simple/integer-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface PositionModel {
  nativeCol: number;
  nativeColOff: number;
  nativeRow: number;
  nativeRowOff: number;
}

class CellPositionXform extends BaseXform<PositionModel> {
  declare private tag: string;
  declare public map: { [key: string]: IntegerXform };
  declare public parser?: BaseXform;

  constructor(options: { tag: string }) {
    super();

    this.tag = options.tag;
    this.map = {
      "xdr:col": new IntegerXform({ tag: "xdr:col", zero: true }),
      "xdr:colOff": new IntegerXform({ tag: "xdr:colOff", zero: true }),
      "xdr:row": new IntegerXform({ tag: "xdr:row", zero: true }),
      "xdr:rowOff": new IntegerXform({ tag: "xdr:rowOff", zero: true })
    };
    this.model = { nativeCol: 0, nativeColOff: 0, nativeRow: 0, nativeRowOff: 0 };
  }

  render(xmlStream: XmlSink, model: PositionModel): void {
    xmlStream.openNode(this.tag);

    this.map["xdr:col"].render(xmlStream, model.nativeCol);
    this.map["xdr:colOff"].render(xmlStream, model.nativeColOff);

    this.map["xdr:row"].render(xmlStream, model.nativeRow);
    this.map["xdr:rowOff"].render(xmlStream, model.nativeRowOff);

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model = {
          nativeCol: this.map["xdr:col"].model,
          nativeColOff: this.map["xdr:colOff"].model,
          nativeRow: this.map["xdr:row"].model,
          nativeRowOff: this.map["xdr:rowOff"].model
        };
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
}

export { CellPositionXform };
