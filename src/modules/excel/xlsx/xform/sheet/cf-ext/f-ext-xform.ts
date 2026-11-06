import { fromFileFormula, toFileFormula } from "@excel/utils/future-functions";
import { BaseXform } from "@excel/xlsx/xform/base-xform";

class FExtXform extends BaseXform {
  get tag() {
    return "xm:f";
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag, null, typeof model === "string" ? toFileFormula(model) : model);
  }

  parseOpen() {
    this.model = "";
  }

  parseText(text) {
    this.model += text;
  }

  parseClose(name) {
    if (name === this.tag && typeof this.model === "string") {
      this.model = fromFileFormula(this.model);
    }
    return name !== this.tag;
  }
}

export { FExtXform };
