import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface BooleanXformOptions {
  tag: string;
  attr: string;
  /**
   * Write `false` as `<tag val="0"/>` rather than leaving the element out.
   *
   * In a cell font the two are the same thing — no `<b>` is not bold — so the element is omitted. In a
   * differential format they are not: an absent `<b>` leaves the cell's own weight alone, while `<b val="0"/>`
   * *removes* bold. Omitting it turned "un-bold these cells" into "change nothing".
   */
  writeFalse?: boolean;
}

/**
 * A `CT_BooleanProperty` — an element whose presence means `true` unless its `val` says otherwise.
 *
 * `val` is `xsd:boolean`, so `"0"` and `"false"` are both false. This used to set `true` whenever the element
 * appeared, which read `<b val="0"/>` as bold — the exact opposite of what the file says.
 */
class BooleanXform extends BaseXform {
  declare private tag: string;
  declare private attr: string;
  declare private writeFalse: boolean;

  constructor(options: BooleanXformOptions) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
    this.writeFalse = options.writeFalse === true;
  }

  render(xmlStream: XmlSink, model?: boolean): void {
    if (model) {
      xmlStream.openNode(this.tag);
      xmlStream.closeNode();
    } else if (model === false && this.writeFalse) {
      xmlStream.leafNode(this.tag, { [this.attr]: "0" });
    }
  }

  parseOpen(node: ParseOpenTag): void {
    if (node.name === this.tag) {
      const value = node.attributes[this.attr];
      this.model = value === undefined || !(value === "0" || value === "false");
    }
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { BooleanXform };
