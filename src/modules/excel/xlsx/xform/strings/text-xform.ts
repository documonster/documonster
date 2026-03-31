import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { decodeOoxmlEscape, encodeOoxmlEscape } from "@utils/utils";

//   <t xml:space="preserve"> is </t>

class TextXform extends BaseXform {
  declare private _text: string[];

  get tag(): string {
    return "t";
  }

  render(xmlStream: any, model: string): void {
    xmlStream.openNode("t");
    if (/^\s|\n|\s$/.test(model)) {
      xmlStream.addAttribute("xml:space", "preserve");
    }
    xmlStream.writeText(encodeOoxmlEscape(model));
    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    if (node.name === "t") {
      this._text = [];
      this.model = "";
      return true;
    }
    return false;
  }

  parseText(text: string): void {
    this._text.push(text);
    // Keep model up-to-date with raw (undecoded) text for consumers that read
    // model before parseClose is called (e.g. RichTextXform).
    // OOXML _xHHHH_ decoding happens in parseClose for efficiency.
    this.model = this._text.length === 1 ? text : this._text.join("");
  }

  parseClose(): boolean {
    // Decode OOXML escapes once, only if needed
    const raw = this._text.length === 1 ? this._text[0] : this._text.join("");
    this.model = raw.includes("_x") ? decodeOoxmlEscape(raw) : raw;
    return false;
  }
}

export { TextXform };
