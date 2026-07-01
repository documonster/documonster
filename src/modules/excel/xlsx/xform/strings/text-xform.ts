import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { decodeOoxmlEscape, encodeOoxmlEscape } from "@utils/utils";
import type { ParseOpenTag, XmlSink } from "@xml/types";

//   <t xml:space="preserve"> is </t>

class TextXform extends BaseXform {
  get tag(): string {
    return "t";
  }

  render(xmlStream: XmlSink, model: string): void {
    xmlStream.openNode("t");
    if (/^\s|\n|\s$/.test(model)) {
      xmlStream.addAttribute("xml:space", "preserve");
    }
    xmlStream.writeText(encodeOoxmlEscape(model));
    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (node.name === "t") {
      this.model = "";
      return true;
    }
    return false;
  }

  parseText(text: string): void {
    // Accumulate text incrementally. In the common case (single text event),
    // model is still "" so the first assignment is just `text`.
    // For multi-chunk text, we concat directly — avoids array + join overhead.
    // model is kept up-to-date for consumers that read it before parseClose
    // (e.g. RichTextXform reads model on close without calling parseClose).
    this.model = this.model ? this.model + text : text;
  }

  parseClose(): boolean {
    // Decode OOXML escapes once, only if needed
    const raw = this.model as string;
    if (raw.includes("_x")) {
      this.model = decodeOoxmlEscape(raw);
    }
    return false;
  }
}

export { TextXform };
