import type { RichText } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { PhoneticTextXform } from "@excel/xlsx/xform/strings/phonetic-text-xform";
import { RichTextXform } from "@excel/xlsx/xform/strings/rich-text-xform";
import { TextXform } from "@excel/xlsx/xform/strings/text-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

// <si>
//   <r></r><r></r>...
// </si>
// <si>
//   <t></t>
// </si>

type SharedStringModel = string | { richText: RichText[] };

class SharedStringXform extends BaseXform<SharedStringModel> {
  declare public map: Record<string, BaseXform>;
  declare public parser?: BaseXform;

  constructor(model?: SharedStringModel) {
    super();

    this.model = model;

    this.map = {
      r: new RichTextXform(),
      t: new TextXform(),
      rPh: new PhoneticTextXform()
    };
  }

  get tag(): string {
    return "si";
  }

  render(xmlStream: XmlSink, model?: SharedStringModel): void {
    xmlStream.openNode(this.tag);
    if (
      model &&
      typeof model === "object" &&
      Object.prototype.hasOwnProperty.call(model, "richText") &&
      model.richText
    ) {
      if (model.richText.length) {
        model.richText.forEach(text => {
          this.map.r.render(xmlStream, text);
        });
      } else {
        this.map.t.render(xmlStream, "");
      }
    } else if (model !== undefined && model !== null) {
      this.map.t.render(xmlStream, model as string);
    }
    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (node.name === this.tag) {
      // Empty accumulator; `<r>` children attach a richText array, `<t>`
      // replaces the model with a plain string.
      this.model = {} as SharedStringModel;
      return true;
    }
    this.parser = this.map[node.name];
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        switch (name) {
          case "r": {
            const richModel = this.model as { richText: RichText[] };
            let rt = richModel.richText;
            if (!rt) {
              rt = richModel.richText = [];
            }
            rt.push(this.parser.model as RichText);
            break;
          }
          case "t":
            this.model = this.parser.model;
            break;
          default:
            break;
        }
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

export { SharedStringXform };
