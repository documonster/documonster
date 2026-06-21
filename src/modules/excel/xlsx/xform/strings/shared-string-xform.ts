import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { PhoneticTextXform } from "@excel/xlsx/xform/strings/phonetic-text-xform";
import { RichTextXform } from "@excel/xlsx/xform/strings/rich-text-xform";
import { TextXform } from "@excel/xlsx/xform/strings/text-xform";
import type { XmlSink } from "@xml/types";

// <si>
//   <r></r><r></r>...
// </si>
// <si>
//   <t></t>
// </si>

type SharedStringModel = string | { richText: any[] };

class SharedStringXform extends BaseXform {
  declare public map: { [key: string]: any };
  declare public parser: any;

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

  parseOpen(node: any): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (node.name === this.tag) {
      this.model = {};
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
            let rt = (this.model as any).richText;
            if (!rt) {
              rt = (this.model as any).richText = [];
            }
            rt.push(this.parser.model);
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
