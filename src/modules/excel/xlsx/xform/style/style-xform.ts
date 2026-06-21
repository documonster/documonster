import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { AlignmentXform } from "@excel/xlsx/xform/style/alignment-xform";
import { ProtectionXform } from "@excel/xlsx/xform/style/protection-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

// <xf numFmtId="[numFmtId]" fontId="[fontId]" fillId="[fillId]" borderId="[xf.borderId]" xfId="[xfId]">
//   Optional <alignment>
//   Optional <protection>
// </xf>

interface StyleModel {
  numFmtId?: number;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  xfId?: number;
  alignment?: any;
  protection?: any;
  checkbox?: boolean;
  xfComplementIndex?: number;
  pivotButton?: boolean;
  applyNumberFormat?: boolean;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
  applyProtection?: boolean;
}

interface StyleOptions {
  xfId?: boolean;
}

// Style assists translation from style model to/from xlsx
class StyleXform extends BaseXform {
  declare private xfId: boolean;
  declare public map: { alignment: AlignmentXform; protection: ProtectionXform };
  declare public parser?: BaseXform;

  constructor(options?: StyleOptions) {
    super();

    this.xfId = !!(options && options.xfId);
    this.map = {
      alignment: new AlignmentXform(),
      protection: new ProtectionXform()
    };
  }

  get tag(): string {
    return "xf";
  }

  render(xmlStream: XmlSink, model: StyleModel): void {
    xmlStream.openNode("xf", {
      numFmtId: model.numFmtId ?? 0,
      fontId: model.fontId ?? 0,
      fillId: model.fillId ?? 0,
      borderId: model.borderId ?? 0
    });
    if (this.xfId) {
      xmlStream.addAttribute("xfId", model.xfId ?? 0);
    }

    if (model.applyNumberFormat || model.numFmtId) {
      xmlStream.addAttribute("applyNumberFormat", "1");
    }
    if (model.applyFont || model.fontId) {
      xmlStream.addAttribute("applyFont", "1");
    }
    if (model.applyFill || model.fillId) {
      xmlStream.addAttribute("applyFill", "1");
    }
    if (model.applyBorder || model.borderId) {
      xmlStream.addAttribute("applyBorder", "1");
    }
    if (model.applyAlignment || model.alignment) {
      xmlStream.addAttribute("applyAlignment", "1");
    }
    if (model.applyProtection || model.protection) {
      xmlStream.addAttribute("applyProtection", "1");
    }
    if (model.pivotButton) {
      xmlStream.addAttribute("pivotButton", "1");
    }

    /**
     * Rendering tags causes close of XML stream.
     * Therefore adding attributes must be done before rendering tags.
     */

    if (model.alignment) {
      this.map.alignment.render(xmlStream, model.alignment);
    }
    if (model.protection) {
      this.map.protection.render(xmlStream, model.protection);
    }

    // Add checkbox extLst if needed
    if (model.checkbox && model.xfComplementIndex !== undefined) {
      xmlStream.openNode("extLst");
      xmlStream.openNode("ext", {
        "xmlns:xfpb": "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag",
        uri: "{C7286773-470A-42A8-94C5-96B5CB345126}"
      });
      xmlStream.leafNode("xfpb:xfComplement", { i: model.xfComplementIndex });
      xmlStream.closeNode();
      xmlStream.closeNode();
    }

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    // used during sax parsing of xml to build font object
    switch (node.name) {
      case "xf": {
        this.model = {
          numFmtId: parseInt(node.attributes.numFmtId, 10),
          fontId: parseInt(node.attributes.fontId, 10),
          fillId: parseInt(node.attributes.fillId, 10),
          borderId: parseInt(node.attributes.borderId, 10)
        };
        if (this.xfId) {
          this.model.xfId = parseInt(node.attributes.xfId, 10);
        }
        if (node.attributes.pivotButton === "1") {
          this.model.pivotButton = true;
        }
        // Preserve apply* flags from original file
        const applyFlags = [
          "applyNumberFormat",
          "applyFont",
          "applyFill",
          "applyBorder",
          "applyAlignment",
          "applyProtection"
        ] as const;
        for (const flag of applyFlags) {
          if (node.attributes[flag] === "1") {
            this.model[flag] = true;
          }
        }
        return true;
      }
      case "alignment":
        this.parser = this.map.alignment;
        this.parser.parseOpen(node);
        return true;
      case "protection":
        this.parser = this.map.protection;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        if (this.map.protection === this.parser) {
          this.model.protection = this.parser.model;
        } else {
          this.model.alignment = this.parser.model;
        }
        this.parser = undefined;
      }
      return true;
    }
    return name !== "xf";
  }
}

export { StyleXform };
