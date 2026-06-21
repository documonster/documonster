import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { OutlinePropertiesXform } from "@excel/xlsx/xform/sheet/outline-properties-xform";
import { PageSetupPropertiesXform } from "@excel/xlsx/xform/sheet/page-setup-properties-xform";
import { ColorXform } from "@excel/xlsx/xform/style/color-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface SheetPropertiesModel {
  tabColor?: any;
  pageSetup?: any;
  outlineProperties?: any;
}

class SheetPropertiesXform extends BaseXform {
  declare public map: { [key: string]: any };
  declare public parser?: any;

  constructor() {
    super();

    this.map = {
      tabColor: new ColorXform("tabColor"),
      pageSetUpPr: new PageSetupPropertiesXform(),
      outlinePr: new OutlinePropertiesXform()
    };
  }

  get tag(): string {
    return "sheetPr";
  }

  render(xmlStream: XmlSink, model?: SheetPropertiesModel): void {
    if (!model) {
      return;
    }
    // Check if any child xform would produce output (mirror their internal conditions)
    const hasTabColor = !!model.tabColor;
    const hasPageSetup = !!(model.pageSetup && model.pageSetup.fitToPage);
    const hasOutline =
      model.outlineProperties !== undefined &&
      model.outlineProperties !== null &&
      (model.outlineProperties.summaryBelow !== undefined ||
        model.outlineProperties.summaryRight !== undefined);

    if (hasTabColor || hasPageSetup || hasOutline) {
      xmlStream.openNode("sheetPr");
      this.map.tabColor.render(xmlStream, model.tabColor);
      this.map.pageSetUpPr.render(xmlStream, model.pageSetup);
      this.map.outlinePr.render(xmlStream, model.outlineProperties);
      xmlStream.closeNode();
    }
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (node.name === this.tag) {
      this.reset();
      return true;
    }
    if (this.map[node.name]) {
      this.parser = this.map[node.name];
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  }

  parseText(text: string): boolean {
    if (this.parser) {
      this.parser.parseText(text);
      return true;
    }
    return false;
  }

  parseClose(_name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(_name)) {
        this.parser = undefined;
      }
      return true;
    }
    if (this.map.tabColor.model || this.map.pageSetUpPr.model || this.map.outlinePr.model) {
      this.model = {};
      if (this.map.tabColor.model) {
        this.model.tabColor = this.map.tabColor.model;
      }
      if (this.map.pageSetUpPr.model) {
        this.model.pageSetup = this.map.pageSetUpPr.model;
      }
      if (this.map.outlinePr.model) {
        this.model.outlineProperties = this.map.outlinePr.model;
      }
    } else {
      this.model = null;
    }
    return false;
  }
}

export { SheetPropertiesXform };
