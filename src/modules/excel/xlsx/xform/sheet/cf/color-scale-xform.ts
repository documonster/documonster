import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { CfvoXform } from "@excel/xlsx/xform/sheet/cf/cfvo-xform";
import { ColorXform } from "@excel/xlsx/xform/style/color-xform";
import type { XmlSink } from "@xml/types";

class ColorScaleXform extends CompositeXform {
  cfvoXform: CfvoXform;
  colorXform: ColorXform;

  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform()),
      color: (this.colorXform = new ColorXform())
    };
  }

  get tag(): string {
    return "colorScale";
  }

  render(xmlStream: XmlSink, model: any): void {
    xmlStream.openNode(this.tag);

    model.cfvo.forEach((cfvo: any) => {
      this.cfvoXform.render(xmlStream, cfvo);
    });
    model.color.forEach((color: any) => {
      this.colorXform.render(xmlStream, color);
    });

    xmlStream.closeNode();
  }

  createNewModel(node: any): any {
    return {
      cfvo: [],
      color: []
    };
  }

  onParserClose(name: string, parser: any): void {
    this.model[name].push(parser.model);
  }
}

export { ColorScaleXform };
