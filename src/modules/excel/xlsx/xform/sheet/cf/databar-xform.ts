import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { CfvoXform } from "@excel/xlsx/xform/sheet/cf/cfvo-xform";
import { ColorXform } from "@excel/xlsx/xform/style/color-xform";
import type { XmlSink } from "@xml/types";

class DatabarXform extends CompositeXform {
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
    return "dataBar";
  }

  render(xmlStream: XmlSink, model: any): void {
    xmlStream.openNode(this.tag);

    model.cfvo.forEach((cfvo: any) => {
      this.cfvoXform.render(xmlStream, cfvo);
    });
    this.colorXform.render(xmlStream, model.color);

    xmlStream.closeNode();
  }

  createNewModel(): any {
    return {
      cfvo: []
    };
  }

  onParserClose(name: string, parser: any): void {
    switch (name) {
      case "cfvo":
        this.model.cfvo.push(parser.model);
        break;
      case "color":
        this.model.color = parser.model;
        break;
    }
  }
}

export { DatabarXform };
