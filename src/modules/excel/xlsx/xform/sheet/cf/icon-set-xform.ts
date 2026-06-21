import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { CfvoXform } from "@excel/xlsx/xform/sheet/cf/cfvo-xform";
import type { XmlSink } from "@xml/types";

class IconSetXform extends CompositeXform {
  cfvoXform: CfvoXform;

  constructor() {
    super();

    this.map = {
      cfvo: (this.cfvoXform = new CfvoXform())
    };
  }

  get tag(): string {
    return "iconSet";
  }

  render(xmlStream: XmlSink, model: any): void {
    xmlStream.openNode(this.tag, {
      iconSet: BaseXform.toStringAttribute(model.iconSet, "3TrafficLights"),
      reverse: BaseXform.toBoolAttribute(model.reverse, false),
      showValue: BaseXform.toBoolAttribute(model.showValue, true)
    });

    model.cfvo.forEach((cfvo: any) => {
      this.cfvoXform.render(xmlStream, cfvo);
    });

    xmlStream.closeNode();
  }

  createNewModel({ attributes }: any): any {
    return {
      iconSet: BaseXform.toStringValue(attributes.iconSet, "3TrafficLights"),
      reverse: BaseXform.toBoolValue(attributes.reverse),
      showValue: BaseXform.toBoolValue(attributes.showValue),
      cfvo: []
    };
  }

  onParserClose(name: string, parser: any): void {
    this.model[name].push(parser.model);
  }
}

export { IconSetXform };
