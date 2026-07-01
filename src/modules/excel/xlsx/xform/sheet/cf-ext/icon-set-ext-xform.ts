import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { CfIconExtXform } from "@excel/xlsx/xform/sheet/cf-ext/cf-icon-ext-xform";
import { CfvoExtXform } from "@excel/xlsx/xform/sheet/cf-ext/cfvo-ext-xform";
import type { ParseOpenTag } from "@xml/types";

class IconSetExtXform extends CompositeXform {
  cfvoXform: CfvoExtXform;
  cfIconXform: CfIconExtXform;

  constructor() {
    super();

    this.map = {
      "x14:cfvo": (this.cfvoXform = new CfvoExtXform()),
      "x14:cfIcon": (this.cfIconXform = new CfIconExtXform())
    };
  }

  get tag() {
    return "x14:iconSet";
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      iconSet: BaseXform.toStringAttribute(model.iconSet),
      reverse: BaseXform.toBoolAttribute(model.reverse, false),
      showValue: BaseXform.toBoolAttribute(model.showValue, true),
      custom: BaseXform.toBoolAttribute(model.icons, false)
    });

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });

    if (model.icons) {
      model.icons.forEach((icon, i) => {
        icon.iconId = i;
        this.cfIconXform.render(xmlStream, icon);
      });
    }

    xmlStream.closeNode();
  }

  createNewModel({ attributes }: ParseOpenTag) {
    return {
      cfvo: [],
      iconSet: BaseXform.toStringValue(attributes.iconSet, "3TrafficLights"),
      reverse: BaseXform.toBoolValue(attributes.reverse, false),
      showValue: BaseXform.toBoolValue(attributes.showValue, true)
    };
  }

  onParserClose(name, parser) {
    const [, prop] = name.split(":");
    switch (prop) {
      case "cfvo":
        this.model.cfvo.push(parser.model);
        break;

      case "cfIcon":
        if (!this.model.icons) {
          this.model.icons = [];
        }
        this.model.icons.push(parser.model);
        break;

      default:
        this.model[prop] = parser.model;
        break;
    }
  }
}

export { IconSetExtXform };
