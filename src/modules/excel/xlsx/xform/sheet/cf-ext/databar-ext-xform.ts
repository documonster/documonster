import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { ColorXform } from "@excel/xlsx/xform/style/color-xform";
import { CfvoExtXform } from "@excel/xlsx/xform/sheet/cf-ext/cfvo-ext-xform";

class DatabarExtXform extends CompositeXform {
  cfvoXform: CfvoExtXform;
  borderColorXform: ColorXform;
  negativeBorderColorXform: ColorXform;
  negativeFillColorXform: ColorXform;
  axisColorXform: ColorXform;

  constructor() {
    super();

    this.map = {
      "x14:cfvo": (this.cfvoXform = new CfvoExtXform()),
      "x14:borderColor": (this.borderColorXform = new ColorXform("x14:borderColor")),
      "x14:negativeBorderColor": (this.negativeBorderColorXform = new ColorXform(
        "x14:negativeBorderColor"
      )),
      "x14:negativeFillColor": (this.negativeFillColorXform = new ColorXform(
        "x14:negativeFillColor"
      )),
      "x14:axisColor": (this.axisColorXform = new ColorXform("x14:axisColor"))
    };
  }

  static isExt(_rule) {
    // All data bars require an ext section for proper rendering in modern Excel
    // (Excel 2010+). Without the ext section, data bars may not display or Excel
    // may show a repair dialog. The ext section carries additional formatting
    // attributes (minLength, maxLength, border, gradient, axis, etc.) that the
    // primary <dataBar> element does not support.
    return true;
  }

  get tag() {
    return "x14:dataBar";
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      minLength: BaseXform.toIntAttribute(model.minLength, 0, true),
      maxLength: BaseXform.toIntAttribute(model.maxLength, 100, true),
      border: BaseXform.toBoolAttribute(model.border, false),
      gradient: BaseXform.toBoolAttribute(model.gradient, true),
      negativeBarColorSameAsPositive: BaseXform.toBoolAttribute(
        model.negativeBarColorSameAsPositive,
        true
      ),
      negativeBarBorderColorSameAsPositive: BaseXform.toBoolAttribute(
        model.negativeBarBorderColorSameAsPositive,
        true
      ),
      axisPosition: BaseXform.toAttribute(model.axisPosition, "auto"),
      direction: BaseXform.toAttribute(model.direction, "leftToRight")
    });

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });

    this.borderColorXform.render(xmlStream, model.borderColor);
    this.negativeBorderColorXform.render(xmlStream, model.negativeBorderColor);
    this.negativeFillColorXform.render(xmlStream, model.negativeFillColor);
    this.axisColorXform.render(xmlStream, model.axisColor);

    xmlStream.closeNode();
  }

  createNewModel({ attributes }) {
    return {
      cfvo: [],
      minLength: BaseXform.toIntValue(attributes.minLength, 0),
      maxLength: BaseXform.toIntValue(attributes.maxLength, 100),
      border: BaseXform.toBoolValue(attributes.border, false),
      gradient: BaseXform.toBoolValue(attributes.gradient, true),
      negativeBarColorSameAsPositive: BaseXform.toBoolValue(
        attributes.negativeBarColorSameAsPositive,
        true
      ),
      negativeBarBorderColorSameAsPositive: BaseXform.toBoolValue(
        attributes.negativeBarBorderColorSameAsPositive,
        true
      ),
      axisPosition: BaseXform.toStringValue(attributes.axisPosition, "auto"),
      direction: BaseXform.toStringValue(attributes.direction, "leftToRight")
    };
  }

  onParserClose(name, parser) {
    const [, prop] = name.split(":");
    switch (prop) {
      case "cfvo":
        this.model.cfvo.push(parser.model);
        break;

      default:
        this.model[prop] = parser.model;
        break;
    }
  }
}

export { DatabarExtXform };
