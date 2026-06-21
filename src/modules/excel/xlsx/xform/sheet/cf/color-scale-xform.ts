import type { Color, Cvfo } from "@excel/types";
import type { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { CfvoXform } from "@excel/xlsx/xform/sheet/cf/cfvo-xform";
import { ColorXform } from "@excel/xlsx/xform/style/color-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface ColorScaleModel {
  cfvo: Cvfo[];
  color: Partial<Color>[];
}

class ColorScaleXform extends CompositeXform<ColorScaleModel> {
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

  render(xmlStream: XmlSink, model: ColorScaleModel): void {
    xmlStream.openNode(this.tag);

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });
    model.color.forEach(color => {
      this.colorXform.render(xmlStream, color);
    });

    xmlStream.closeNode();
  }

  createNewModel(_node?: ParseOpenTag): ColorScaleModel {
    return {
      cfvo: [],
      color: []
    };
  }

  onParserClose(name: string, parser: BaseXform): void {
    (this.model as unknown as Record<string, unknown[]>)[name].push(parser.model);
  }
}

export { ColorScaleXform };
