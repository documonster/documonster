import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import type { XmlSink } from "@xml/types";

/** The x14 conditional-formatting extension reference: just the rule's x14 id. */
interface ExtLstRefModel {
  x14Id?: string;
}

class X14IdXform extends BaseXform<string> {
  get tag(): string {
    return "x14:id";
  }

  render(xmlStream: XmlSink, model?: string): void {
    xmlStream.leafNode(this.tag, undefined, model);
  }

  parseOpen(): void {
    this.model = "";
  }

  parseText(text: string): void {
    this.model += text;
  }

  parseClose(name: string): boolean {
    return name !== this.tag;
  }
}

class ExtXform extends CompositeXform {
  idXform: X14IdXform;
  declare public model: ExtLstRefModel;

  constructor() {
    super();

    this.map = {
      "x14:id": (this.idXform = new X14IdXform())
    };
  }

  get tag(): string {
    return "ext";
  }

  render(xmlStream: XmlSink, model: ExtLstRefModel): void {
    xmlStream.openNode(this.tag, {
      uri: "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}",
      "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    });

    this.idXform.render(xmlStream, model.x14Id);

    xmlStream.closeNode();
  }

  createNewModel(): ExtLstRefModel {
    return {};
  }

  onParserClose(name: string, parser: BaseXform): void {
    this.model.x14Id = parser.model as string;
  }
}

class ExtLstRefXform extends CompositeXform {
  declare public model: ExtLstRefModel;

  constructor() {
    super();
    this.map = {
      ext: new ExtXform()
    };
  }

  get tag(): string {
    return "extLst";
  }

  render(xmlStream: XmlSink, model: ExtLstRefModel): void {
    if (!model.x14Id) {
      return;
    }
    xmlStream.openNode(this.tag);
    this.map!.ext.render(xmlStream, model);
    xmlStream.closeNode();
  }

  createNewModel(): ExtLstRefModel {
    return {};
  }

  onParserClose(name: string, parser: BaseXform): void {
    Object.assign(this.model, parser.model);
  }
}

export { ExtLstRefXform };
