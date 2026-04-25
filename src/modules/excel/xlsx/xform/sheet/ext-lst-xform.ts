import { renderSparklineGroups, parseSparklineGroups } from "@excel/sparkline";
import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { ConditionalFormattingsExtXform } from "@excel/xlsx/xform/sheet/cf-ext/conditional-formattings-ext-xform";

class ExtXform extends CompositeXform {
  declare public map: { [key: string]: any };
  declare public model: any;
  declare private conditionalFormattings: ConditionalFormattingsExtXform;

  constructor() {
    super();
    this.map = {
      "x14:conditionalFormattings": (this.conditionalFormattings =
        new ConditionalFormattingsExtXform())
    };
  }

  get tag(): string {
    return "ext";
  }

  hasContent(model: any): boolean {
    return this.conditionalFormattings.hasContent(model.conditionalFormattings);
  }

  prepare(model: any): void {
    this.conditionalFormattings.prepare(model.conditionalFormattings);
  }

  render(xmlStream: any, model: any): void {
    xmlStream.openNode("ext", {
      uri: "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
      "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    });

    this.conditionalFormattings.render(xmlStream, model.conditionalFormattings);

    xmlStream.closeNode();
  }

  createNewModel(): any {
    return {};
  }

  onParserClose(name: string, parser: any): void {
    this.model[name] = parser.model;
  }
}

/**
 * Lightweight xform for the sparkline extension block.
 * Renders/captures `<ext uri="{05C60535-...}">` containing x14:sparklineGroups.
 */
class SparklineExtXform {
  private _parsedModel: any = null;

  get tag(): string {
    return "ext";
  }

  hasContent(sparklineGroups: unknown): boolean {
    return Array.isArray(sparklineGroups) && sparklineGroups.length > 0;
  }

  render(xmlStream: any, sparklineGroups: any[]): void {
    if (!this.hasContent(sparklineGroups)) {
      return;
    }
    xmlStream.openNode("ext", {
      uri: "{05C60535-1F16-4fd2-B633-F4F36F0041E1}",
      "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    });
    xmlStream.writeRaw(renderSparklineGroups(sparklineGroups));
    xmlStream.closeNode();
  }

  parse(xml: string): any[] {
    return parseSparklineGroups(xml);
  }
}

class ExtLstXform extends CompositeXform {
  declare public map: { [key: string]: any };
  declare public model: any;
  declare private ext: ExtXform;
  declare private sparklineExt: SparklineExtXform;

  constructor() {
    super();
    this.sparklineExt = new SparklineExtXform();

    this.map = {
      ext: (this.ext = new ExtXform())
    };
  }

  get tag(): string {
    return "extLst";
  }

  prepare(model: any, _options?: any): void {
    this.ext.prepare(model);
  }

  hasContent(model: any): boolean {
    return this.ext.hasContent(model) || this.sparklineExt.hasContent(model?.sparklineGroups);
  }

  render(xmlStream: any, model: any): void {
    if (!this.hasContent(model)) {
      return;
    }

    xmlStream.openNode("extLst");
    if (this.ext.hasContent(model)) {
      this.ext.render(xmlStream, model);
    }
    if (this.sparklineExt.hasContent(model?.sparklineGroups)) {
      this.sparklineExt.render(xmlStream, model.sparklineGroups);
    }
    xmlStream.closeNode();
  }

  createNewModel(): any {
    return {};
  }

  onParserClose(name: string, parser: any): void {
    this.model[name] = parser.model;
  }
}

export { ExtLstXform };
