import { renderSparklineGroups, parseSparklineGroups } from "@excel/core/sparkline";
import type { SparklineGroup } from "@excel/core/sparkline";
import type { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { ConditionalFormattingsExtXform } from "@excel/xlsx/xform/sheet/cf-ext/conditional-formattings-ext-xform";
import type { XmlSink } from "@xml/types";

/** The worksheet `<extLst>` model: conditional-formatting and sparkline extensions. */
interface ExtLstModel {
  conditionalFormattings?: unknown;
  sparklineGroups?: SparklineGroup[];
}

class ExtXform extends CompositeXform {
  declare public map: Record<string, BaseXform>;
  declare public model: ExtLstModel;
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

  hasContent(model: ExtLstModel): boolean {
    return this.conditionalFormattings.hasContent(model.conditionalFormattings);
  }

  prepare(model: ExtLstModel): void {
    this.conditionalFormattings.prepare(model.conditionalFormattings);
  }

  render(xmlStream: XmlSink, model: ExtLstModel): void {
    xmlStream.openNode("ext", {
      uri: "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
      "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    });

    this.conditionalFormattings.render(xmlStream, model.conditionalFormattings);

    xmlStream.closeNode();
  }

  createNewModel(): ExtLstModel {
    return {};
  }

  onParserClose(name: string, parser: BaseXform): void {
    (this.model as Record<string, unknown>)[name] = parser.model;
  }
}

/**
 * Lightweight xform for the sparkline extension block.
 * Renders/captures `<ext uri="{05C60535-...}">` containing x14:sparklineGroups.
 */
class SparklineExtXform {
  get tag(): string {
    return "ext";
  }

  hasContent(sparklineGroups: unknown): boolean {
    return Array.isArray(sparklineGroups) && sparklineGroups.length > 0;
  }

  render(xmlStream: XmlSink, sparklineGroups: SparklineGroup[]): void {
    if (!this.hasContent(sparklineGroups)) {
      return;
    }
    // The canonical Microsoft-registered extension uri for
    // `x14:sparklineGroups` is
    //   {05C60535-1F16-4fd2-B633-F4F36F0B64E0}
    // as emitted by Excel 2010 through Excel 365. Earlier
    // revisions of this file used the WRONG uri
    // `{05C60535-1F16-4fd2-B633-F4F36F0041E1}` (a typo that
    // looked plausible but doesn't match any registered
    // extension). Excel's MC processor — which is designed to
    // silently skip unknown extension uris — dropped the whole
    // `<ext>` element on load, erasing every sparkline the
    // workbook defined. Verified against a Microsoft Excel 2021
    // sparkline reference (`tmp/ccccc.xlsx`); switching to the
    // correct uri restored sparkline rendering.
    xmlStream.openNode("ext", {
      uri: "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}",
      "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
    });
    xmlStream.writeRaw(renderSparklineGroups(sparklineGroups));
    xmlStream.closeNode();
  }

  parse(xml: string): SparklineGroup[] {
    return parseSparklineGroups(xml);
  }
}

class ExtLstXform extends CompositeXform {
  declare public map: Record<string, BaseXform>;
  declare public model: ExtLstModel;
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

  prepare(model: ExtLstModel, _options?: unknown): void {
    this.ext.prepare(model);
  }

  hasContent(model: ExtLstModel): boolean {
    return this.ext.hasContent(model) || this.sparklineExt.hasContent(model?.sparklineGroups);
  }

  render(xmlStream: XmlSink, model: ExtLstModel): void {
    if (!this.hasContent(model)) {
      return;
    }

    xmlStream.openNode("extLst");
    if (this.ext.hasContent(model)) {
      this.ext.render(xmlStream, model);
    }
    if (this.sparklineExt.hasContent(model?.sparklineGroups)) {
      this.sparklineExt.render(xmlStream, model.sparklineGroups!);
    }
    xmlStream.closeNode();
  }

  createNewModel(): ExtLstModel {
    return {};
  }

  onParserClose(name: string, parser: BaseXform): void {
    (this.model as Record<string, unknown>)[name] = parser.model;
  }
}

export { ExtLstXform };
