import type { ConditionalFormattingOptions, ConditionalFormattingRule } from "@excel/types";
import type { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CompositeXform } from "@excel/xlsx/xform/composite-xform";
import { CfRuleXform } from "@excel/xlsx/xform/sheet/cf/cf-rule-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

class ConditionalFormattingXform extends CompositeXform<ConditionalFormattingOptions> {
  constructor() {
    super();

    this.map = {
      cfRule: new CfRuleXform()
    };
  }

  get tag(): string {
    return "conditionalFormatting";
  }

  render(xmlStream: XmlSink, model: ConditionalFormattingOptions): void {
    // if there are no primitive rules, exit now
    if (!model.rules.some(CfRuleXform.isPrimitive)) {
      return;
    }

    xmlStream.openNode(this.tag, { sqref: model.ref });

    model.rules.forEach(rule => {
      if (CfRuleXform.isPrimitive(rule)) {
        (rule as ConditionalFormattingRule & { ref?: string }).ref = model.ref;
        this.map!.cfRule.render(xmlStream, rule);
      }
    });

    xmlStream.closeNode();
  }

  createNewModel({ attributes }: ParseOpenTag): ConditionalFormattingOptions {
    return {
      ref: attributes.sqref,
      rules: []
    };
  }

  onParserClose(name: string, parser: BaseXform): void {
    this.model!.rules.push(parser.model as ConditionalFormattingRule);
  }
}

export { ConditionalFormattingXform };
