import type {
  Color,
  ConditionalFormattingOptions,
  ConditionalFormattingRule,
  Style
} from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { ConditionalFormattingXform } from "@excel/xlsx/xform/sheet/cf/conditional-formatting-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

/** A CF rule as carried through (de)serialisation: it gains a transient dxfId. */
type SerializedCfRule = ConditionalFormattingRule & {
  dxfId?: number;
  cfvo?: { type: string }[];
  color?: Partial<Color>;
};

/** The style-manager surface the CF xform needs from the shared writer options. */
interface CfStyleManager {
  addDxfStyle(style: Partial<Style>): number;
  getDxfStyle(id: number): Partial<Style> | undefined;
}
interface CfPrepareOptions {
  styles: object;
}

class ConditionalFormattingsXform extends BaseXform<ConditionalFormattingOptions[]> {
  cfXform: ConditionalFormattingXform;
  parser?: BaseXform;

  constructor() {
    super();

    this.cfXform = new ConditionalFormattingXform();
  }

  get tag(): string {
    return "conditionalFormatting";
  }

  reset(): void {
    this.model = [];
  }

  prepare(model: ConditionalFormattingOptions[], options: CfPrepareOptions): void {
    const styles = options.styles as CfStyleManager;
    // ensure each rule has a priority value
    let nextPriority = model.reduce(
      (p: number, cf) => Math.max(p, ...cf.rules.map(rule => rule.priority ?? 0)),
      1
    );
    model.forEach(cf => {
      cf.rules.forEach(ruleModel => {
        const rule = ruleModel as SerializedCfRule;
        if (!rule.priority) {
          rule.priority = nextPriority++;
        }

        if (rule.style) {
          rule.dxfId = styles.addDxfStyle(rule.style);
        }

        // Ensure dataBar rules have required cfvo and color properties
        if (rule.type === "dataBar") {
          if (!rule.cfvo || rule.cfvo.length < 2) {
            rule.cfvo = [{ type: "min" }, { type: "max" }];
          }
          if (!rule.color) {
            // Default blue color for data bars (same as Excel's default)
            rule.color = { argb: "FF638EC6" };
          }
        }
      });
    });
  }

  render(xmlStream: XmlSink, model?: ConditionalFormattingOptions[]): void {
    (model ?? []).forEach(cf => {
      this.cfXform.render(xmlStream, cf);
    });
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case "conditionalFormatting":
        this.parser = this.cfXform;
        this.parser.parseOpen(node);
        return true;

      default:
        return false;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.push(this.parser.model as ConditionalFormattingOptions);
        this.parser = undefined;
        return false;
      }
      return true;
    }
    return false;
  }

  reconcile(model: ConditionalFormattingOptions[], options: CfPrepareOptions): void {
    const styles = options.styles as CfStyleManager;
    model.forEach(cf => {
      cf.rules.forEach(ruleModel => {
        const rule = ruleModel as SerializedCfRule;
        if (rule.dxfId !== undefined) {
          rule.style = styles.getDxfStyle(rule.dxfId);
          delete rule.dxfId;
        }
      });
    });
  }
}

export { ConditionalFormattingsXform };
