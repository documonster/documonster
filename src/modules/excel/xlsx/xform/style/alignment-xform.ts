import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { validInt, parseBoolean } from "@utils/utils";
import type { ParseOpenTag, XmlSink } from "@xml/types";

const Enums = {
  ReadingOrder: {
    LeftToRight: 1,
    RightToLeft: 2
  }
};

interface AlignmentModel {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  textRotation?: number | "vertical";
  readingOrder?: "ltr" | "rtl" | "context";
  justifyLastLine?: boolean;
  relativeIndent?: number;
}

const validation = {
  horizontalValues: [
    "general",
    "left",
    "center",
    "right",
    "fill",
    "centerContinuous",
    "distributed",
    "justify"
  ].reduce((p: { [key: string]: boolean }, v: string) => {
    p[v] = true;
    return p;
  }, {}),
  horizontal(value: string): string | undefined {
    return this.horizontalValues[value] ? value : undefined;
  },

  verticalValues: ["top", "middle", "bottom", "distributed", "justify"].reduce(
    (p: { [key: string]: boolean }, v: string) => {
      p[v] = true;
      return p;
    },
    {}
  ),
  vertical(value: string): string | undefined {
    if (value === "middle") {
      return "center";
    }
    return this.verticalValues[value] ? value : undefined;
  },
  wrapText(value: boolean): boolean | undefined {
    return value ? true : undefined;
  },
  shrinkToFit(value: boolean): boolean | undefined {
    return value ? true : undefined;
  },
  textRotation(value: number | "vertical"): number | "vertical" | undefined {
    switch (value) {
      case "vertical":
        return value;
      default: {
        const numValue = validInt(value as number);
        return numValue !== undefined && numValue >= -90 && numValue <= 90 ? numValue : undefined;
      }
    }
  },
  indent(value: number): number {
    const numValue = validInt(value);
    return Math.max(0, numValue!);
  },
  readingOrder(value: "ltr" | "rtl" | "context"): number | undefined {
    switch (value) {
      case "context":
        return 0;
      case "ltr":
        return Enums.ReadingOrder.LeftToRight;
      case "rtl":
        return Enums.ReadingOrder.RightToLeft;
      default:
        return undefined;
    }
  }
};

const textRotationXform = {
  toXml(textRotation: number | "vertical"): number | undefined {
    const validated = validation.textRotation(textRotation);
    if (validated) {
      if (validated === "vertical") {
        return 255;
      }

      const tr = Math.round(validated);
      if (tr >= 0 && tr <= 90) {
        return tr;
      }

      if (tr < 0 && tr >= -90) {
        return 90 - tr;
      }
    }
    return undefined;
  },
  toModel(textRotation: string): number | "vertical" | undefined {
    const tr = validInt(textRotation);
    if (tr !== undefined) {
      if (tr === 255) {
        return "vertical";
      }
      if (tr >= 0 && tr <= 90) {
        return tr;
      }
      if (tr > 90 && tr <= 180) {
        return 90 - tr;
      }
    }
    return undefined;
  }
};

/** `ST_ReadingOrder`: 0 context, 1 left-to-right, 2 right-to-left. */
const READING_ORDER_NAMES: Readonly<Record<string, "context" | "ltr" | "rtl">> = {
  "0": "context",
  "1": "ltr",
  "2": "rtl"
};

// Alignment encapsulates translation from style.alignment model to/from xlsx
class AlignmentXform extends BaseXform {
  /**
   * Whether this is the alignment of a differential format (`<dxf>`).
   *
   * A cell's alignment can drop a default — no `wrapText` is the same as `wrapText="0"` — so it does, and its
   * output stays as it always was. A differential format cannot: it changes only what it states, so
   * `wrapText="0"`, `indent="0"`, `textRotation="0"`, `horizontal="general"` and `readingOrder="0"` are each an
   * instruction to *reset* that property, and leaving them out turned the instruction into "change nothing".
   */
  declare private differential: boolean;

  constructor(differential = false) {
    super();
    this.differential = differential;
  }

  get tag(): string {
    return "alignment";
  }

  render(xmlStream: XmlSink, model: AlignmentModel): void {
    // Collect valid attributes first, only write if any exist
    const attrs: Record<string, string | number> = {};
    const differential = this.differential;
    function add(name: string, value: string | number | boolean | undefined): void {
      if (differential ? value !== undefined && value !== false : value) {
        attrs[name] = value as string | number;
      }
    }
    const flag = (value: boolean | undefined): string | false | undefined =>
      value === undefined ? undefined : value ? "1" : differential ? "0" : false;
    // `general` is the default: a cell never states it, a differential format may.
    add(
      "horizontal",
      model.horizontal === "general" && !differential
        ? undefined
        : validation.horizontal(model.horizontal!)
    );
    add("vertical", validation.vertical(model.vertical!));
    add("wrapText", flag(model.wrapText));
    add("shrinkToFit", flag(model.shrinkToFit));
    add("indent", model.indent === undefined ? undefined : validation.indent(model.indent));
    add(
      "relativeIndent",
      differential && model.relativeIndent !== undefined
        ? validInt(model.relativeIndent)
        : undefined
    );
    add("justifyLastLine", flag(model.justifyLastLine));
    add(
      "textRotation",
      model.textRotation === 0 ? 0 : textRotationXform.toXml(model.textRotation!)
    );
    add(
      "readingOrder",
      model.readingOrder === "context" && !differential
        ? undefined
        : validation.readingOrder(model.readingOrder!)
    );
    // A cell's zeros are defaults and are dropped, exactly as before; see `differential`.
    if (!differential) {
      for (const [name, value] of Object.entries(attrs)) {
        if (value === 0) {
          delete attrs[name];
        }
      }
    }

    if (Object.keys(attrs).length > 0) {
      xmlStream.leafNode("alignment", attrs);
    }
  }

  parseOpen(node: ParseOpenTag): void {
    const model: Record<string, unknown> = {};
    const { attributes } = node;

    let valid = false;
    function add(present: unknown, name: string, value: unknown): void {
      if (present !== undefined && value !== undefined) {
        model[name] = value;
        valid = true;
      }
    }
    add(attributes.horizontal, "horizontal", attributes.horizontal);
    add(
      attributes.vertical,
      "vertical",
      attributes.vertical === "center" ? "middle" : attributes.vertical
    );
    add(attributes.wrapText, "wrapText", parseBoolean(attributes.wrapText));
    add(attributes.shrinkToFit, "shrinkToFit", parseBoolean(attributes.shrinkToFit));
    add(attributes.justifyLastLine, "justifyLastLine", parseBoolean(attributes.justifyLastLine));
    add(attributes.indent, "indent", validInt(attributes.indent));
    add(attributes.relativeIndent, "relativeIndent", validInt(attributes.relativeIndent));
    add(
      attributes.textRotation,
      "textRotation",
      textRotationXform.toModel(attributes.textRotation)
    );
    // `0` is *context*, not left-to-right: reading every value but "2" as `ltr` turned a reset into an override.
    add(attributes.readingOrder, "readingOrder", READING_ORDER_NAMES[attributes.readingOrder]);

    this.model = valid ? model : null;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { AlignmentXform };
