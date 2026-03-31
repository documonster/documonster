import { BaseXform } from "@excel/xlsx/xform/base-xform";

const validation = {
  boolean(value: boolean | undefined, dflt: boolean): boolean {
    if (value === undefined) {
      return dflt;
    }
    return value;
  }
};

interface ProtectionModel {
  locked?: boolean;
  hidden?: boolean;
}

// Protection encapsulates translation from style.protection model to/from xlsx
class ProtectionXform extends BaseXform {
  get tag(): string {
    return "protection";
  }

  render(xmlStream: any, model: ProtectionModel): void {
    const attrs: Record<string, string> = {};
    const locked = validation.boolean(model.locked, true) ? undefined : "0";
    const hidden = validation.boolean(model.hidden, false) ? "1" : undefined;
    if (locked !== undefined) {
      attrs.locked = locked;
    }
    if (hidden !== undefined) {
      attrs.hidden = hidden;
    }

    if (Object.keys(attrs).length > 0) {
      xmlStream.leafNode("protection", attrs);
    }
  }

  parseOpen(node: any): void {
    const model: ProtectionModel = {
      locked: !(node.attributes.locked === "0"),
      hidden: node.attributes.hidden === "1"
    };

    // only want to record models that differ from defaults
    const isSignificant = !model.locked || model.hidden;

    this.model = isSignificant ? model : null;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { ProtectionXform };
