import type { XmlSink } from "@xml/types";
import { XmlWriter } from "@xml/writer";

/* 'virtual' methods used as a form of documentation */

/**
 * Base class for Xforms.
 *
 * Holds the model and the per-element callbacks (`parseOpen`, `parseText`, `parseClose`) and the renderer.
 * The loop that *drives* a parse lives in `xform/parse-xform.ts` rather than here, because it is the one
 * piece that needs the SAX parser: a method on this class would be retained with every xform the writer
 * renders, and the writer would carry the XML parser along with them.
 */
class BaseXform<TModel = any> {
  declare public map?: { [key: string]: any };
  public model?: TModel;

  // ============================================================
  // Virtual Interface
  prepare(_model?: any, _options?: any): void {
    // optional preparation (mutation) of model so it is ready for write
  }

  render(_xmlStream?: XmlSink, _model?: any): void {
    // convert model to xml
  }

  parseOpen(_node: any): boolean | void {
    // XML node opened. Subclasses that participate in a composite parse return
    // a boolean indicating whether they consumed the node; leaf/standalone
    // xforms return nothing (void).
    return false;
  }

  parseText(_text: string): void {
    // chunk of text encountered for current node
  }

  parseCdata(text: string): void {
    // CDATA and ordinary text have the same parsed value. Xforms that need to
    // preserve the lexical CDATA wrapper may override this method.
    this.parseText(text);
  }

  parseClose(_name: string): boolean {
    // XML node closed
    return false;
  }

  reconcile(_model: any, _options?: any): void {
    // optional post-parse step (opposite to prepare)
  }

  // ============================================================
  reset(): void {
    // to make sure parses don't bleed to next iteration
    this.model = undefined;

    // if we have a map - reset them too
    if (this.map) {
      Object.values(this.map).forEach(xform => {
        if (xform instanceof BaseXform) {
          xform.reset();
        } else if (xform.xform) {
          xform.xform.reset();
        }
      });
    }
  }

  mergeModel(obj: Partial<TModel>): void {
    // set obj's props to this.model
    const base: object = (this.model as object | undefined) ?? {};
    this.model = Object.assign(base, obj) as TModel;
  }

  get xml(): string {
    // convenience function to get the xml of this.model
    // useful for manager types that are built during the prepare phase
    return this.toXml(this.model);
  }

  toXml(model?: any): string {
    const xmlStream = new XmlWriter();
    this.render(xmlStream, model);
    return xmlStream.xml;
  }

  // ============================================================
  // Useful Utilities
  static toAttribute(value: any, dflt?: any, always: boolean = false): string | undefined {
    if (value === undefined) {
      if (always) {
        return dflt;
      }
    } else if (always || value !== dflt) {
      return value.toString();
    }
    return undefined;
  }

  static toStringAttribute(value: any, dflt?: any, always: boolean = false): string | undefined {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toStringValue(attr: any, dflt?: any): any {
    return attr === undefined ? dflt : attr;
  }

  static toBoolAttribute(value: any, dflt?: any, always: boolean = false): string | undefined {
    if (value === undefined) {
      if (always) {
        return dflt;
      }
    } else if (always || value !== dflt) {
      return value ? "1" : "0";
    }
    return undefined;
  }

  static toBoolValue(attr: any, dflt?: any): boolean {
    return attr === undefined ? dflt : attr === "1";
  }

  static toIntAttribute(value: any, dflt?: any, always: boolean = false): string | undefined {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toIntValue(attr: any, dflt?: any): number {
    return attr === undefined ? dflt : parseInt(attr, 10);
  }

  static toFloatAttribute(value: any, dflt?: any, always: boolean = false): string | undefined {
    return BaseXform.toAttribute(value, dflt, always);
  }

  static toFloatValue(attr: any, dflt?: any): number {
    return attr === undefined ? dflt : parseFloat(attr);
  }
}

export { BaseXform };
