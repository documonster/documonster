import { parseSax } from "@xml/sax";
import { XmlWriter } from "@xml/writer";
import type { XmlSink } from "@xml/types";

/* 'virtual' methods used as a form of documentation */

interface ParseEvent {
  eventType: string;
  value: any;
}

// HAN CELL namespace prefix normalization
// HAN CELL uses non-standard namespace prefixes (ep:, cp:, dc:, etc.)
const HAN_CELL_PREFIXES = new Set(["ep", "cp", "dc", "dcterms", "dcmitype", "vt"]);
const SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

// Detect HAN CELL mode from first tag. Returns:
// - undefined: normal file (no prefix handling needed)
// - null: HAN CELL file without spreadsheetml prefix (uses static prefixes only)
// - string: HAN CELL file with spreadsheetml prefix (e.g., "x")
function detectHanCellPrefix(
  tagName: string,
  attrs: Record<string, string>
): string | null | undefined {
  for (const key in attrs) {
    if (key.length > 6 && key.startsWith("xmlns:")) {
      const prefix = key.slice(6);
      // Check for spreadsheetml namespace prefix
      if (attrs[key] === SPREADSHEETML_NS) {
        return prefix;
      }
      // Check if xmlns declares a known HAN CELL prefix (e.g., xmlns:dc, xmlns:dcterms)
      if (HAN_CELL_PREFIXES.has(prefix)) {
        return null;
      }
    }
  }
  // Check if tag name has a known static prefix
  const i = tagName.indexOf(":");
  return i !== -1 && HAN_CELL_PREFIXES.has(tagName.slice(0, i)) ? null : undefined;
}

// Strip known namespace prefix from element name
function stripPrefix(name: string, nsPrefix: string | null): string {
  const i = name.indexOf(":");
  if (i === -1) {
    return name;
  }
  const p = name.slice(0, i);
  return p === nsPrefix || HAN_CELL_PREFIXES.has(p) ? name.slice(i + 1) : name;
}

// Base class for Xforms
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

  parseOpen(_node: any): void {
    // XML node opened
  }

  parseText(_text: string): void {
    // chunk of text encountered for current node
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

  mergeModel(obj: any): void {
    // set obj's props to this.model
    this.model = Object.assign(this.model || ({} as any), obj);
  }

  async parse(saxParser: AsyncIterable<ParseEvent[]>): Promise<TModel | undefined> {
    // IMPORTANT:
    // Do not return early once parsing is "done".
    // In true streaming scenarios, `parseSax(stream)` is backed by a Node.js
    // Readable async iterator. Returning early would close the iterator, which
    // destroys the underlying stream and can surface as AbortError (ABORT_ERR).
    let done = false;
    let finalModel: TModel | undefined;

    // HAN CELL compatibility: 0 = not checked, 1 = normal file, 2 = HAN CELL file
    let nsMode = 0;
    let nsPrefix: string | null = null;

    for await (const events of saxParser) {
      if (done) {
        continue;
      }
      for (const { eventType, value } of events) {
        if (eventType === "opentag") {
          // Fast path for normal Excel files (majority case)
          if (nsMode === 1) {
            this.parseOpen(value);
            continue;
          }
          // First tag - detect mode
          if (nsMode === 0) {
            const prefix = detectHanCellPrefix(value.name, value.attributes);
            if (prefix === undefined) {
              nsMode = 1;
              this.parseOpen(value);
              continue;
            }
            nsMode = 2;
            nsPrefix = prefix;
          }
          // HAN CELL mode - strip prefix
          value.name = stripPrefix(value.name, nsPrefix);
          this.parseOpen(value);
        } else if (eventType === "text") {
          this.parseText(value);
        } else if (eventType === "closetag") {
          // Fast path for normal files
          if (nsMode === 1) {
            if (!this.parseClose(value.name)) {
              done = true;
              finalModel = this.model;
              break;
            }
            continue;
          }
          // HAN CELL mode - strip prefix
          if (!this.parseClose(stripPrefix(value.name, nsPrefix))) {
            done = true;
            finalModel = this.model;
            break;
          }
        }
      }
    }

    return done ? finalModel : this.model;
  }

  async parseStream(stream: any): Promise<TModel | undefined> {
    return this.parse(parseSax(stream));
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
