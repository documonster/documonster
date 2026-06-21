import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { DateXform } from "@excel/xlsx/xform/simple/date-xform";
import { IntegerXform } from "@excel/xlsx/xform/simple/integer-xform";
import { StringXform } from "@excel/xlsx/xform/simple/string-xform";
import type { XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

interface CoreModel {
  creator?: string;
  title?: string;
  subject?: string;
  description?: string;
  identifier?: string;
  language?: string;
  keywords?: string;
  category?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  revision?: number;
  version?: string;
  contentStatus?: string;
  contentType?: string;
  created?: Date;
  modified?: Date;
}

// Rendering uses namespace prefixes, parsing uses unqualified names (SAX strips prefixes)
const PROPS = {
  creator: "dc:creator",
  title: "dc:title",
  subject: "dc:subject",
  description: "dc:description",
  identifier: "dc:identifier",
  language: "dc:language",
  keywords: "cp:keywords",
  category: "cp:category",
  lastModifiedBy: "cp:lastModifiedBy",
  lastPrinted: "cp:lastPrinted",
  revision: "cp:revision",
  version: "cp:version",
  contentStatus: "cp:contentStatus",
  contentType: "cp:contentType",
  created: "dcterms:created",
  modified: "dcterms:modified"
} as const;

class CoreXform extends BaseXform {
  declare public map: { [key: string]: any };
  declare public parser: any;

  constructor() {
    super();

    this.map = {
      creator: new StringXform({ tag: PROPS.creator }),
      title: new StringXform({ tag: PROPS.title }),
      subject: new StringXform({ tag: PROPS.subject }),
      description: new StringXform({ tag: PROPS.description }),
      identifier: new StringXform({ tag: PROPS.identifier }),
      language: new StringXform({ tag: PROPS.language }),
      keywords: new StringXform({ tag: PROPS.keywords }),
      category: new StringXform({ tag: PROPS.category }),
      lastModifiedBy: new StringXform({ tag: PROPS.lastModifiedBy }),
      lastPrinted: new DateXform({ tag: PROPS.lastPrinted, format: CoreXform.DateFormat }),
      revision: new IntegerXform({ tag: PROPS.revision }),
      version: new StringXform({ tag: PROPS.version }),
      contentStatus: new StringXform({ tag: PROPS.contentStatus }),
      contentType: new StringXform({ tag: PROPS.contentType }),
      created: new DateXform({
        tag: PROPS.created,
        attrs: CoreXform.DateAttrs,
        format: CoreXform.DateFormat
      }),
      modified: new DateXform({
        tag: PROPS.modified,
        attrs: CoreXform.DateAttrs,
        format: CoreXform.DateFormat
      })
    };
  }

  render(xmlStream: XmlSink, model: CoreModel): void {
    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode("cp:coreProperties", CoreXform.CORE_PROPERTY_ATTRIBUTES);

    for (const key of Object.keys(PROPS)) {
      this.map[key].render(xmlStream, model[key as keyof CoreModel]);
    }

    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (node.name !== "coreProperties") {
      this.parser = this.map[node.name];
      if (this.parser) {
        this.parser.parseOpen(node);
      }
    }
    return true;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    if (name === "coreProperties") {
      this.model = {};
      for (const key of Object.keys(PROPS)) {
        const val = this.map[key].model;
        if (val !== undefined && val !== "") {
          this.model[key] = val;
        }
      }
      return false;
    }
    return true;
  }

  static DateFormat(dt: Date): string {
    return dt.toISOString().replace(/[.]\d{3}/, "");
  }

  static DateAttrs = { "xsi:type": "dcterms:W3CDTF" };

  static CORE_PROPERTY_ATTRIBUTES = {
    "xmlns:cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "xmlns:dc": "http://purl.org/dc/elements/1.1/",
    "xmlns:dcterms": "http://purl.org/dc/terms/",
    "xmlns:dcmitype": "http://purl.org/dc/dcmitype/",
    "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
  };
}

export { CoreXform };
