/**
 * DOCX Module - Relationship Manager
 *
 * Manages OPC relationships for the DOCX package.
 * Generates .rels files for package-level and part-level relationships.
 */

import type { XmlSink } from "@xml/types";

import { NS_PKG_RELS, STD_DOC_ATTRIBUTES } from "./constants";

/** A single OPC relationship. */
export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: "External";
}

/**
 * Manages relationships and generates .rels XML.
 */
export class RelationshipManager {
  private readonly _rels: Relationship[] = [];
  private _nextId = 1;

  /** Add a relationship and return its assigned rId. */
  add(type: string, target: string, targetMode?: "External"): string {
    const id = `rId${this._nextId++}`;
    this._rels.push({ id, type, target, targetMode });
    return id;
  }

  /** Add a relationship with a specific ID. */
  addWithId(id: string, type: string, target: string, targetMode?: "External"): void {
    this._rels.push({ id, type, target, targetMode });
    // Keep nextId above any manually-assigned IDs
    const num = parseInt(id.replace("rId", ""), 10);
    if (!isNaN(num) && num >= this._nextId) {
      this._nextId = num + 1;
    }
  }

  /** Get all relationships. */
  get relationships(): readonly Relationship[] {
    return this._rels;
  }

  /** Get the number of relationships. */
  get count(): number {
    return this._rels.length;
  }

  /** Render the relationships XML to a sink. */
  render(xml: XmlSink): void {
    xml.openXml(STD_DOC_ATTRIBUTES);
    xml.openNode("Relationships", { xmlns: NS_PKG_RELS });
    for (const rel of this._rels) {
      const attrs: Record<string, string> = {
        Id: rel.id,
        Type: rel.type,
        Target: rel.target
      };
      if (rel.targetMode) {
        attrs.TargetMode = rel.targetMode;
      }
      xml.leafNode("Relationship", attrs);
    }
    xml.closeNode();
  }
}
