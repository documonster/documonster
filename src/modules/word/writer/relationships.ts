/**
 * DOCX Module - Relationship Manager
 *
 * Manages OPC relationships for the DOCX package.
 * Generates .rels files for package-level and part-level relationships.
 *
 * Provides both an object-oriented RelationshipsState interface and
 * backward-compatible free functions (addRelationship, renderRelationships, etc.).
 */

import type { XmlSink } from "@xml/types";

import { NS_PKG_RELS, STD_DOC_ATTRIBUTES } from "../constants";
import { DocxWriteError } from "../errors";

/** A single OPC relationship. */
export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: "External";
}

/** Rich relationship set with validation and dedup capabilities. */
export interface RelationshipsState {
  readonly rels: readonly Relationship[];
  /** Add a relationship. If a matching (type, target, targetMode) already exists, returns its ID. */
  add(type: string, target: string, targetMode?: "External"): string;
  /** Add with a specific ID. Throws if ID already exists. */
  addWithId(id: string, type: string, target: string, targetMode?: "External"): void;
  /** Find existing relationship by type and target. */
  findByTypeAndTarget(type: string, target: string): Relationship | undefined;
  /** Check if an rId is already used. */
  hasId(id: string): boolean;
  /** Get count. */
  count(): number;
  /** Validate all relationships. Returns error messages (empty = valid). */
  validate(): string[];
  /** Render to XML. */
  render(sink: XmlSink): void;
}

/** Internal mutable storage backing RelationshipsState. */
interface RelationshipsInternal {
  readonly _rels: Relationship[];
  _nextId: number;
}

/** Create a new empty RelationshipsState. */
export function createRelationships(): RelationshipsState {
  const internal: RelationshipsInternal = { _rels: [], _nextId: 1 };

  const state: RelationshipsState = {
    get rels(): readonly Relationship[] {
      return internal._rels;
    },

    add(type: string, target: string, targetMode?: "External"): string {
      // Dedup: reuse existing (type, target, targetMode) if found
      const existing = internal._rels.find(
        r => r.type === type && r.target === target && r.targetMode === targetMode
      );
      if (existing) {
        return existing.id;
      }
      const id = `rId${internal._nextId++}`;
      internal._rels.push({ id, type, target, targetMode });
      return id;
    },

    addWithId(id: string, type: string, target: string, targetMode?: "External"): void {
      if (internal._rels.some(r => r.id === id)) {
        throw new DocxWriteError(`Relationship ID "${id}" already exists`);
      }
      internal._rels.push({ id, type, target, targetMode });
      // Keep nextId above any manually-assigned IDs
      const num = parseInt(id.replace("rId", ""), 10);
      if (!isNaN(num) && num >= internal._nextId) {
        internal._nextId = num + 1;
      }
    },

    findByTypeAndTarget(type: string, target: string): Relationship | undefined {
      return internal._rels.find(r => r.type === type && r.target === target);
    },

    hasId(id: string): boolean {
      return internal._rels.some(r => r.id === id);
    },

    count(): number {
      return internal._rels.length;
    },

    validate(): string[] {
      const errors: string[] = [];

      // Check duplicate IDs
      const idSet = new Set<string>();
      for (const rel of internal._rels) {
        if (idSet.has(rel.id)) {
          errors.push(`Duplicate relationship ID: "${rel.id}"`);
        }
        idSet.add(rel.id);
      }

      // Check external targets without TargetMode
      for (const rel of internal._rels) {
        if (
          (rel.target.startsWith("http://") ||
            rel.target.startsWith("https://") ||
            rel.target.startsWith("mailto:")) &&
          !rel.targetMode
        ) {
          errors.push(
            `Relationship "${rel.id}" has external target "${rel.target}" but no TargetMode="External"`
          );
        }
      }

      return errors;
    },

    render(sink: XmlSink): void {
      renderRelationships(state, sink);
    }
  };

  return state;
}

// =============================================================================
// Backward-Compatible Free Functions
// =============================================================================

/** Add a relationship and return its assigned rId (free function alias). */
export function addRelationship(
  state: RelationshipsState,
  type: string,
  target: string,
  targetMode?: "External"
): string {
  return state.add(type, target, targetMode);
}

/** Add a relationship with a specific ID (free function alias). */
export function addRelationshipWithId(
  state: RelationshipsState,
  id: string,
  type: string,
  target: string,
  targetMode?: "External"
): void {
  state.addWithId(id, type, target, targetMode);
}

/** Get the number of relationships (free function alias). */
export function getRelationshipCount(state: RelationshipsState): number {
  return state.count();
}

/** Render the relationships XML to a sink (free function alias). */
export function renderRelationships(state: RelationshipsState, xml: XmlSink): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("Relationships", { xmlns: NS_PKG_RELS });
  for (const rel of state.rels) {
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
