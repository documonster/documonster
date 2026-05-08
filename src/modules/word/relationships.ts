/**
 * DOCX Module - Relationship Manager
 *
 * Manages OPC relationships for the DOCX package.
 * Generates .rels files for package-level and part-level relationships.
 * Uses a plain data record + free functions for tree-shakeability.
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

/** Internal state for relationships (plain record, not a class). */
export interface RelationshipsState {
  readonly rels: Relationship[];
  nextId: number;
}

/** Create a new empty RelationshipsState. */
export function createRelationships(): RelationshipsState {
  return { rels: [], nextId: 1 };
}

/** Add a relationship and return its assigned rId. */
export function addRelationship(
  state: RelationshipsState,
  type: string,
  target: string,
  targetMode?: "External"
): string {
  const id = `rId${state.nextId++}`;
  (state.rels as Relationship[]).push({ id, type, target, targetMode });
  return id;
}

/** Add a relationship with a specific ID. */
export function addRelationshipWithId(
  state: RelationshipsState,
  id: string,
  type: string,
  target: string,
  targetMode?: "External"
): void {
  (state.rels as Relationship[]).push({ id, type, target, targetMode });
  // Keep nextId above any manually-assigned IDs
  const num = parseInt(id.replace("rId", ""), 10);
  if (!isNaN(num) && num >= state.nextId) {
    state.nextId = num + 1;
  }
}

/** Get the number of relationships. */
export function getRelationshipCount(state: RelationshipsState): number {
  return state.rels.length;
}

/** Render the relationships XML to a sink. */
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
