/**
 * DOCX Reader Context
 *
 * Encapsulates per-parse session state so that concurrent readDocx() calls
 * are safe. Previously this state lived in a module-level `_session` object.
 */

import { parseXml } from "@xml/dom";

import { STRICT_TO_TRANSITIONAL_REL } from "../constants";

/**
 * A parsed OPC relationship entry (from a .rels part).
 */
export interface ParsedRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

/**
 * Per-parse session context. Holds relationship maps and any other state
 * that needs to be shared across parse functions during a single readDocx()
 * invocation without relying on module-level mutable state.
 */
export interface ReaderContext {
  /** Relationship map for the current part being parsed (document, header, footer, etc.) */
  relMap: Map<string, ParsedRelationship>;
}

/**
 * Create a fresh reader context for a new parse session.
 */
export function createReaderContext(): ReaderContext {
  return {
    relMap: new Map()
  };
}

// =============================================================================
// Relationships Parser
// =============================================================================

/**
 * Parse a `.rels` XML string into an array of `ParsedRelationship` entries.
 *
 * Normalizes ISO 29500 Strict relationship types to Transitional equivalents
 * via `STRICT_TO_TRANSITIONAL_REL`.
 */
export function parseRelationships(xmlStr: string): ParsedRelationship[] {
  const doc = parseXml(xmlStr);
  const rels: ParsedRelationship[] = [];

  for (const child of doc.root.children) {
    if (child.type === "element" && child.name === "Relationship") {
      let relType = child.attributes["Type"] ?? "";
      // Normalize ISO 29500 Strict relationship types to Transitional
      const transitional = STRICT_TO_TRANSITIONAL_REL.get(relType);
      if (transitional) {
        relType = transitional;
      }
      rels.push({
        id: child.attributes["Id"] ?? "",
        type: relType,
        target: child.attributes["Target"] ?? "",
        targetMode: child.attributes["TargetMode"]
      });
    }
  }

  return rels;
}
