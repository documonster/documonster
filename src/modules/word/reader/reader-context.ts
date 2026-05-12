/**
 * DOCX Reader Context
 *
 * Encapsulates per-parse session state so that concurrent readDocx() calls
 * are safe. Previously this state lived in a module-level `_session` object.
 */

import { parseXml } from "@xml/dom";

import { STRICT_TO_TRANSITIONAL_REL } from "../constants";
import { DEFAULT_SECURITY_POLICY, type WordSecurityPolicy } from "../security/policy";
import type { FormField, RunProperties } from "../types";

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
 * State for the OOXML field state machine (`<w:fldChar>`).
 *
 * OOXML allows complex fields (TOC, INDEX, SEQ, REF, …) to span multiple
 * paragraphs: `<w:fldChar fldCharType="begin">` may be in paragraph A while
 * the matching `end` is in paragraph C. This state therefore lives on the
 * shared reader context rather than as locals in `parseParagraph`.
 *
 * It is intentionally swappable: when entering a self-contained part such as
 * a header, footer, footnote, endnote or comment we save and reset the field
 * state so an unterminated field in the body never bleeds into those parts.
 */
export interface FieldParseState {
  state: "none" | "instrText" | "cached";
  instr: string;
  cached: string;
  runProps: RunProperties | undefined;
  formField: FormField | undefined;
}

/**
 * Per-parse session context. Holds relationship maps and any other state
 * that needs to be shared across parse functions during a single readDocx()
 * invocation without relying on module-level mutable state.
 */
export interface ReaderContext {
  /** Relationship map for the current part being parsed (document, header, footer, etc.) */
  relMap: Map<string, ParsedRelationship>;
  /** Field state machine; shared across paragraphs to support cross-paragraph fields. */
  field: FieldParseState;
  /**
   * Resolved security policy for this read. Parsers consult specific
   * fields (e.g. `allowExternalTargets`) to decide whether to keep certain
   * pieces of information that may be dangerous in untrusted documents.
   * Defaults to `DEFAULT_SECURITY_POLICY` when not set.
   */
  securityPolicy: Required<WordSecurityPolicy>;
}

/**
 * Create a fresh field state in the "none" state.
 */
export function createFieldState(): FieldParseState {
  return {
    state: "none",
    instr: "",
    cached: "",
    runProps: undefined,
    formField: undefined
  };
}

/**
 * Create a fresh reader context for a new parse session.
 */
export function createReaderContext(securityPolicy?: Required<WordSecurityPolicy>): ReaderContext {
  return {
    relMap: new Map(),
    field: createFieldState(),
    securityPolicy: securityPolicy ?? DEFAULT_SECURITY_POLICY
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
