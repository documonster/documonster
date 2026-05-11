/**
 * DOCX Module - Glossary Document (Building Blocks)
 *
 * Provides types and utilities for working with Glossary Document parts,
 * which contain AutoText entries, Quick Parts, and other Building Blocks.
 *
 * INTEGRATION STATUS: This module provides data structures and query helpers
 * for building blocks. The actual reading/writing of glossary parts from/to
 * DOCX archives is handled via the opaqueParts round-trip mechanism —
 * glossary parts in existing files are preserved as opaque parts during
 * read/write. This module is useful for:
 * - Building glossary data structures programmatically
 * - Querying/filtering building block collections
 * - Preparing data for future direct glossary part writing
 *
 * To add glossary content to a document currently, include it as an
 * OpaquePart with path "word/glossary/document.xml".
 */

import { generateGuid } from "../core/internal-utils";
import type { BodyContent, SectionProperties } from "../types";

// =============================================================================
// Types
// =============================================================================

/** Building block gallery category. */
export type BuildingBlockGallery =
  | "autoText"
  | "quickParts"
  | "coverPages"
  | "tableOfContents"
  | "headers"
  | "footers"
  | "pageNumbers"
  | "tables"
  | "textBoxes"
  | "watermarks"
  | "equations"
  | "bibliographies"
  | "custom1"
  | "custom2"
  | "custom3"
  | "custom4"
  | "custom5";

/** A single building block (AutoText/Quick Part) entry. */
export interface BuildingBlock {
  /** Name of the building block (displayed in gallery). */
  readonly name: string;
  /** Gallery this block belongs to. */
  readonly gallery: BuildingBlockGallery;
  /** Category within the gallery. */
  readonly category?: string;
  /** Description/tooltip. */
  readonly description?: string;
  /** The content of the building block. */
  readonly content: readonly BodyContent[];
  /** Section properties specific to this building block. */
  readonly sectionProperties?: SectionProperties;
  /** Unique identifier (GUID). */
  readonly guid?: string;
}

/** The glossary document model. */
export interface GlossaryDocument {
  /** Building block entries. */
  readonly blocks: readonly BuildingBlock[];
  /** Raw parts preserved for round-trip (style, settings, fontTable). */
  readonly rawParts?: ReadonlyMap<string, Uint8Array>;
}

// =============================================================================
// Glossary Document Builder
// =============================================================================

/**
 * Create a building block entry.
 *
 * @param name - Display name of the building block.
 * @param gallery - Which gallery it belongs to.
 * @param content - The body content of the building block.
 * @param options - Additional options.
 * @returns A BuildingBlock instance.
 *
 * @example
 * ```ts
 * const autoText = createBuildingBlock("Greeting", "autoText", [
 *   paragraph([text("Dear Sir/Madam,")])
 * ]);
 * ```
 */
export function createBuildingBlock(
  name: string,
  gallery: BuildingBlockGallery,
  content: readonly BodyContent[],
  options?: {
    category?: string;
    description?: string;
    sectionProperties?: SectionProperties;
  }
): BuildingBlock {
  return {
    name,
    gallery,
    category: options?.category ?? "General",
    description: options?.description,
    content,
    sectionProperties: options?.sectionProperties,
    guid: generateGuid()
  };
}

/**
 * Create a glossary document from building blocks.
 */
export function createGlossaryDocument(blocks: readonly BuildingBlock[]): GlossaryDocument {
  return { blocks };
}

/**
 * Find a building block by name and gallery.
 */
export function findBuildingBlock(
  glossary: GlossaryDocument,
  name: string,
  gallery?: BuildingBlockGallery
): BuildingBlock | undefined {
  return glossary.blocks.find(
    b => b.name === name && (gallery === undefined || b.gallery === gallery)
  );
}

/**
 * List all building blocks in a specific gallery.
 */
export function listBuildingBlocks(
  glossary: GlossaryDocument,
  gallery: BuildingBlockGallery
): readonly BuildingBlock[] {
  return glossary.blocks.filter(b => b.gallery === gallery);
}

/**
 * Extract AutoText entries from a glossary document.
 */
export function getAutoTextEntries(glossary: GlossaryDocument): readonly BuildingBlock[] {
  return listBuildingBlocks(glossary, "autoText");
}

/**
 * Extract Quick Parts from a glossary document.
 */
export function getQuickParts(glossary: GlossaryDocument): readonly BuildingBlock[] {
  return listBuildingBlocks(glossary, "quickParts");
}
