/**
 * DOCX Module - Glossary Document (Building Blocks)
 *
 * Provides types and utilities for working with Glossary Document parts,
 * which contain AutoText entries, Quick Parts, and other Building Blocks.
 *
 * INTEGRATION STATUS: This module provides the glossary data model and query
 * helpers. To embed a glossary in a document, assign a {@link GlossaryDocument}
 * to `doc.glossary`; the packager then serialises it to
 * `word/glossary/document.xml`, registers the `glossaryDocument` relationship,
 * and adds the `[Content_Types].xml` override (the canonical OOXML location
 * Word reads Quick Parts / AutoText from). `Document.read` reads a file's glossary
 * into the same model — its entries as `blocks`, which these helpers list and
 * find — and writes it back unchanged unless `blocks` were edited. This module is useful for:
 * - Building glossary data structures programmatically
 * - Querying/filtering building block collections
 * - Assembling a glossary to attach via `doc.glossary`
 */

import {
  checkHandle,
  checkNodeArray,
  checkOneOf,
  checkRecord,
  checkString,
  invalidArgument
} from "@word/core/argument-checks";
import { GALLERY_TO_OOXML } from "@word/core/glossary-galleries";
import { generateGuid } from "@word/core/internal-utils";
import type {
  BodyContent,
  BuildingBlock,
  BuildingBlockGallery,
  GlossaryDocument,
  SectionProperties
} from "@word/types";

// =============================================================================
// Types
// =============================================================================

// The glossary data model (BuildingBlockGallery / BuildingBlock /
// GlossaryDocument) lives in ../types so that DocxDocument can reference it
// without an advanced/ → types cycle. Re-exported here for API compatibility.
export type { BuildingBlock, BuildingBlockGallery, GlossaryDocument } from "@word/types";

/** Every {@link BuildingBlockGallery}. */
const GALLERIES: ReadonlySet<string> = new Set(Object.keys(GALLERY_TO_OOXML));

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
  const where = "Glossary.createBlock";
  checkString(name, "name", where);
  if (name === "") {
    invalidArgument(where, "name", "a non-empty string", name);
  }
  checkOneOf(gallery, GALLERIES, "gallery", where);
  checkNodeArray(content, "content", where);
  checkRecord(options, "options", where, { optional: true });
  checkString(options?.category, "options.category", where, { optional: true });
  checkString(options?.description, "options.description", where, { optional: true });
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
  checkNodeArray(blocks, "blocks", "Glossary.createDocument");
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
  checkHandle(glossary, "glossary", "glossary", "Glossary.findBlock");
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
  checkHandle(glossary, "glossary", "glossary", "Glossary.listBlocks");
  return glossary.blocks.filter(b => b.gallery === gallery);
}

/**
 * Extract AutoText entries from a glossary document.
 */
export function getAutoTextEntries(glossary: GlossaryDocument): readonly BuildingBlock[] {
  checkHandle(glossary, "glossary", "glossary", "Glossary.listAutoTextEntries");
  return listBuildingBlocks(glossary, "autoText");
}

/**
 * Extract Quick Parts from a glossary document.
 */
export function getQuickParts(glossary: GlossaryDocument): readonly BuildingBlock[] {
  checkHandle(glossary, "glossary", "glossary", "Glossary.listQuickParts");
  return listBuildingBlocks(glossary, "quickParts");
}
