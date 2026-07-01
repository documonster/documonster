/**
 * Compatibility Mode API
 *
 * Get and set the Word compatibility mode of a document.
 */

import type { Mutable } from "@word/core/internal-utils";
import type { DocxDocument, DocumentSettings } from "@word/types";

// =============================================================================
// Types
// =============================================================================

/** Word compatibility mode versions. */
export type CompatibilityMode = 11 | 12 | 14 | 15;

// =============================================================================
// Public API
// =============================================================================

/**
 * Get the compatibility mode of a document.
 *
 * The mode is stored in `settings.compatibilityMode` (the canonical scalar
 * field populated by the reader). For backward compatibility we also honour an
 * explicit `compatibilityMode` entry in `settings.compatSettings` if present,
 * since the writer accepts that advanced-override path.
 *
 * Returns 15 (Word 2013+) by default if nothing is stored.
 *
 * @param doc - The document to inspect.
 * @returns The compatibility mode version number.
 */
export function getCompatibilityMode(doc: DocxDocument): CompatibilityMode {
  const settings = doc.settings;
  if (!settings) {
    return 15;
  }

  // Prefer an explicit override entry in compatSettings (advanced path) so a
  // hand-authored value wins, then fall back to the canonical scalar field.
  const overrideEntry = settings.compatSettings?.find(s => s.name === "compatibilityMode");
  const raw = overrideEntry?.val ?? settings.compatibilityMode;
  if (raw === undefined) {
    return 15;
  }

  const n = typeof raw === "number" ? raw : parseInt(raw, 10);
  if (n === 11 || n === 12 || n === 14 || n === 15) {
    return n;
  }
  return 15;
}

/**
 * Set the compatibility mode of a document (mutates settings in place).
 *
 * Writes the canonical `settings.compatibilityMode` scalar field and removes
 * any stale `compatibilityMode` override entry from `settings.compatSettings`
 * so the two sources never disagree.
 *
 * @param doc - The document to modify (mutated in place).
 * @param mode - The target compatibility mode (11=Word 2003, 12=Word 2007, 14=Word 2010, 15=Word 2013+).
 */
export function setCompatibilityMode(doc: DocxDocument, mode: CompatibilityMode): void {
  const settings: Mutable<DocumentSettings> = doc.settings ? { ...doc.settings } : {};
  settings.compatibilityMode = mode;

  // Drop any stale override entry so getCompatibilityMode/writer don't read a
  // conflicting value from the array. The scalar field is now authoritative.
  if (settings.compatSettings) {
    const filtered = settings.compatSettings.filter(s => s.name !== "compatibilityMode");
    if (filtered.length > 0) {
      settings.compatSettings = filtered;
    } else {
      delete settings.compatSettings;
    }
  }

  (doc as Mutable<DocxDocument>).settings = settings;
}
