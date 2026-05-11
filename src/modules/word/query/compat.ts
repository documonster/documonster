/**
 * Compatibility Mode API
 *
 * Get and set the Word compatibility mode of a document.
 */

import { type Mutable } from "../core/internal-utils";
import type { DocxDocument, DocumentSettings, CompatSetting } from "../types";

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
 * Looks at the `compatSetting` named "compatibilityMode" in document settings.
 * Returns 15 (Word 2013+) by default if not found.
 *
 * @param doc - The document to inspect.
 * @returns The compatibility mode version number.
 */
export function getCompatibilityMode(doc: DocxDocument): CompatibilityMode {
  if (!doc.settings?.compatSettings) {
    return 15;
  }
  const modeSetting = doc.settings.compatSettings.find(s => s.name === "compatibilityMode");
  if (!modeSetting?.val) {
    return 15;
  }
  const n = parseInt(modeSetting.val, 10);
  if (n === 11 || n === 12 || n === 14 || n === 15) {
    return n;
  }
  return 15;
}

/**
 * Set the compatibility mode of a document (mutates settings in place).
 *
 * @param doc - The document to modify (mutated in place).
 * @param mode - The target compatibility mode (11=Word 2003, 12=Word 2007, 14=Word 2010, 15=Word 2013+).
 */
export function setCompatibilityMode(doc: DocxDocument, mode: CompatibilityMode): void {
  const settings: Mutable<DocumentSettings> = doc.settings ? { ...doc.settings } : {};
  const compatSettings: CompatSetting[] = settings.compatSettings
    ? [...settings.compatSettings]
    : [];
  const idx = compatSettings.findIndex(s => s.name === "compatibilityMode");
  const entry: CompatSetting = {
    name: "compatibilityMode",
    uri: "http://schemas.microsoft.com/office/word",
    val: String(mode)
  };
  if (idx >= 0) {
    compatSettings[idx] = entry;
  } else {
    compatSettings.push(entry);
  }
  settings.compatSettings = compatSettings;
  (doc as Mutable<DocxDocument>).settings = settings;
}
