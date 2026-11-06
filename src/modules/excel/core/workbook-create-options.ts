/**
 * What `Workbook.create` takes, with its schema.
 *
 * Kept apart from `workbook-io-types`, which `Workbook.create` must not import: that module reaches the
 * package readers through `await import()`, and a module the workbook core imports statically pulls those
 * dynamic imports into every graph rooted at the core — which a single-file IIFE bundle (`Documonster.Word`,
 * whose Excel bridge starts at the core) cannot hold.
 */
import type { SyntaxProbe } from "@excel/core/defined-names";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

/** What `Workbook.create` takes — and what a read that returns a new workbook accepts beside its own options. */
export interface WorkbookCreateOptions {
  /** Classifies defined-name text as formula or opaque when reading; see {@link SyntaxProbe}. */
  readonly formulaSyntaxProbe?: SyntaxProbe;
}

export const createOptionsSchema = s.object(
  {
    // Typed as the interface declares it; `null` is accepted at run time, as it always was.
    formulaSyntaxProbe: s
      .custom<SyntaxProbe>(v => v === null || typeof v === "function", "a function")
      .optional()
  },
  { readonly: true }
);
const _createOptionsMatch: SchemaMatches<typeof createOptionsSchema, WorkbookCreateOptions> = true;
void _createOptionsMatch;
