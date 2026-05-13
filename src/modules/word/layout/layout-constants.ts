/**
 * Shared layout constants for the layout / pagination / SVG render modules.
 *
 * These defaults are used when a `SectionProperties` does not specify a
 * page size or margin. Keeping them in a single file avoids drift between
 * `layout.ts`, `layout-full.ts`, and `render-page.ts`.
 *
 * The values match Microsoft Word's defaults for US Letter paper:
 *   - Page size: 8.5 in × 11 in (12240 × 15840 twips)
 *   - Margins:   1 in on all sides (1440 twips)
 */

/** Default page width in twips (US Letter, 8.5 in). */
export const DEFAULT_PAGE_WIDTH_TWIPS = 12240;

/** Default page height in twips (US Letter, 11 in). */
export const DEFAULT_PAGE_HEIGHT_TWIPS = 15840;

/** Default page margin in twips (1 in on each side). */
export const DEFAULT_PAGE_MARGIN_TWIPS = 1440;
