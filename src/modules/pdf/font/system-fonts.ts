/**
 * System font discovery for the PDF module — the policy, not the search.
 *
 * The search is `@utils/font-discovery`, at Layer 0, because the glyph rasteriser
 * in `draw` needs the same answer and `draw` cannot import `pdf`. See that
 * module's header for why it moved and what the split had been costing.
 *
 * ## Why this file still exists
 *
 * Discovery has to be told which code points genuinely require an embedded face,
 * and for PDF that is not all of them: a Type3 glyph draws every arrow, box
 * character, dingbat and enclosed numeral, so those must **not** disqualify an
 * otherwise-correct CJK font. Getting this wrong is not hypothetical — two
 * task-list checkboxes once disqualified every Chinese face on the machine and
 * set 508 Chinese characters in a Japanese hand.
 *
 * Binding `requiresEmbeddedFace` here means every PDF pipeline asks the same
 * question, which is the property the original module documented and is worth
 * keeping now that the search itself is shared with a caller whose answer differs.
 *
 * @module
 */

import { requiresEmbeddedFace } from "@pdf/font/type3-repertoire";
import type { CjkLanguage } from "@utils/cjk";
import { findSystemFontForCodePoints as findFace } from "@utils/font-discovery";
import type { TtfFont } from "@utils/font-ttf";

export type { SystemFontCandidate } from "@utils/font-discovery";
export {
  FAMILIES_BY_LANGUAGE,
  FONT_FILES_BY_LANGUAGE,
  PREFERRED_FONTS,
  discoverSystemFont,
  discoverSystemFontCandidates,
  preferredFontFiles,
  resetFontDiscoveryCache,
  _isReadBufferHeldForTest,
  _setCandidatesForTest
} from "@utils/font-discovery";

/**
 * Find an installed face that can draw `codePoints`, for PDF embedding.
 *
 * Identical to {@link findFace} except that symbols Type3 can draw are not
 * treated as requirements. See `@utils/font-discovery` for the selection rules
 * — language first, coverage second, partial coverage over none.
 *
 * The positional signature is kept so that nothing else in `pdf` changes: the shared
 * search takes an options object because it also accepts a target weight and slant,
 * which the PDF embedder does not yet ask for.
 */
export function findSystemFontForCodePoints(
  codePoints: ReadonlySet<number>,
  preferredFamilies: readonly string[] = [],
  language?: CjkLanguage
): TtfFont | null {
  return findFace(codePoints, {
    families: preferredFamilies,
    ...(language === undefined ? {} : { language }),
    requiresFace: requiresEmbeddedFace
  });
}
