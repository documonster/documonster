/**
 * DOCX Module - Style Mapping DSL
 *
 * A domain-specific language for customizing how DOCX elements are
 * mapped to HTML output during conversion. A style-mapping approach
 * designed for this project's architecture.
 *
 * The DSL allows users to define rules that control the HTML output
 * for specific Word styles, without modifying the core renderer.
 *
 * Syntax examples:
 * ```
 * p[style-name='Heading 1'] => h1.custom-heading
 * p[style-name='Quote'] => blockquote.pull-quote
 * r[style-name='Emphasis'] => em.highlight
 * p[style-name='Code'] => pre > code
 * table[style-name='GridTable'] => table.data-table
 * ```
 */

import type { SchemaMatches } from "@utils/schema";
import {
  checkArray,
  checkBoolean,
  checkFinite,
  checkHandle,
  checkOneOf,
  checkRecord,
  checkString,
  invalidArgument
} from "@word/core/argument-checks";
import { checkSchema, model, w } from "@word/core/option-schema";

// =============================================================================
// Types
// =============================================================================

/** Target element type in the mapping rule. */
export type MappingSourceType = "p" | "r" | "table" | "image" | "b" | "i" | "u" | "strike";

/** A condition for matching source elements. */
export interface MappingCondition {
  /** Attribute to match on. */
  readonly attribute:
    | "style-name"
    | "style-id"
    | "outline-level"
    | "numbering-level"
    | "is-bold"
    | "is-italic";
  /** Value to match (exact or regex). */
  readonly value: string;
  /** Use regex matching. */
  readonly regex?: boolean;
}

/** Target HTML element specification. */
export interface MappingTarget {
  /** HTML tag name. */
  readonly tagName: string;
  /** CSS class(es) to add. */
  readonly className?: string;
  /** Additional HTML attributes. */
  readonly attributes?: Readonly<Record<string, string>>;
  /** Nested element (for `pre > code` style mappings). */
  readonly child?: MappingTarget;
  /** Whether to preserve separator. */
  readonly separator?: string;
  /** Whether this is a "fresh" mapping (wrap each paragraph separately). */
  readonly fresh?: boolean;
}

/** A single style mapping rule. */
export interface StyleMappingRule {
  /** Source element type. */
  readonly source: MappingSourceType;
  /** Conditions that must all be satisfied. */
  readonly conditions: readonly MappingCondition[];
  /** Target HTML output. */
  readonly target: MappingTarget;
  /** Priority (higher = checked first). Default: 0. */
  readonly priority?: number;
}

/** A compiled set of style mapping rules. */
export interface StyleMap {
  /** All rules, sorted by priority (highest first). */
  readonly rules: readonly StyleMappingRule[];
}

/** Options for the style mapping parser. */
export interface StyleMapOptions {
  /** Base rules to include (before user rules). */
  readonly base?: StyleMap;
  /** Whether to include default mappings (Heading → h1-h6, etc.). */
  readonly includeDefaults?: boolean;
}

/** `base` is a style map, which `checkStyleMap` checks. */
const styleMapOptionsSchema = w.record({
  base: model<StyleMap>(),
  includeDefaults: w.boolean().optional()
});
const _styleMapOptionsMatch: SchemaMatches<typeof styleMapOptionsSchema, StyleMapOptions> = true;
void _styleMapOptionsMatch;

// =============================================================================
// Default Mappings
// =============================================================================

const SOURCE_TYPES: ReadonlySet<string> = new Set<MappingSourceType>([
  "p",
  "r",
  "table",
  "image",
  "b",
  "i",
  "u",
  "strike"
]);
const CONDITION_ATTRIBUTES: ReadonlySet<string> = new Set<MappingCondition["attribute"]>([
  "style-name",
  "style-id",
  "outline-level",
  "numbering-level",
  "is-bold",
  "is-italic"
]);

function checkStyleMap(map: unknown, field: string, where: string): void {
  checkRecord(map, field, where);
  checkRules((map as StyleMap).rules, `${field}.rules`, where);
}

function checkTarget(target: unknown, field: string, where: string): void {
  checkRecord(target, field, where);
  const t = target as MappingTarget;
  if (typeof t.tagName !== "string" || !/^[A-Za-z][A-Za-z0-9-]*$/.test(t.tagName)) {
    invalidArgument(where, `${field}.tagName`, "an HTML tag name", t.tagName);
  }
  checkString(t.className, `${field}.className`, where, { optional: true });
  checkString(t.separator, `${field}.separator`, where, { optional: true });
  checkBoolean(t.fresh, `${field}.fresh`, where, { optional: true });
  if (t.child !== undefined) {
    checkTarget(t.child, `${field}.child`, where);
  }
}

/** Rules as `Styles.create` takes them; a regex condition must compile. */
function checkRules(rules: unknown, field: string, where: string): void {
  checkArray(rules, field, where);
  (rules as unknown[]).forEach((rule, i) => {
    const at = `${field}[${i}]`;
    checkRecord(rule, at, where);
    const r = rule as StyleMappingRule;
    checkOneOf(r.source, SOURCE_TYPES, `${at}.source`, where);
    checkArray(r.conditions, `${at}.conditions`, where);
    r.conditions.forEach((cond, j) => {
      const c = `${at}.conditions[${j}]`;
      checkRecord(cond, c, where);
      checkOneOf(cond.attribute, CONDITION_ATTRIBUTES, `${c}.attribute`, where);
      checkString(cond.value, `${c}.value`, where);
      checkBoolean(cond.regex, `${c}.regex`, where, { optional: true });
      if (cond.regex) {
        try {
          new RegExp(cond.value);
        } catch {
          invalidArgument(where, `${c}.value`, "a valid regular expression", cond.value);
        }
      }
    });
    checkTarget(r.target, `${at}.target`, where);
    checkFinite(r.priority, `${at}.priority`, where, { optional: true });
  });
}

/** Default style mappings (standard Word styles → semantic HTML). */
export const DEFAULT_STYLE_MAP: StyleMap = {
  rules: [
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Heading 1" }],
      target: { tagName: "h1" },
      priority: 10
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Heading 2" }],
      target: { tagName: "h2" },
      priority: 10
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Heading 3" }],
      target: { tagName: "h3" },
      priority: 10
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Heading 4" }],
      target: { tagName: "h4" },
      priority: 10
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Heading 5" }],
      target: { tagName: "h5" },
      priority: 10
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Heading 6" }],
      target: { tagName: "h6" },
      priority: 10
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Title" }],
      target: { tagName: "h1", className: "document-title" },
      priority: 10
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Subtitle" }],
      target: { tagName: "h2", className: "document-subtitle" },
      priority: 10
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Quote" }],
      target: { tagName: "blockquote" },
      priority: 5
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "Intense Quote" }],
      target: { tagName: "blockquote", className: "intense" },
      priority: 5
    },
    {
      source: "p",
      conditions: [{ attribute: "style-name", value: "List Paragraph" }],
      target: { tagName: "p", className: "list-paragraph" },
      priority: 3
    },
    {
      source: "r",
      conditions: [{ attribute: "style-name", value: "Strong" }],
      target: { tagName: "strong" },
      priority: 5
    },
    {
      source: "r",
      conditions: [{ attribute: "style-name", value: "Emphasis" }],
      target: { tagName: "em" },
      priority: 5
    },
    {
      source: "r",
      conditions: [{ attribute: "style-name", value: "Hyperlink" }],
      target: { tagName: "a" },
      priority: 5
    },
    {
      source: "r",
      conditions: [{ attribute: "style-name", value: "Code" }],
      target: { tagName: "code" },
      priority: 5
    },
    { source: "b", conditions: [], target: { tagName: "strong" }, priority: 1 },
    { source: "i", conditions: [], target: { tagName: "em" }, priority: 1 },
    { source: "u", conditions: [], target: { tagName: "u" }, priority: 1 },
    { source: "strike", conditions: [], target: { tagName: "s" }, priority: 1 }
  ]
};

// =============================================================================
// DSL Parser
// =============================================================================

/**
 * Parse a style mapping DSL string into a StyleMap.
 *
 * @param dsl - The DSL string (one rule per line).
 * @param options - Parser options.
 * @returns A compiled StyleMap.
 *
 * @example
 * ```ts
 * const map = parseStyleMap(`
 *   p[style-name='Heading 1'] => h1.title
 *   p[style-name='Code Block'] => pre > code
 *   r[style-name='Emphasis'] => em.custom-em
 * `);
 * ```
 */
export function parseStyleMap(dsl: string, options?: StyleMapOptions): StyleMap {
  const where = "Styles.parse";
  checkString(dsl, "dsl", where);
  checkSchema(styleMapOptionsSchema.optional(), options, where);
  if (options?.base !== undefined) {
    checkStyleMap(options.base, "options.base", where);
  }
  // User-defined rules from the DSL.
  const userRules: StyleMappingRule[] = [];

  const lines = dsl
    .split("\n")
    .map(l => l.trim())
    .filter(l => l && !l.startsWith("//") && !l.startsWith("#"));

  for (const line of lines) {
    const rule = parseRule(line);
    if (rule) {
      userRules.push(rule);
    }
  }

  // Default / base rules requested via `includeDefaults`.
  const defaultRules: StyleMappingRule[] = [];
  if (options?.includeDefaults !== false && options?.base) {
    defaultRules.push(...options.base.rules);
  } else if (options?.includeDefaults !== false && !options?.base) {
    defaultRules.push(...DEFAULT_STYLE_MAP.rules);
  }

  // An explicit DSL rule should always win over a default rule for the same
  // element — that is the whole point of providing one. Default rules,
  // however, carry their own priorities (e.g. "Heading 1" => h1 has priority
  // 10) that can exceed the fixed priority `parseRule` assigns to user rules.
  // To guarantee user intent wins while preserving the relative priority
  // ordering *within* each group, lift every user rule above the highest
  // default priority. When there are no defaults this offset is 0 and user
  // priorities are untouched.
  const maxDefaultPriority = defaultRules.reduce((m, r) => Math.max(m, r.priority ?? 0), 0);
  const userOffset = defaultRules.length > 0 ? maxDefaultPriority + 1 : 0;
  const liftedUserRules =
    userOffset === 0
      ? userRules
      : userRules.map(r => ({ ...r, priority: (r.priority ?? 0) + userOffset }));

  // User rules come first so that, after a stable sort, an explicit DSL rule
  // also wins over any default rule that happens to share its (lifted)
  // priority. The sort below only reorders by priority; equal priorities
  // preserve this user-before-default ordering.
  const rules = [...liftedUserRules, ...defaultRules];

  // Sort by priority (highest first). Array.prototype.sort is stable, so
  // rules of equal priority keep their relative order (user rules first).
  rules.sort((a, b) => (b.priority ?? 0) - (a.priority ?? 0));

  return { rules };
}

/**
 * Create a style map programmatically (without DSL parsing).
 */
export function createStyleMap(rules: readonly StyleMappingRule[]): StyleMap {
  checkRules(rules, "rules", "Styles.create");
  const sorted = [...rules].sort((a, b) => (b.priority ?? 0) - (a.priority ?? 0));
  return { rules: sorted };
}

/**
 * Merge multiple style maps together.
 * Later maps take precedence (higher priority) over earlier ones.
 */
export function mergeStyleMaps(...maps: readonly StyleMap[]): StyleMap {
  maps.forEach((map, i) => checkStyleMap(map, `maps[${i}]`, "Styles.merge"));
  const allRules: StyleMappingRule[] = [];
  for (let i = 0; i < maps.length; i++) {
    // Add base priority offset based on position
    for (const rule of maps[i]!.rules) {
      allRules.push({
        ...rule,
        priority: (rule.priority ?? 0) + i * 100
      });
    }
  }
  allRules.sort((a, b) => (b.priority ?? 0) - (a.priority ?? 0));
  return { rules: allRules };
}

/**
 * Match a source element against the style map and return the target.
 *
 * @param map - The style map to match against.
 * @param source - Source element type.
 * @param attributes - Attributes of the source element.
 * @returns The matching target, or undefined if no match.
 */
export function matchStyleMap(
  map: StyleMap,
  source: MappingSourceType,
  attributes: Readonly<Record<string, string>>
): MappingTarget | undefined {
  checkHandle(map, "styleMap", "map", "Styles.match");
  checkRecord(attributes, "attributes", "Styles.match");
  for (const rule of map.rules) {
    if (rule.source !== source) {
      continue;
    }

    const allMatch = rule.conditions.every(cond => {
      const value = attributes[cond.attribute];
      if (value === undefined) {
        return false;
      }
      if (cond.regex) {
        // The pattern source comes from a style-map DSL that may be
        // attacker-controlled. We can't introspect a regex for
        // catastrophic-backtracking risk, but we can:
        //   1) reject the rule outright when the pattern is malformed
        //      (the previous code threw, which broke the whole map);
        //   2) cap the input string we feed to .test() so a quadratic
        //      pattern can't run forever even if one slips through.
        let re: RegExp;
        try {
          re = new RegExp(cond.value);
        } catch {
          return false;
        }
        const STYLE_MAP_REGEX_INPUT_CAP = 4096;
        const limited =
          value.length > STYLE_MAP_REGEX_INPUT_CAP
            ? value.slice(0, STYLE_MAP_REGEX_INPUT_CAP)
            : value;
        return re.test(limited);
      }
      return value === cond.value;
    });

    if (allMatch) {
      return rule.target;
    }
  }

  return undefined;
}

// =============================================================================
// Internal: Rule Parser
// =============================================================================

function parseRule(line: string): StyleMappingRule | null {
  // Format: source[conditions] => target
  const arrowIdx = line.indexOf("=>");
  if (arrowIdx < 0) {
    return null;
  }

  const sourcePart = line.substring(0, arrowIdx).trim();
  const targetPart = line.substring(arrowIdx + 2).trim();

  // Parse source
  const sourceMatch = sourcePart.match(/^(p|r|table|image|b|i|u|strike)(.*)$/);
  if (!sourceMatch) {
    return null;
  }

  const sourceType = sourceMatch[1] as MappingSourceType;
  const conditionsStr = sourceMatch[2]!;

  // Parse conditions: [attr='value'][attr2='value2']
  const conditions: MappingCondition[] = [];
  const condRegex = /\[([^=]+)='([^']+)'\]/g;
  let condMatch: RegExpExecArray | null;
  while ((condMatch = condRegex.exec(conditionsStr)) !== null) {
    conditions.push({
      attribute: condMatch[1] as MappingCondition["attribute"],
      value: condMatch[2]!
    });
  }

  // Also support regex conditions: [attr~='pattern']
  const regexCondRegex = /\[([^~]+)~='([^']+)'\]/g;
  while ((condMatch = regexCondRegex.exec(conditionsStr)) !== null) {
    conditions.push({
      attribute: condMatch[1] as MappingCondition["attribute"],
      value: condMatch[2]!,
      regex: true
    });
  }

  // Parse target
  const target = parseTarget(targetPart);
  if (!target) {
    return null;
  }

  return {
    source: sourceType,
    conditions,
    target,
    priority: conditions.length > 0 ? 5 : 1
  };
}

function parseTarget(targetStr: string): MappingTarget | null {
  // Format: tagName.className > childTag.childClass
  const parts = targetStr.split(">").map(s => s.trim());

  const parseSingle = (s: string): MappingTarget | null => {
    // tag.class1.class2[attr=value]
    const tagMatch = s.match(/^([a-zA-Z][a-zA-Z0-9]*)((?:\.[a-zA-Z_-][a-zA-Z0-9_-]*)*)?(\[.*\])?$/);
    if (!tagMatch) {
      return null;
    }

    const tagName = tagMatch[1]!;
    const className = tagMatch[2] ? tagMatch[2].substring(1).replace(/\./g, " ") : undefined;

    // Parse attributes [key=value]. Implemented as a linear scan rather
    // than a global regex so that adversarial mapping strings cannot
    // trigger CodeQL's `js/polynomial-redos`. The regex form
    // `/\[([^=]+)=([^\]]+)\]/g` exhibits backtracking on inputs like
    // `[xxxxxxxxxxxxx`.
    let attributes: Record<string, string> | undefined;
    const attrSection = tagMatch[3];
    if (attrSection) {
      attributes = {};
      const len = attrSection.length;
      let i = 0;
      while (i < len) {
        if (attrSection.charCodeAt(i) !== 0x5b /* '[' */) {
          i++;
          continue;
        }
        const eq = attrSection.indexOf("=", i + 1);
        const close = attrSection.indexOf("]", i + 1);
        if (close < 0) {
          break;
        }
        if (eq < 0 || eq > close) {
          // No `=` inside this `[...]` — skip past it.
          i = close + 1;
          continue;
        }
        const key = attrSection.slice(i + 1, eq);
        const rawValue = attrSection.slice(eq + 1, close);
        // Strip a single leading and trailing quote (' or "). Two
        // independent linear strips replace the previous
        // `/^['"]|['"]$/g` regex.
        let v = rawValue;
        if (v.length >= 2) {
          const first = v.charCodeAt(0);
          if (first === 0x27 || first === 0x22) {
            v = v.slice(1);
          }
        }
        if (v.length >= 1) {
          const last = v.charCodeAt(v.length - 1);
          if (last === 0x27 || last === 0x22) {
            v = v.slice(0, -1);
          }
        }
        attributes[key] = v;
        i = close + 1;
      }
    }

    return { tagName, className: className || undefined, attributes };
  };

  if (parts.length === 0) {
    return null;
  }

  const first = parseSingle(parts[0]!);
  if (!first) {
    return null;
  }

  if (parts.length > 1) {
    const child = parseSingle(parts[1]!);
    if (child) {
      return { ...first, child };
    }
  }

  return first;
}
