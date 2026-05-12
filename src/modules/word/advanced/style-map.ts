/**
 * DOCX Module - Style Mapping DSL
 *
 * A domain-specific language for customizing how DOCX elements are
 * mapped to HTML output during conversion. Inspired by mammoth.js's
 * style mapping approach but designed for this project's architecture.
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

// =============================================================================
// Default Mappings
// =============================================================================

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
  const rules: StyleMappingRule[] = [];

  // Include defaults if requested
  if (options?.includeDefaults !== false && options?.base) {
    rules.push(...options.base.rules);
  } else if (options?.includeDefaults !== false && !options?.base) {
    rules.push(...DEFAULT_STYLE_MAP.rules);
  }

  const lines = dsl
    .split("\n")
    .map(l => l.trim())
    .filter(l => l && !l.startsWith("//") && !l.startsWith("#"));

  for (const line of lines) {
    const rule = parseRule(line);
    if (rule) {
      rules.push(rule);
    }
  }

  // Sort by priority (highest first)
  rules.sort((a, b) => (b.priority ?? 0) - (a.priority ?? 0));

  return { rules };
}

/**
 * Create a style map programmatically (without DSL parsing).
 */
export function createStyleMap(rules: readonly StyleMappingRule[]): StyleMap {
  const sorted = [...rules].sort((a, b) => (b.priority ?? 0) - (a.priority ?? 0));
  return { rules: sorted };
}

/**
 * Merge multiple style maps together.
 * Later maps take precedence (higher priority) over earlier ones.
 */
export function mergeStyleMaps(...maps: readonly StyleMap[]): StyleMap {
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

    // Parse attributes [key=value]
    let attributes: Record<string, string> | undefined;
    if (tagMatch[3]) {
      attributes = {};
      const attrRegex = /\[([^=]+)=([^\]]+)\]/g;
      let attrMatch: RegExpExecArray | null;
      while ((attrMatch = attrRegex.exec(tagMatch[3])) !== null) {
        attributes[attrMatch[1]!] = attrMatch[2]!.replace(/^['"]|['"]$/g, "");
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
