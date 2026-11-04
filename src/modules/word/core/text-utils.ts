/**
 * DOCX Module - Text Utilities & Type Guards
 *
 * Single source of truth for type guards and text extraction helpers
 * used across the word module. Eliminates duplicated implementations
 * of `isRun`, `extractParagraphText`, `extractMathText`, etc.
 */

import type {
  DeletedRun,
  MovedFromRun,
  Paragraph,
  ParagraphChild,
  Run,
  MathContent,
  Hyperlink,
  HyperlinkNestedChild,
  InsertedRun,
  MovedToRun,
  RangeMarker,
  SymbolContent
} from "@word/types";

// =============================================================================
// Type Guards — Paragraph Children
// =============================================================================

/**
 * Check if a paragraph child is a Run.
 *
 * A Run is the only ParagraphChild without a `type` discriminator field.
 * All other paragraph children (Hyperlink, InsertedRun, DeletedRun, MovedFromRun,
 * MovedToRun, BookmarkStart/End, CommentRangeStart/End) have an explicit `type`
 * string. Therefore, a child is a Run iff it has a `content` array AND no `type`
 * field — this is more robust than checking `content` alone (since some typed
 * children also have `content`-like properties).
 */
export function isRun(child: ParagraphChild): child is Run {
  if ("type" in child) {
    return false;
  }
  return "content" in child && Array.isArray((child as Run).content);
}

/** Check if a paragraph child is a Hyperlink. */
export function isHyperlink(child: ParagraphChild): child is Hyperlink {
  return "type" in child && (child as { type: string }).type === "hyperlink";
}

/**
 * The content of a hyperlink in document order: its runs with its
 * {@link Hyperlink.nestedChildren} among them. Range markers keep riding on the
 * items they rode on; {@link Hyperlink.rangeMarkersAfter} is not included.
 */
export function hyperlinkContent(link: Hyperlink): ParagraphChild[] {
  const nested = link.nestedChildren;
  if (!nested || nested.length === 0) {
    return [...link.children];
  }
  const out: ParagraphChild[] = [];
  let next = 0;
  link.children.forEach((run, i) => {
    while (next < nested.length && nested[next]!.before <= i) {
      out.push(nested[next++]!.child);
    }
    out.push(run);
  });
  while (next < nested.length) {
    out.push(nested[next++]!.child);
  }
  return out;
}

/**
 * `link` with its content replaced by `items`, in document order (the inverse
 * of {@link hyperlinkContent}): runs become its children and everything else
 * nested children among them. A range marker among the items rides on the
 * item after it, or goes in front of `rangeMarkersAfter` at the end.
 */
export function withHyperlinkContent(
  link: Hyperlink,
  items: readonly (ParagraphChild | RangeMarker)[]
): Hyperlink {
  const children: Run[] = [];
  const nested: HyperlinkNestedChild[] = [];
  let pending: RangeMarker[] = [];
  const ride = <T extends ParagraphChild>(item: T): T => {
    if (pending.length === 0) {
      return item;
    }
    const before = [...pending, ...(item.rangeMarkersBefore ?? [])];
    pending = [];
    return { ...item, rangeMarkersBefore: before };
  };
  for (const item of items) {
    if (isRangeMarker(item)) {
      pending.push(item);
    } else if (isRun(item)) {
      children.push(ride(item));
    } else {
      nested.push({ before: children.length, child: ride(item as Exclude<ParagraphChild, Run>) });
    }
  }
  const { nestedChildren: _nested, rangeMarkersAfter, ...rest } = link;
  const after = [...pending, ...(rangeMarkersAfter ?? [])];
  return {
    ...rest,
    children,
    ...(nested.length > 0 ? { nestedChildren: nested } : {}),
    ...(after.length > 0 ? { rangeMarkersAfter: after } : {})
  };
}

/** The `type` of every {@link RangeMarker} kind. */
export const RANGE_MARKER_TYPES: readonly RangeMarker["type"][] = [
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
  "customXmlInsRangeStart",
  "customXmlInsRangeEnd",
  "customXmlDelRangeStart",
  "customXmlDelRangeEnd",
  "customXmlMoveFromRangeStart",
  "customXmlMoveFromRangeEnd",
  "customXmlMoveToRangeStart",
  "customXmlMoveToRangeEnd",
  "permStart",
  "permEnd",
  "bidiStart",
  "bidiEnd",
  "smartTagStart",
  "smartTagEnd"
];

const RANGE_MARKER_TYPE_SET: ReadonlySet<string> = new Set(RANGE_MARKER_TYPES);

/** Check if a model node (paragraph child or body entry) is a {@link RangeMarker}. */
export function isRangeMarker(node: object): node is RangeMarker {
  return "type" in node && RANGE_MARKER_TYPE_SET.has((node as { type: string }).type);
}

/** A {@link RangeMarker} that stands for a wrapper element around paragraph content. */
export type SpanMarker = Extract<RangeMarker, { type: SpanMarkerType }>;
type SpanMarkerType = "bidiStart" | "bidiEnd" | "smartTagStart" | "smartTagEnd";

const SPAN_MARKER_TYPES: ReadonlySet<string> = new Set<SpanMarkerType>([
  "bidiStart",
  "bidiEnd",
  "smartTagStart",
  "smartTagEnd"
]);

/**
 * Whether a range marker is a span marker (`bidiStart`/`bidiEnd`,
 * `smartTagStart`/`smartTagEnd`): it pairs by nesting rather than by id, and
 * means nothing outside the paragraph content it was in.
 */
export function isSpanMarker(marker: RangeMarker): marker is SpanMarker {
  return SPAN_MARKER_TYPES.has(marker.type);
}

// =============================================================================
// Text Extraction — Runs
// =============================================================================

/**
 * Extract plain text from a single run's content array.
 */
export function extractRunText(run: Run): string {
  let text = "";
  for (const c of run.content) {
    if (c.type === "text") {
      text += c.text;
    } else if (c.type === "tab") {
      text += "\t";
    } else if (c.type === "break") {
      text += "\n";
    } else if (c.type === "noBreakHyphen") {
      text += "-";
    } else if (c.type === "softHyphen") {
      text += "\u00AD";
    } else if (c.type === "field" && c.cachedValue) {
      text += c.cachedValue;
    }
  }
  return text;
}

/**
 * The character a `w:sym` stands for. `w:char` is a hexadecimal code point
 * (often in the Private Use Area of a symbol font); an unparseable code yields
 * an empty string rather than the raw hex digits.
 */
export function symbolText(symbol: SymbolContent): string {
  const code = /^[0-9a-f]+$/i.test(symbol.char) ? parseInt(symbol.char, 16) : NaN;
  return Number.isInteger(code) && code <= 0x10ffff ? String.fromCodePoint(code) : "";
}

// =============================================================================
// Text Extraction — Paragraphs
// =============================================================================

/**
 * Extract concatenated plain text from a paragraph's children (runs + hyperlinks).
 * This recursively handles hyperlinks and tracked changes. What lies inside
 * a field's instruction — between the begin and separate characters of a
 * field held in pieces (`FieldContent.fieldChar`), such as a field nested in
 * an `IF` — is not shown, so not included.
 */
export function extractParagraphText(para: Paragraph): string {
  let pieces = false;
  for (const child of para.children) {
    if (holdsPiece(child)) {
      pieces = true;
      break;
    }
  }
  let text = "";
  if (!pieces) {
    for (const child of para.children) {
      text += extractChildText(child);
    }
    return text;
  }
  // Each open field: whether its separate character has been met.
  const open: boolean[] = [];
  const visit = (run: Run, visible: boolean): void => {
    for (const c of run.content) {
      if (c.type === "field" && c.fieldChar) {
        if (c.fieldChar === "begin") {
          open.push(false);
        } else if (c.fieldChar === "separate") {
          if (open.length > 0) {
            open[open.length - 1] = true;
          }
        } else {
          open.pop();
        }
      } else if (visible && !open.includes(false)) {
        text += extractRunText({ content: [c] });
      }
    }
  };
  const walk = (children: readonly ParagraphChild[]): void => {
    for (const child of children) {
      if (isRun(child)) {
        visit(child, true);
        continue;
      }
      const type = (child as { type: string }).type;
      if (type === "hyperlink") {
        walk(hyperlinkContent(child as Hyperlink));
      } else if (type === "insertedRun" || type === "movedToRun") {
        visit((child as InsertedRun | MovedToRun).run, true);
      } else if (type === "deletedRun" || type === "movedFromRun") {
        // Not visible, but its field characters still open and close fields.
        visit((child as DeletedRun | MovedFromRun).run, false);
      }
    }
  };
  walk(para.children);
  return text;
}

function runHoldsPiece(run: Run): boolean {
  for (const c of run.content) {
    if (c.type === "field" && c.fieldChar) {
      return true;
    }
  }
  return false;
}

function holdsPiece(child: ParagraphChild): boolean {
  if (isRun(child)) {
    return runHoldsPiece(child);
  }
  const type = (child as { type: string }).type;
  if (type === "hyperlink") {
    const link = child as Hyperlink;
    return (
      link.children.some(runHoldsPiece) ||
      (link.nestedChildren?.some(n => holdsPiece(n.child)) ?? false)
    );
  }
  const run = (child as { run?: Run }).run;
  return run !== undefined && runHoldsPiece(run);
}

/**
 * Extract text from a single paragraph child (run, hyperlink, or tracked change).
 */
export function extractChildText(child: ParagraphChild): string {
  if (isRun(child)) {
    return extractRunText(child);
  }
  if ("type" in child) {
    const typed = child as { type: string };
    switch (typed.type) {
      case "hyperlink": {
        let text = "";
        for (const c of hyperlinkContent(child as Hyperlink)) {
          text += extractChildText(c);
        }
        return text;
      }
      case "insertedRun":
        return extractRunText((child as InsertedRun).run);
      case "deletedRun":
        return ""; // Deleted content is not visible
      case "movedFromRun":
        return ""; // Moved-from is not visible at source
      case "movedToRun":
        return extractRunText((child as MovedToRun).run);
      default:
        return "";
    }
  }
  return "";
}

// =============================================================================
// Text Extraction — Math
// =============================================================================

/**
 * Extract plain text representation from math content elements.
 */
export function extractMathText(content: readonly MathContent[]): string {
  let text = "";
  for (const item of content) {
    switch (item.type) {
      case "mathRun":
        text += item.text;
        break;
      case "mathFraction":
        text += `(${extractMathText(item.numerator)})/(${extractMathText(item.denominator)})`;
        break;
      case "mathRadical":
        if (item.hideDegree) {
          text += `√(${extractMathText(item.content)})`;
        } else if (item.degree) {
          text += `${extractMathText(item.degree)}√(${extractMathText(item.content)})`;
        } else {
          text += `√(${extractMathText(item.content)})`;
        }
        break;
      case "mathSuperScript":
        text += `${extractMathText(item.base)}^(${extractMathText(item.superScript)})`;
        break;
      case "mathSubScript":
        text += `${extractMathText(item.base)}_(${extractMathText(item.subScript)})`;
        break;
      case "mathSubSuperScript":
        text += `${extractMathText(item.base)}_(${extractMathText(item.subScript)})^(${extractMathText(item.superScript)})`;
        break;
      case "mathPreSubSuperScript":
        text += `_(${extractMathText(item.preSubScript)})^(${extractMathText(item.preSuperScript)})${extractMathText(item.base)}`;
        break;
      case "mathDelimiter":
        text += item.beginChar ?? "(";
        text += item.content.map(c => extractMathText(c)).join(item.separatorChar ?? ",");
        text += item.endChar ?? ")";
        break;
      case "mathNary":
        text += item.char ?? "∑";
        if (item.sub) {
          text += `_(${extractMathText(item.sub)})`;
        }
        if (item.sup) {
          text += `^(${extractMathText(item.sup)})`;
        }
        text += extractMathText(item.content);
        break;
      case "mathFunction":
        text += extractMathText(item.name);
        text += extractMathText(item.content);
        break;
      case "mathLimit":
        text += extractMathText(item.base);
        text += item.limitType === "upper" ? "^" : "_";
        text += `(${extractMathText(item.limit)})`;
        break;
      case "mathMatrix":
        text += "[";
        text += item.rows.map(row => row.map(cell => extractMathText(cell)).join(",")).join(";");
        text += "]";
        break;
      case "mathAccent":
        text += extractMathText(item.content);
        if (item.char) {
          text += item.char;
        }
        break;
      case "mathBar":
        text += extractMathText(item.content);
        break;
      case "mathBox":
        text += extractMathText(item.content);
        break;
      case "mathEquationArray":
        text += item.rows.map(row => extractMathText(row)).join("\n");
        break;
      case "mathGroupChar":
        text += extractMathText(item.base);
        break;
      case "mathPhantom":
        text += extractMathText(item.content);
        break;
      case "mathBorderBox":
        text += extractMathText(item.content);
        break;
    }
  }
  return text;
}
