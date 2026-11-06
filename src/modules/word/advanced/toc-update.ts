/**
 * Regenerating a table of contents as its `TOC` instruction asks.
 *
 * The entries are chosen by the switches: `\o` (outline levels) and `\t`
 * (styles with levels) pick headings, `\c` and `\a` the captions numbered by a
 * `SEQ` field of that label (a table of figures). Each entry is the target's
 * text and — unless `\n` omits it for its level — a `PAGEREF` field to a
 * `_Toc` bookmark on the target, after a right-aligned leader tab or the `\p`
 * separator; with `\h` it is a hyperlink to that bookmark. A target without a
 * `_Toc` bookmark gets one, as Word adds it, so a second update finds the
 * same names. `\w` and `\x` keep the tabs and line breaks of the target's
 * text, which are spaces otherwise; `\u` and `\z` change nothing here (the
 * outline level a paragraph carries itself always counts, and there is no web
 * view).
 *
 * Not built: entries from `TC` fields (`\f`, `\l`), the bookmark-limited range
 * (`\b`) and chapter-number prefixes (`\s`, `\d`).
 */

import { forEachField, mapBlockLists, resolveSplitFields } from "@word/core/field-view";
import type { Mutable } from "@word/core/internal-utils";
import { extractParagraphText } from "@word/core/text-utils";
import { tocSwitches } from "@word/core/toc-instruction";
import { walkDocument } from "@word/core/walker";
import type { LayoutResult } from "@word/layout/layout";
import { resolveHeadingLevel } from "@word/query/heading";
import { indexStyles } from "@word/query/style-resolve";
import type {
  BodyContent,
  BookmarkEnd,
  BookmarkStart,
  DocxDocument,
  Hyperlink,
  Paragraph,
  ParagraphChild,
  Run,
  TabStopLeader
} from "@word/types";

/** What a `TOC` instruction asks for. */
interface TocSpec {
  /** `\o`: the outline levels taken. */
  readonly levels: { readonly min: number; readonly max: number } | undefined;
  /** `\t`: style (name or id, lower-cased) → level. */
  readonly styles: ReadonlyMap<string, number>;
  /** `\c` / `\a`: the caption label, and whether the entry drops it with its number. */
  readonly caption: { readonly label: string; readonly bare: boolean } | undefined;
  readonly hyperlink: boolean;
  /** `\n`: the levels without a page number (`true`: every level). */
  readonly omitPages: { readonly min: number; readonly max: number } | true | undefined;
  /** `\p`: what separates an entry from its page number, instead of a tab. */
  readonly separator: string | undefined;
  readonly keepTabs: boolean;
  readonly keepBreaks: boolean;
}

function levelRange(value: string | undefined): { min: number; max: number } | undefined {
  const m = value ? /^\s*(\d+)\s*-\s*(\d+)\s*$/.exec(value) : null;
  return m ? { min: parseInt(m[1]!, 10), max: parseInt(m[2]!, 10) } : undefined;
}

function readTocSpec(instruction: string): TocSpec {
  let levels: TocSpec["levels"];
  const styles = new Map<string, number>();
  let caption: TocSpec["caption"];
  let hyperlink = false;
  let omitPages: TocSpec["omitPages"];
  let separator: string | undefined;
  let keepTabs = false;
  let keepBreaks = false;
  for (const sw of tocSwitches(instruction)) {
    switch (sw.letter) {
      case "o":
        levels = levelRange(sw.value) ?? { min: 1, max: 9 };
        break;
      case "t": {
        // "Style,1,Other,2" — or with ';' between the pairs, as some locales write it.
        const parts = (sw.value ?? "").split(/[;,]/).map(s => s.trim());
        for (let i = 0; i + 1 < parts.length; i += 2) {
          const level = parseInt(parts[i + 1]!, 10);
          if (parts[i] && Number.isFinite(level)) {
            styles.set(parts[i]!.toLowerCase(), level);
          }
        }
        break;
      }
      case "c":
      case "a":
        if (sw.value) {
          caption = { label: sw.value, bare: sw.letter === "a" };
        }
        break;
      case "h":
        hyperlink = true;
        break;
      case "n":
        omitPages = levelRange(sw.value) ?? true;
        break;
      case "p":
        separator = sw.value;
        break;
      case "w":
        keepTabs = true;
        break;
      case "x":
        keepBreaks = true;
        break;
      default:
        break;
    }
  }
  // A bare `TOC` lists the headings, as it always has here.
  if (!levels && styles.size === 0 && !caption) {
    levels = { min: 1, max: 9 };
  }
  return { levels, styles, caption, hyperlink, omitPages, separator, keepTabs, keepBreaks };
}

/** A paragraph a table of contents may list. */
interface Candidate {
  readonly paragraph: Paragraph;
  /** Index of the body block holding it. */
  readonly top: number;
  readonly heading: number | undefined;
  /** Its style's id and name, lower-cased. */
  readonly styleKeys: readonly string[];
  /** The `SEQ` identifiers its fields number, lower-cased. */
  readonly sequences: readonly string[];
}

/**
 * The paragraphs of the body a table of contents may list, in document
 * order — not those of a table of contents itself.
 */
function collectCandidates(doc: DocxDocument): Candidate[] {
  const styles = indexStyles(doc);
  const names = new Map<string, string>();
  for (const s of doc.styles ?? []) {
    if (s.name) {
      names.set(s.styleId, s.name.toLowerCase());
    }
  }
  const out: Candidate[] = [];
  const visitList = (blocks: readonly object[], top: number | undefined): void => {
    const skipped = new Set<number>();
    for (const f of resolveSplitFields(blocks)?.fields ?? []) {
      if (/^\s*TOC\b/i.test(f.instruction)) {
        for (let b = f.block; b <= (f.endBlock ?? blocks.length - 1); b++) {
          skipped.add(b);
        }
      }
    }
    blocks.forEach((block, i) => {
      const at = top ?? i;
      if (skipped.has(i) || !("type" in block)) {
        return;
      }
      const b = block as BodyContent;
      switch (b.type) {
        case "paragraph": {
          const heading = resolveHeadingLevel(doc, b, styles);
          const style = b.properties?.style;
          const sequences: string[] = [];
          forEachField([b], field => {
            const m = /^\s*SEQ\s+(\S+)/i.exec(field.instruction);
            if (m && !field.deleted) {
              sequences.push(m[1]!.toLowerCase());
            }
          });
          out.push({
            paragraph: b,
            top: at,
            heading: heading?.kind === "heading" ? heading.level : undefined,
            styleKeys: style ? [style.toLowerCase(), names.get(style) ?? ""] : [],
            sequences
          });
          break;
        }
        case "table":
          for (const row of b.rows) {
            for (const cell of row.cells) {
              visitList(cell.content, at);
            }
          }
          break;
        case "sdt":
          visitList(b.content, at);
          break;
        case "textBox":
          visitList(b.content ?? [], at);
          break;
        default:
          break;
      }
    });
  };
  visitList(doc.body, undefined);
  return out;
}

interface Selected {
  readonly candidate: Candidate;
  readonly level: number;
  readonly caption: boolean;
}

function select(spec: TocSpec, candidates: readonly Candidate[]): Selected[] {
  const out: Selected[] = [];
  const label = spec.caption?.label.toLowerCase();
  for (const candidate of candidates) {
    if (label !== undefined && candidate.sequences.includes(label)) {
      out.push({ candidate, level: 1, caption: true });
      continue;
    }
    const styled = candidate.styleKeys.map(k => spec.styles.get(k)).find(l => l !== undefined);
    if (styled !== undefined) {
      out.push({ candidate, level: styled, caption: false });
      continue;
    }
    const h = candidate.heading;
    if (spec.levels && h !== undefined && h >= spec.levels.min && h <= spec.levels.max) {
      out.push({ candidate, level: h, caption: false });
    }
  }
  return out;
}

/** The `_Toc` bookmark a paragraph already carries. */
function tocBookmarkOf(paragraph: Paragraph): string | undefined {
  const starts = [
    ...(paragraph.rangeMarkersBefore ?? []),
    ...paragraph.children.flatMap(c => [...(c.rangeMarkersBefore ?? []), c])
  ];
  for (const s of starts) {
    if (
      (s as { type?: string }).type === "bookmarkStart" &&
      /^_Toc/.test((s as BookmarkStart).name)
    ) {
      return (s as BookmarkStart).name;
    }
  }
  return undefined;
}

/** A table of contents plan: the document with its targets bookmarked, and the entries. */
export interface TocPlan {
  /** The document, with a `_Toc` bookmark on every target that lacked one. */
  readonly doc: DocxDocument;
  /**
   * The entry paragraphs a `TOC` instruction lists, numbered from `layout` (of
   * {@link doc}): each target's page as `pageText` shows it — the number its
   * section gives that page, in its section's format, as Word lists it.
   */
  entries(
    instruction: string,
    layout: LayoutResult,
    pageText: (page: number) => string,
    leader?: TabStopLeader
  ): Paragraph[];
  /** The `TOC` styles the entries of `instruction` use. */
  styles(instruction: string): { readonly levels: readonly number[]; readonly figures: boolean };
}

/**
 * Plan regenerating the tables of contents with `instructions`: find what
 * they list and bookmark it.
 */
export function planTableOfContents(doc: DocxDocument, instructions: readonly string[]): TocPlan {
  const candidates = collectCandidates(doc);
  const specs = new Map(instructions.map(i => [i, readTocSpec(i)]));
  const names = new Map<Paragraph, string>();
  const used = new Set<string>();
  let maxId = 0;
  walkDocument(
    doc,
    {
      visitBookmarkStart(b) {
        used.add(b.name);
        maxId = Math.max(maxId, b.id);
      }
    },
    { includeComments: true }
  );
  let serial = 100000000;
  const added = new Map<Paragraph, Paragraph>();
  for (const spec of specs.values()) {
    for (const { candidate } of select(spec, candidates)) {
      const p = candidate.paragraph;
      if (names.has(p)) {
        continue;
      }
      const existing = tocBookmarkOf(p);
      if (existing) {
        names.set(p, existing);
        continue;
      }
      let name: string;
      do {
        name = `_Toc${++serial}`;
      } while (used.has(name));
      used.add(name);
      names.set(p, name);
      const id = ++maxId;
      const start: BookmarkStart = { type: "bookmarkStart", id, name };
      const end: BookmarkEnd = { type: "bookmarkEnd", id, name };
      added.set(p, { ...p, children: [start, ...p.children, end] });
    }
  }
  const nextDoc =
    added.size === 0
      ? doc
      : {
          ...doc,
          body: mapBlockLists(doc.body, list => {
            let out: object[] | undefined;
            list.forEach((b, i) => {
              const replaced = added.get(b as Paragraph);
              if (replaced) {
                out ??= list.slice();
                out[i] = replaced;
              }
            });
            return (out ?? list) as typeof list;
          })
        };

  const section = doc.sectionProperties;
  const width =
    (section?.pageSize?.width ?? 12240) -
    (section?.margins?.left ?? 1440) -
    (section?.margins?.right ?? 1440);

  return {
    doc: nextDoc,
    entries(instruction, layout, pageText, leader = "dot") {
      const spec = specs.get(instruction) ?? readTocSpec(instruction);
      return select(spec, candidates).map(s => {
        const name = names.get(s.candidate.paragraph)!;
        const page = layout.bookmarkPages.get(name) ?? layout.contentPages[s.candidate.top] ?? 1;
        return entryParagraph(spec, s, entryText(spec, s), name, pageText(page), width, leader);
      });
    },
    styles(instruction) {
      const spec = specs.get(instruction) ?? readTocSpec(instruction);
      const selected = select(spec, candidates);
      return {
        levels: [...new Set(selected.filter(s => !s.caption).map(s => s.level))],
        figures: selected.some(s => s.caption)
      };
    }
  };
}

function entryText(spec: TocSpec, s: Selected): string {
  let text = extractParagraphText(s.candidate.paragraph);
  if (!spec.keepTabs) {
    text = text.replace(/\t/g, " ");
  }
  if (!spec.keepBreaks) {
    text = text.replace(/\n/g, " ");
  }
  if (s.caption && spec.caption?.bare) {
    const label = spec.caption.label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    text = text.replace(
      new RegExp(`^\\s*${label}\\s+[\\w.-]*\\d[\\w.-]*\\s*[:.\\-–—]?\\s*`, "i"),
      ""
    );
  }
  return text.trim();
}

function omitsPage(spec: TocSpec, level: number): boolean {
  const omit = spec.omitPages;
  return omit === true || (omit !== undefined && level >= omit.min && level <= omit.max);
}

function entryParagraph(
  spec: TocSpec,
  s: Selected,
  text: string,
  bookmark: string,
  page: string,
  width: number,
  leader: TabStopLeader
): Paragraph {
  const runs: Run[] = [{ content: [{ type: "text", text }] }];
  const withPage = !omitsPage(spec, s.level);
  if (withPage) {
    runs.push(
      spec.separator !== undefined
        ? { content: [{ type: "text", text: spec.separator }] }
        : { content: [{ type: "tab" }] },
      {
        content: [{ type: "field", instruction: ` PAGEREF ${bookmark} \\h `, cachedValue: page }]
      }
    );
  }
  const children: ParagraphChild[] = spec.hyperlink
    ? [{ type: "hyperlink", anchor: bookmark, history: true, children: runs } as Hyperlink]
    : runs;
  const properties: Mutable<NonNullable<Paragraph["properties"]>> = {
    style: s.caption ? "TableofFigures" : `TOC${Math.max(1, Math.min(9, s.level))}`
  };
  if (withPage && spec.separator === undefined) {
    properties.tabs = [{ type: "right", position: width, leader }];
  }
  return { type: "paragraph", properties, children };
}
