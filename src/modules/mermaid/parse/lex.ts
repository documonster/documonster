/**
 * Shared lexical helpers.
 *
 * Mermaid's surface syntax is line-oriented but not line-delimited: a statement can be
 * ended by a newline or a semicolon, and both forms appear in the same document. These
 * helpers turn a source string into the statement list every parser walks, so none of
 * them has to rediscover where a statement ends or what a comment looks like.
 */

/** One statement, and where it starts in the source (1-based line and column). */
export interface Statement {
  readonly text: string;
  readonly line: number;
  readonly column: number;
}

/**
 * Split source into statements.
 *
 * `%%` starts a comment that runs to the end of the line — but only outside quotes and
 * outside the `%%{ … }%%` init block, which is a directive rather than a comment. The
 * init block is skipped whole: this module has no configuration to apply from it, and
 * reading its contents as statements would be worse than ignoring it.
 */
export function splitStatements(source: string): string[] {
  return splitStatementsAt(source).map(statement => statement.text);
}

/** {@link splitStatements}, with each statement's position — for reporting what was skipped. */
export function splitStatementsAt(source: string): Statement[] {
  const out: Statement[] = [];
  const lines = stripFrontmatter(source.split(/\r\n|\r|\n/));
  for (let index = 0; index < lines.length; index++) {
    const raw = stripComment(lines[index]);
    if (raw.trim() === "") {
      continue;
    }
    let offset = 0;
    for (const part of splitOnSemicolons(raw)) {
      const text = part.trim();
      if (text !== "") {
        out.push({ text, line: index + 1, column: offset + part.indexOf(text) + 1 });
      }
      offset += part.length + 1;
    }
  }
  return out;
}

/**
 * Collects the statements a lenient parser could not read, so the diagram can say what it
 * left out instead of drawing a confident picture of part of the source.
 */
export class SkipLog {
  private readonly entries: Array<{ text: string; line: number; column: number }> = [];

  /** Record `statement` as not (fully) understood. */
  add(statement: Statement): void {
    this.entries.push({ text: statement.text, line: statement.line, column: statement.column });
  }

  /** Record a raw line (for the indentation-based grammars) as not understood. */
  addLine(text: string, line: number): void {
    const trimmed = text.trim();
    this.entries.push({ text: trimmed, line, column: text.indexOf(trimmed) + 1 });
  }

  /** `{ skipped }` when anything was skipped, else nothing — spread into the diagram. */
  report(): { skipped?: readonly { text: string; line: number; column: number }[] } {
    return this.entries.length === 0 ? {} : { skipped: this.entries };
  }
}

/**
 * The source's lines, with any frontmatter removed.
 *
 * For the grammars whose structure lives in whitespace — a mind map, a kanban board — and
 * so cannot go through {@link splitStatements}, which trims.
 */
export function linesOf(source: string): string[] {
  // Comments are stripped here with the quote-aware rule every other grammar gets. The three
  // callers each had their own `replace(/%%.*$/, "")`, which is not quote-aware: a mind-map
  // node called `100%% coverage` was truncated to `100` while the same literal survived in a
  // flowchart. Indentation is left alone, because for these grammars it *is* the syntax.
  return stripFrontmatter(source.split(/\r\n|\r|\n/)).map(line => stripComment(line));
}

/**
 * Drop a YAML frontmatter block.
 *
 * Mermaid's own documentation puts the title in frontmatter, so a diagram written the
 * recommended way begins with `---`. Left in place the dispatcher read that as the diagram
 * type and refused the whole document — every one of the supported types, for a construct
 * that carries no geometry.
 *
 * Only a block that *opens* on the very first non-empty line counts; `---` elsewhere is a
 * link in a flowchart.
 */
function stripFrontmatter(lines: readonly string[]): string[] {
  let first = 0;
  while (first < lines.length && lines[first].trim() === "") {
    first++;
  }
  if (lines[first]?.trim() !== "---") {
    return [...lines];
  }
  for (let i = first + 1; i < lines.length; i++) {
    if (lines[i].trim() === "---") {
      // Keep the line count so a reported line number still matches the source.
      return [
        ...lines.slice(0, first).map(() => ""),
        ...lines.slice(first, i + 1).map(() => ""),
        ...lines.slice(i + 1)
      ];
    }
  }
  return [...lines];
}

/**
 * Remove a trailing `%%` comment.
 *
 * Quote-aware, because a label may legitimately contain the characters: `A["50%% done"]`
 * is a node whose text ends in a percent sign, not a node followed by a comment.
 */
function stripComment(line: string): string {
  const directive = /^\s*%%\{[\s\S]*\}%%\s*$/;
  if (directive.test(line)) {
    return "";
  }
  let quote: '"' | null = null;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      quote = quote === null ? '"' : null;
      continue;
    }
    if (quote === null && ch === "%" && line[i + 1] === "%") {
      return line.slice(0, i);
    }
  }
  return line;
}

/**
 * Split on semicolons that end a statement.
 *
 * Quote-aware *and* bracket-aware. A semicolon also terminates Mermaid's numeric entity
 * (`#35;` is a literal `#`), so a label containing one would otherwise be cut in half and
 * the rest of the line read as a second statement.
 */
function splitOnSemicolons(line: string): string[] {
  const parts: string[] = [];
  let start = 0;
  let quote: '"' | null = null;
  let depth = 0;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      quote = quote === null ? '"' : null;
      continue;
    }
    if (quote !== null) {
      continue;
    }
    if (ch === "[" || ch === "(" || ch === "{") {
      depth++;
    } else if (ch === "]" || ch === ")" || ch === "}") {
      depth = Math.max(0, depth - 1);
    } else if (ch === ";" && depth === 0) {
      parts.push(line.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(line.slice(start));
  return parts;
}

/**
 * Strip one layer of matching quotes from a label.
 *
 * Mermaid uses quotes to let a label contain the characters that would otherwise end it,
 * so the quotes are syntax and must not survive into the text.
 */
export function unquote(text: string): string {
  const trimmed = text.trim();
  if (trimmed.length >= 2 && trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1);
  }
  return trimmed;
}

/**
 * Resolve the escapes a label may carry.
 *
 * `<br>` is Mermaid's line break in every diagram type, and `#35;`-style entities are how
 * an author writes a character the syntax would otherwise consume.
 */
export function decodeLabel(text: string): string {
  return unquote(text)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/#(\d+);/g, (_all, code: string) => String.fromCodePoint(Number(code)))
    .replace(/#([a-z]+);/gi, (all, name: string) => ENTITIES[name.toLowerCase()] ?? all)
    .trim();
}

const ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  nbsp: "\u00a0",
  semi: ";",
  colon: ":",
  hash: "#"
};

/**
 * The first occurrence of `token` that is not inside a quoted label.
 *
 * A class diagram looking for `-->` and a flowchart looking for a closing bracket were the
 * same scan written twice: a label may legitimately contain the very characters that would
 * otherwise end the statement, so a plain `indexOf` finds the wrong one. Returns `-1` when
 * there is none, like `indexOf`.
 */
export function indexOfUnquoted(text: string, token: string, from = 0): number {
  let quoted = false;
  for (let i = from; i < text.length; i++) {
    if (text[i] === '"') {
      quoted = !quoted;
      continue;
    }
    if (!quoted && text.startsWith(token, i)) {
      return i;
    }
  }
  return -1;
}
