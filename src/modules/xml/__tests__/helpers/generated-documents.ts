/**
 * Generated XML documents for the property tests of the byte-fed readers: records nested in
 * records of the same name, self-closing records, CDATA sections, comments and processing
 * instructions holding tags that are not tags, declared entities, characters of every UTF-8 length,
 * CRLF and lone CR line ends, several records on a line or the whole file on one.
 */

/** A deterministic generator, so a failure names a seed rather than a flake. */
export function random(seed: number): () => number {
  let state = seed >>> 0 || 1;
  return () => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return (state >>> 0) / 0x100000000;
  };
}

export interface GeneratedDocument {
  readonly xml: string;
  /** The DOCTYPE declaration, if the document has one: to give a record's text its entities. */
  readonly doctype: string;
}

export interface GenerateOptions {
  /** Put fake tags inside comments, CDATA sections and processing instructions. */
  readonly decoys?: boolean;
  /** Nest records inside records of the same name. */
  readonly sameNameNesting?: boolean;
  /** Write the whole document on one line. */
  readonly oneLine?: boolean;
  /** Give the document a DOCTYPE with declared entities. */
  readonly doctype?: boolean;
}

const TEXTS = [
  "plain",
  "a &amp; b",
  "&lt;tag&gt;",
  "&#x41;&#66;",
  "ü€",
  "😀",
  "日本語",
  " spaced "
];

export function generateDocument(seed: number, options: GenerateOptions = {}): GeneratedDocument {
  const next = random(seed);
  const pick = <T>(items: readonly T[]): T => items[Math.floor(next() * items.length)];
  const chance = (p: number): boolean => next() < p;
  const oneLine = options.oneLine ?? false;
  const breaks = ["\n", "\r\n", "\r", "\n  "];
  const br = (): string => (oneLine ? "" : pick(breaks));
  const doctype = options.doctype
    ? `<!DOCTYPE doc [ <!-- ] it's > --> <?pi ] ?> <!ENTITY e1 "ä &amp; b"> <!ENTITY e2 '[&e1;]'> ]>`
    : "";
  const texts = options.doctype ? [...TEXTS, "&e1;", "&e2;"] : TEXTS;
  let id = 0;

  const decoy = (): string =>
    pick([
      "<!-- <r id='x'> fake -->",
      "<![CDATA[ <r>not a record</r> ]]>",
      "<?pi <r> ?>",
      "<!-- </r> -->"
    ]);

  const content = (depth: number): string => {
    let out = "";
    const count = 1 + Math.floor(next() * 3);
    for (let i = 0; i < count; i++) {
      const roll = next();
      if (roll < 0.35) {
        out += pick(texts);
      } else if (roll < 0.5) {
        out += `<c${chance(0.5) ? ` k="${pick(texts)}"` : ""}/>`;
      } else if (roll < 0.65 && depth < 3) {
        out += `<n>${content(depth + 1)}</n>`;
      } else if (roll < 0.75 && depth < 3 && options.sameNameNesting) {
        out += record(depth + 1);
      } else if (roll < 0.85 && options.decoys) {
        out += decoy();
      } else {
        out += br();
      }
    }
    return out;
  };

  const record = (depth: number): string => {
    const attributes = `id="${id++}"${chance(0.3) ? ` t='${pick(texts)}'` : ""}`;
    if (chance(0.2)) {
      return `<r ${attributes}/>`;
    }
    return `<r ${attributes}>${content(depth)}</r>`;
  };

  let body = "";
  const items = 3 + Math.floor(next() * 6);
  for (let i = 0; i < items; i++) {
    const roll = next();
    if (roll < 0.55) {
      body += record(0);
    } else if (roll < 0.7) {
      body += `<o>${record(1)}${content(1)}</o>`;
    } else if (roll < 0.8 && options.decoys) {
      body += decoy();
    } else if (roll < 0.9) {
      body += `<!-- note -->`;
    }
    body += br();
  }
  const declaration = chance(0.5) ? `<?xml version="1.0" encoding="UTF-8"?>${br()}` : "";
  const xml = `${declaration}${doctype}${doctype ? br() : ""}<doc a="1">${br()}${body}</doc>${br()}`;
  return { xml, doctype };
}
