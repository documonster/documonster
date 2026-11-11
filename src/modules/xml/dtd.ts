/**
 * The DOCTYPE declaration: checking the parts the SAX parser reads, and the general entities an
 * internal subset declares.
 *
 * Nothing here fetches anything. An external DTD subset, an external entity and a parameter
 * entity are recognised so a document using them is well-formed, and are never read.
 */

import { charRefValue, isNameChar, isNameStartChar, isS, isXmlName } from "@xml/chars";

/** The five entities XML predefines (§4.6). */
const PREDEFINED: Readonly<Record<string, string>> = Object.assign(
  Object.create(null) as Record<string, string>,
  { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" }
);

/** Whether `name` is one of the five predefined entities. */
export function isPredefinedEntity(name: string): boolean {
  return PREDEFINED[name] !== undefined;
}

/** What an `<!ENTITY …>` declaration declares. */
export type EntityDeclaration =
  /** An internal general entity: `value` is its replacement text, character references expanded. */
  | { readonly kind: "internal"; readonly name: string; readonly value: string }
  /** An external general entity (`SYSTEM` / `PUBLIC`), which is never read. */
  | { readonly kind: "external"; readonly name: string }
  /** A parameter entity (`<!ENTITY % …>`), which is never expanded. */
  | { readonly kind: "parameter"; readonly name: string };

/** A small cursor over the text of one declaration. */
class Cursor {
  i = 0;
  constructor(readonly text: string) {}

  get done(): boolean {
    return this.i >= this.text.length;
  }

  /** Skip white space; whether there was any. */
  space(): boolean {
    const start = this.i;
    while (this.i < this.text.length && isS(this.text.charCodeAt(this.i))) {
      this.i++;
    }
    return this.i > start;
  }

  /** Read a Name, or `undefined` when none starts here. */
  name(): string | undefined {
    const start = this.i;
    const t = this.text;
    if (start >= t.length || !isNameStartChar(t.codePointAt(start)!)) {
      return undefined;
    }
    this.i += t.codePointAt(start)! > 0xffff ? 2 : 1;
    while (this.i < t.length) {
      const c = t.codePointAt(this.i)!;
      if (!isNameChar(c)) {
        break;
      }
      this.i += c > 0xffff ? 2 : 1;
    }
    return t.slice(start, this.i);
  }

  /** Read a name token (`Nmtoken`: name characters, any first), or `undefined` when none starts here. */
  nmtoken(): string | undefined {
    const start = this.i;
    const t = this.text;
    while (this.i < t.length) {
      const c = t.codePointAt(this.i)!;
      if (!isNameChar(c)) {
        break;
      }
      this.i += c > 0xffff ? 2 : 1;
    }
    return this.i > start ? t.slice(start, this.i) : undefined;
  }

  /** Read a quoted literal's content, or `undefined` when none starts here. */
  literal(): string | undefined {
    const q = this.text[this.i];
    if (q !== '"' && q !== "'") {
      return undefined;
    }
    const end = this.text.indexOf(q, this.i + 1);
    if (end < 0) {
      return undefined;
    }
    const value = this.text.slice(this.i + 1, end);
    this.i = end + 1;
    return value;
  }

  keyword(word: string): boolean {
    if (this.text.startsWith(word, this.i)) {
      this.i += word.length;
      return true;
    }
    return false;
  }
}

/** XML 1.0 §2.3 `PubidChar`, beside space, carriage return and line feed. */
const PUBID_PUNCTUATION = "-'()+,./:=?;!*#@$_%";

function isPubidLiteral(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    const alphanumeric =
      (c >= 0x30 && c <= 0x39) || (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a);
    if (
      !alphanumeric &&
      c !== 0x20 &&
      c !== 0xd &&
      c !== 0xa &&
      !PUBID_PUNCTUATION.includes(text[i])
    ) {
      return false;
    }
  }
  return true;
}

/** Read `SYSTEM S literal` or `PUBLIC S literal S literal`; a problem, or `undefined` once read. */
function readExternalId(c: Cursor, where: string): string | undefined {
  if (c.keyword("SYSTEM")) {
    if (!c.space() || c.literal() === undefined) {
      return `${where}: SYSTEM must be followed by a quoted system identifier`;
    }
    return undefined;
  }
  if (c.keyword("PUBLIC")) {
    if (!c.space()) {
      return `${where}: PUBLIC must be followed by a quoted public identifier`;
    }
    const pubid = c.literal();
    if (pubid === undefined || !isPubidLiteral(pubid)) {
      return `${where}: PUBLIC must be followed by a quoted public identifier`;
    }
    if (!c.space() || c.literal() === undefined) {
      return `${where}: a public identifier must be followed by a quoted system identifier`;
    }
    return undefined;
  }
  return `${where}: expected a quoted value, SYSTEM or PUBLIC`;
}

/**
 * Check the head of a DOCTYPE declaration — the text after `<!DOCTYPE` and before `[` or `>`:
 * `S Name (S ExternalID)? S?`. A problem, or `undefined` when it is well-formed.
 */
export function checkDoctypeHead(text: string): string | undefined {
  const c = new Cursor(text);
  if (!c.space() || c.name() === undefined) {
    return "DOCTYPE must be followed by the name of the root element";
  }
  const spaced = c.space();
  if (c.done) {
    return undefined;
  }
  if (!spaced) {
    return "malformed DOCTYPE declaration";
  }
  const problem = readExternalId(c, "DOCTYPE");
  if (problem !== undefined) {
    return problem;
  }
  c.space();
  return c.done ? undefined : "malformed DOCTYPE declaration";
}

/**
 * The replacement text of an entity value (§4.5): character references expanded, general entity
 * references kept to be expanded where the entity is used. A problem string when the value is
 * malformed.
 */
function replacementText(value: string): { text: string } | { problem: string } {
  let out = "";
  let from = 0;
  for (let i = 0; i < value.length; i++) {
    const ch = value.charCodeAt(i);
    if (ch === 0x25) {
      return { problem: "a parameter-entity reference may not appear in an entity value here" };
    }
    if (ch !== 0x26) {
      continue;
    }
    const semi = value.indexOf(";", i);
    const ref = semi < 0 ? "" : value.slice(i + 1, semi);
    if (ref.charCodeAt(0) === 0x23) {
      const char = charRefValue(ref);
      if (char === undefined) {
        return { problem: `invalid character reference &${ref}; in an entity value` };
      }
      out += value.slice(from, i) + char;
      from = semi + 1;
    } else if (!isXmlName(ref)) {
      return { problem: "malformed entity reference in an entity value" };
    }
    i = semi;
  }
  return { text: out + value.slice(from) };
}

/**
 * What well-formedness needs to know of an internal entity's replacement text where it may be used
 * in an attribute value — without keeping the text: whether it holds a `<` of its own, whether a
 * reference in it is malformed, and the declared entities it refers to, once per name.
 */
export interface EntitySummary {
  readonly lt: boolean;
  readonly malformed: boolean;
  readonly refs: readonly string[];
}

export function summarizeEntity(text: string): EntitySummary {
  const refs = new Set<string>();
  let malformed = false;
  for (let i = text.indexOf("&"); i >= 0;) {
    const semi = text.indexOf(";", i);
    const ref = semi < 0 ? "" : text.slice(i + 1, semi);
    i = semi < 0 ? -1 : text.indexOf("&", semi + 1);
    if (ref.charCodeAt(0) === 0x23 ? charRefValue(ref) === undefined : !isXmlName(ref)) {
      malformed = true;
    } else if (ref.charCodeAt(0) !== 0x23 && !isPredefinedEntity(ref)) {
      refs.add(ref);
    }
  }
  return { lt: text.includes("<"), malformed, refs: [...refs] };
}

/**
 * What `<!` … `>` in an internal subset declares, or a problem string when it is malformed:
 * `references` are the general entities an attribute list's defaults refer to, to be checked once
 * every declaration is known.
 */
export type MarkupDeclaration =
  | {
      readonly entity?: EntityDeclaration;
      readonly references?: readonly string[];
      readonly attlist?: AttlistDeclaration;
    }
  | string;

/** An attribute an `<!ATTLIST>` declares: whether its type is other than CDATA, and its default as written. */
export interface AttributeDefinition {
  readonly name: string;
  readonly tokenized: boolean;
  /** The default value's literal, references unexpanded; absent for `#REQUIRED` and `#IMPLIED`. */
  readonly value?: string;
}

/** What an `<!ATTLIST>` declares, for a parser that uses the declarations it reads (XML 1.0 §5.1). */
export interface AttlistDeclaration {
  readonly element: string;
  readonly attributes: readonly AttributeDefinition[];
}

/**
 * Read a markup declaration of an internal subset — `text` is what lies between `<!` and `>` — to
 * the grammar of XML 1.0 §3.2 (`elementdecl`), §3.3 (`AttlistDecl`), §4.2 (`EntityDecl`) and §4.7
 * (`NotationDecl`). An ENTITY declaration also says what it declares; the others are only checked.
 */
export function readMarkupDeclaration(text: string): MarkupDeclaration {
  const c = new Cursor(text);
  let problem: string | undefined;
  const references: string[] = [];
  let attlist: AttlistDeclaration | undefined;
  if (c.keyword("ENTITY")) {
    return readEntityDeclaration(c);
  }
  if (c.keyword("ELEMENT")) {
    problem = readElementDeclaration(c);
  } else if (c.keyword("ATTLIST")) {
    const read = readAttlistDeclaration(c, references);
    if (typeof read === "string") {
      problem = read;
    } else {
      attlist = read;
    }
  } else if (c.keyword("NOTATION")) {
    problem = readNotationDeclaration(c);
  } else {
    return "expected ELEMENT, ATTLIST, ENTITY, NOTATION or a comment after <!";
  }
  if (problem === undefined) {
    c.space();
    if (!c.done) {
      problem = `unexpected ${describeAt(c)}`;
    }
  }
  if (problem !== undefined) {
    return problem;
  }
  return {
    ...(references.length > 0 && { references }),
    ...(attlist !== undefined && attlist.attributes.length > 0 && { attlist })
  };
}

/**
 * Check the start of a markup declaration too long to be gathered whole: its keyword, and the name it
 * declares. A problem, or `undefined`.
 */
export function checkDeclarationStart(text: string): string | undefined {
  const c = new Cursor(text);
  for (const keyword of ["ELEMENT", "ATTLIST", "ENTITY", "NOTATION"]) {
    if (c.keyword(keyword)) {
      if (!c.space()) {
        return `<!${keyword} must be followed by a name`;
      }
      if (keyword === "ENTITY" && c.keyword("%") && !c.space()) {
        return "<!ENTITY % must be followed by a name";
      }
      return c.name() === undefined ? `<!${keyword} must be followed by a name` : undefined;
    }
  }
  return "expected ELEMENT, ATTLIST, ENTITY, NOTATION or a comment after <!";
}

/** What is at the cursor, for a message: a parameter-entity reference is named as one. */
function describeAt(c: Cursor): string {
  if (c.done) {
    return "end of the declaration";
  }
  const ch = c.text[c.i];
  return ch === "%"
    ? "parameter-entity reference: one may not appear inside a markup declaration here"
    : `character ${JSON.stringify(ch)}`;
}

/** `S Name`, the name a declaration is about. */
function readDeclaredName(c: Cursor, keyword: string): string | undefined {
  return c.space() && c.name() !== undefined
    ? undefined
    : `<!${keyword} must be followed by a name`;
}

/** `elementdecl`: `S Name S contentspec`. */
function readElementDeclaration(c: Cursor): string | undefined {
  const named = readDeclaredName(c, "ELEMENT");
  if (named !== undefined) {
    return named;
  }
  if (!c.space()) {
    return "<!ELEMENT: expected a content model after the name";
  }
  if (c.keyword("EMPTY") || c.keyword("ANY")) {
    return undefined;
  }
  if (c.text[c.i] !== "(") {
    return `<!ELEMENT: expected EMPTY, ANY or a content model, found ${describeAt(c)}`;
  }
  const at = c.i;
  c.i++;
  c.space();
  if (c.keyword("#PCDATA")) {
    return readMixed(c);
  }
  c.i = at;
  const problem = readContentParticle(c);
  return problem === undefined ? undefined : `<!ELEMENT: ${problem}`;
}

/** `Mixed`, after `( S? #PCDATA`: `(S? '|' S? Name)* S? ')*'`, or `S? ')'` with no names. */
function readMixed(c: Cursor): string | undefined {
  let names = 0;
  for (;;) {
    c.space();
    if (c.text[c.i] === ")") {
      c.i++;
      if (c.text[c.i] === "*") {
        c.i++;
      } else if (names > 0) {
        return "<!ELEMENT: mixed content naming elements must end with )*";
      }
      return undefined;
    }
    if (c.text[c.i] !== "|") {
      return `<!ELEMENT: expected | or ) in mixed content, found ${describeAt(c)}`;
    }
    c.i++;
    c.space();
    if (c.name() === undefined) {
      return `<!ELEMENT: expected an element name in mixed content, found ${describeAt(c)}`;
    }
    names++;
  }
}

/** `cp`: `(Name | choice | seq) ('?' | '*' | '+')?`, where a group is `( cp (S? sep S? cp)* )`. */
function readContentParticle(c: Cursor, depth = 0): string | undefined {
  if (depth > 256) {
    return "content model nested too deeply";
  }
  if (c.text[c.i] === "(") {
    c.i++;
    c.space();
    let separator: string | undefined;
    for (;;) {
      const problem = readContentParticle(c, depth + 1);
      if (problem !== undefined) {
        return problem;
      }
      c.space();
      const ch = c.text[c.i];
      if (ch === ")") {
        c.i++;
        break;
      }
      if (ch !== "|" && ch !== ",") {
        return `expected |, , or ) in a content model, found ${describeAt(c)}`;
      }
      if (separator !== undefined && ch !== separator) {
        return "a content model group may not mix | and ,";
      }
      separator = ch;
      c.i++;
      c.space();
    }
  } else if (c.name() === undefined) {
    return `expected an element name or ( in a content model, found ${describeAt(c)}`;
  }
  const occurrence = c.text[c.i];
  if (occurrence === "?" || occurrence === "*" || occurrence === "+") {
    c.i++;
  }
  return undefined;
}

const TOKENIZED_TYPES = ["IDREFS", "IDREF", "ID", "ENTITIES", "ENTITY", "NMTOKENS", "NMTOKEN"];

/** `AttlistDecl`: `S Name (S Name S AttType S DefaultDecl)*`. */
function readAttlistDeclaration(c: Cursor, references: string[]): AttlistDeclaration | string {
  if (!c.space()) {
    return "<!ATTLIST must be followed by a name";
  }
  const element = c.name();
  if (element === undefined) {
    return "<!ATTLIST must be followed by a name";
  }
  const attributes: AttributeDefinition[] = [];
  for (;;) {
    const spaced = c.space();
    if (c.done) {
      return { element, attributes };
    }
    if (!spaced) {
      return `<!ATTLIST: unexpected ${describeAt(c)}`;
    }
    const attribute = c.name();
    if (attribute === undefined) {
      return `<!ATTLIST: expected an attribute name, found ${describeAt(c)}`;
    }
    if (!c.space()) {
      return `<!ATTLIST: expected the type of ${attribute}`;
    }
    const tokenized = !c.keyword("CDATA");
    const typeProblem = tokenized ? readAttributeType(c, attribute) : undefined;
    if (typeProblem !== undefined) {
      return typeProblem;
    }
    if (!c.space()) {
      return `<!ATTLIST: expected the default of ${attribute}`;
    }
    if (c.keyword("#REQUIRED") || c.keyword("#IMPLIED")) {
      attributes.push({ name: attribute, tokenized });
      continue;
    }
    if (c.keyword("#FIXED") && !c.space()) {
      return `<!ATTLIST: #FIXED must be followed by a value for ${attribute}`;
    }
    const value = c.literal();
    if (value === undefined) {
      return `<!ATTLIST: expected #REQUIRED, #IMPLIED, #FIXED or a quoted default for ${attribute}`;
    }
    const valueProblem = checkAttributeDefault(value, references);
    if (valueProblem !== undefined) {
      return `<!ATTLIST: the default of ${attribute} ${valueProblem}`;
    }
    attributes.push({ name: attribute, tokenized, value });
  }
}

/** `AttType` other than `CDATA`: a tokenized type, `NOTATION S (names)` or an enumeration `(nmtokens)`. */
function readAttributeType(c: Cursor, attribute: string): string | undefined {
  if (c.keyword("NOTATION")) {
    if (!c.space() || c.text[c.i] !== "(") {
      return `<!ATTLIST: NOTATION must be followed by a list of notation names for ${attribute}`;
    }
    return readTokenList(c, attribute, false);
  }
  for (const type of TOKENIZED_TYPES) {
    if (c.keyword(type)) {
      return undefined;
    }
  }
  if (c.text[c.i] === "(") {
    return readTokenList(c, attribute, true);
  }
  return `<!ATTLIST: expected the type of ${attribute}, found ${describeAt(c)}`;
}

/** `'(' S? token (S? '|' S? token)* S? ')'`, of names or of name tokens. */
function readTokenList(c: Cursor, attribute: string, nmtokens: boolean): string | undefined {
  c.i++;
  for (;;) {
    c.space();
    if ((nmtokens ? c.nmtoken() : c.name()) === undefined) {
      return `<!ATTLIST: expected a ${nmtokens ? "name token" : "notation name"} in the type of ${attribute}, found ${describeAt(c)}`;
    }
    c.space();
    const ch = c.text[c.i];
    c.i++;
    if (ch === ")") {
      return undefined;
    }
    if (ch !== "|") {
      c.i--;
      return `<!ATTLIST: expected | or ) in the type of ${attribute}, found ${describeAt(c)}`;
    }
  }
}

/**
 * An attribute default (`AttValue`, §3.3): no `<`, and every `&` a well-formed reference. The general
 * entities it names are added to `references`: whether they are declared is known only at the end
 * of the subset.
 */
function checkAttributeDefault(value: string, references: string[]): string | undefined {
  if (value.includes("<")) {
    return "may not contain <";
  }
  for (let i = value.indexOf("&"); i >= 0;) {
    const semi = value.indexOf(";", i);
    const ref = semi < 0 ? "" : value.slice(i + 1, semi);
    if (ref.charCodeAt(0) === 0x23 ? charRefValue(ref) === undefined : !isXmlName(ref)) {
      return "holds a malformed reference";
    }
    if (ref.charCodeAt(0) !== 0x23 && !isPredefinedEntity(ref)) {
      references.push(ref);
    }
    i = value.indexOf("&", semi + 1);
  }
  return undefined;
}

/** `NotationDecl`: `S Name S (ExternalID | PublicID)`. */
function readNotationDeclaration(c: Cursor): string | undefined {
  const named = readDeclaredName(c, "NOTATION");
  if (named !== undefined) {
    return named;
  }
  if (!c.space()) {
    return "<!NOTATION: expected SYSTEM or PUBLIC after the name";
  }
  if (c.keyword("PUBLIC")) {
    // A public identifier alone is enough here (`PublicID`); a system identifier may follow it.
    if (!c.space()) {
      return "<!NOTATION: PUBLIC must be followed by a quoted public identifier";
    }
    const pubid = c.literal();
    if (pubid === undefined || !isPubidLiteral(pubid)) {
      return "<!NOTATION: PUBLIC must be followed by a quoted public identifier";
    }
    const at = c.i;
    if (c.space() && c.literal() === undefined) {
      c.i = at;
    }
    return undefined;
  }
  return readExternalId(c, "<!NOTATION");
}

function readEntityDeclaration(c: Cursor): MarkupDeclaration {
  if (!c.space()) {
    return "<!ENTITY must be followed by a name";
  }
  const parameter = c.keyword("%");
  if (parameter && !c.space()) {
    return "<!ENTITY % must be followed by a name";
  }
  const name = c.name();
  if (name === undefined) {
    return "<!ENTITY must be followed by a name";
  }
  if (!c.space()) {
    return `<!ENTITY ${name}: expected a value`;
  }
  let entity: EntityDeclaration;
  const literal = c.literal();
  if (literal !== undefined) {
    const replaced = replacementText(literal);
    if ("problem" in replaced) {
      return `<!ENTITY ${name}: ${replaced.problem}`;
    }
    entity = parameter
      ? { kind: "parameter", name }
      : { kind: "internal", name, value: replaced.text };
  } else {
    const problem = readExternalId(c, `<!ENTITY ${name}`);
    if (problem !== undefined) {
      return problem;
    }
    entity = parameter ? { kind: "parameter", name } : { kind: "external", name };
    const spaced = c.space();
    if (!parameter && spaced && c.keyword("NDATA")) {
      if (!c.space() || c.name() === undefined) {
        return `<!ENTITY ${name}: NDATA must be followed by a notation name`;
      }
    }
  }
  c.space();
  return c.done ? { entity } : `<!ENTITY ${name}: unexpected text after the value`;
}

/** Why an entity could not be expanded: the parser reports it as a parse error. */
export class EntityProblem {
  /**
   * @param limit - Whether expanding it would pass the parser's `maxEntityExpansions`, rather than
   *   the entity being malformed.
   */
  constructor(
    readonly message: string,
    readonly limit = false
  ) {}
}

/** An entity being written out by `text`: `value[at]` is where it goes on. */
interface Expansion {
  readonly name: string;
  readonly value: string;
  at: number;
}

/** A step of a walk over the entities a replacement text refers to: `refs[next]` is next. */
interface Visit {
  readonly name: string;
  readonly refs: readonly string[];
  next: number;
  count: number;
  /** The length of its text so far: its own characters, and those of the references done. */
  length: number;
}

/**
 * The internal general entities of a document, expanded where they are used (§4.4): the replacement
 * text is read again for references, and in an attribute value its white space becomes spaces
 * (§3.3.3). Replacement text that holds markup is refused rather than parsed, as is recursion.
 *
 * An entity's cost — how many entity references one use of it expands, nested ones included — is
 * counted before any text is built, and the count stops as soon as it passes the budget the parser
 * gives (its `maxEntityExpansions` still unspent), so a "billion laughs" document is refused without
 * building the string, and without walking all of it either. Both walks keep their own stack rather
 * than recursing, so a long chain of entities each naming the next is refused or expanded, never a
 * `RangeError` from the JavaScript stack. The same walk measures the text each entity expands to, so
 * the parser refuses a text too long before it is built: a few entities used many times stay well
 * within a count of expansions and can still make a string larger than memory. Counts, lengths and
 * texts are memoised.
 */
export class DeclaredEntities {
  private readonly counts = new Map<string, number>();
  /** The length of the text each entity expands to, in UTF-16 units: the same in content and in an attribute value. */
  private readonly lengths = new Map<string, number>();
  private readonly references = new Map<string, readonly string[]>();
  /** The length of each entity's own text: everything but the declared entities it refers to. */
  private readonly ownLengths = new Map<string, number>();
  /** The texts of entities used in content, and in attribute values: only those used, not what they refer to. */
  private readonly content = new Map<string, string>();
  private readonly attribute = new Map<string, string>();

  /** @param values - Replacement text by name, on a null-prototype object. */
  constructor(readonly values: Readonly<Record<string, string>>) {}

  has(name: string): boolean {
    return typeof this.values[name] === "string";
  }

  /**
   * The declared entities the replacement text of `name` refers to, once per reference, in order.
   * A malformed or undefined reference is a problem.
   */
  private referencesOf(name: string): readonly string[] {
    let refs = this.references.get(name);
    if (refs !== undefined) {
      return refs;
    }
    const value = this.values[name];
    const found: string[] = [];
    let own = value.length;
    // Each reference is read whole and stepped over, as `build` reads it: a reference is refused
    // here, before anything it refers to is built, and the length it adds is never negative.
    for (let i = value.indexOf("&"); i >= 0;) {
      const semi = value.indexOf(";", i);
      const ref = semi < 0 ? "" : value.slice(i + 1, semi);
      i = semi < 0 ? -1 : value.indexOf("&", semi + 1);
      if (ref.charCodeAt(0) === 0x23) {
        const char = charRefValue(ref);
        if (char === undefined) {
          throw new EntityProblem(
            `invalid character reference &${ref.length > 40 ? `${ref.slice(0, 40)}…` : ref}; in entity ${name}`
          );
        }
        own -= ref.length + 2 - char.length;
        continue;
      }
      if (!isXmlName(ref)) {
        throw new EntityProblem(`malformed entity reference in the replacement text of ${name}`);
      }
      own -= ref.length + 2;
      if (isPredefinedEntity(ref)) {
        own++;
      } else {
        if (!this.has(ref)) {
          throw new EntityProblem(`undefined entity ${ref} in the replacement text of ${name}`);
        }
        found.push(ref);
      }
    }
    refs = found;
    this.references.set(name, refs);
    this.ownLengths.set(name, own);
    return refs;
  }

  /**
   * How many entity references one use of `name` expands, itself included. Past `budget`, an
   * {@link EntityProblem} marked `limit` — counted no further than needed to know it.
   */
  expansions(name: string, budget: number): number {
    const known = this.counts.get(name);
    if (known !== undefined) {
      if (known > budget) {
        throw new EntityProblem(`entity ${name} expands too many entities`, true);
      }
      return known;
    }
    // The use of `name` itself costs one, before anything it refers to.
    if (budget < 1) {
      throw new EntityProblem(`entity ${name} expands too many entities`, true);
    }
    const stack: Visit[] = [this.visit(name, 1)];
    const active = new Set<string>([name]);
    // The counts on the stack add up to no more than `name`'s: once they pass the budget, so does it.
    let atLeast = 1;
    while (stack.length > 0) {
      const top = stack[stack.length - 1];
      if (top.next === top.refs.length) {
        stack.pop();
        active.delete(top.name);
        this.counts.set(top.name, top.count);
        this.lengths.set(top.name, top.length);
        if (stack.length > 0) {
          stack[stack.length - 1].count += top.count;
          stack[stack.length - 1].length += top.length;
        }
        continue;
      }
      const ref = top.refs[top.next++];
      const counted = this.counts.get(ref);
      if (counted !== undefined) {
        top.count += counted;
        top.length += this.lengths.get(ref)!;
        atLeast += counted;
      } else if (active.has(ref)) {
        throw new EntityProblem(`entity ${ref} refers to itself`);
      } else {
        active.add(ref);
        stack.push(this.visit(ref, 1));
        atLeast++;
      }
      if (atLeast > budget) {
        throw new EntityProblem(`entity ${name} expands too many entities`, true);
      }
    }
    return this.counts.get(name)!;
  }

  /**
   * The length of the text `name` expands to, in UTF-16 units — known without building it, so a
   * text too long to build is refused before it is. Call {@link expansions} first: it measures it.
   */
  length(name: string): number {
    return this.lengths.get(name)!;
  }

  private visit(name: string, count: number): Visit {
    const refs = this.referencesOf(name);
    return { name, refs, next: 0, count, length: this.ownLengths.get(name)! };
  }

  /**
   * The text `name` expands to in content, or in an attribute value. Call {@link expansions} first:
   * it finds what refers to itself and checks every reference, and bounds what this builds.
   *
   * Written straight into one string, references expanded in place as they are met: what an entity
   * refers to is not kept as a text of its own. Keeping each — the memo this used to be — held the
   * whole chain of texts a long chain of entities passes through, which no limit on what is used
   * bounds. Only a text that was used is kept, so what is kept is never more than what was counted.
   */
  text(name: string, inAttribute: boolean): string {
    const memo = inAttribute ? this.attribute : this.content;
    const known = memo.get(name);
    if (known !== undefined) {
      return known;
    }
    let out = "";
    const stack: Expansion[] = [{ name, value: this.values[name], at: 0 }];
    while (stack.length > 0) {
      const frame = stack[stack.length - 1];
      const value = frame.value;
      let from = frame.at;
      let nested = false;
      for (let i = from; i < value.length; i++) {
        const ch = value.charCodeAt(i);
        if (ch === 0x3c) {
          throw new EntityProblem(
            inAttribute
              ? `entity ${frame.name} puts a < in an attribute value`
              : `entity ${frame.name} contains markup, which is not expanded`
          );
        }
        if (inAttribute && (ch === 0x9 || ch === 0xa || ch === 0xd)) {
          out += value.slice(from, i) + " ";
          from = i + 1;
        } else if (
          !inAttribute &&
          ch === 0x3e &&
          i >= 2 &&
          value.charCodeAt(i - 1) === 0x5d &&
          value.charCodeAt(i - 2) === 0x5d
        ) {
          // Character data may not hold a `]]>` (§2.4), and replacement text in content is read as
          // content. Only one of its own: `]` and `]>` from two entities, or `]]&gt;`, are not one.
          throw new EntityProblem(
            `entity ${frame.name} contains ]]>, which character data may not`
          );
        } else if (ch === 0x26) {
          const semi = value.indexOf(";", i);
          const ref = value.slice(i + 1, semi);
          out += value.slice(from, i);
          from = semi + 1;
          i = semi;
          if (ref.charCodeAt(0) === 0x23) {
            // Checked when the references were read.
            out += charRefValue(ref)!;
            continue;
          }
          const fixed = PREDEFINED[ref] ?? memo.get(ref);
          if (fixed !== undefined) {
            out += fixed;
            continue;
          }
          frame.at = from;
          stack.push({ name: ref, value: this.values[ref], at: 0 });
          nested = true;
          break;
        }
      }
      if (!nested) {
        out += value.slice(from);
        stack.pop();
      }
    }
    memo.set(name, out);
    return out;
  }
}

/** XML 1.0 §2.8 `XMLDecl` after `<?xml`, its white space and its `?>` removed: `VersionInfo EncodingDecl? SDDecl? S?`. */
const XML_DECLARATION =
  /^version[ \t\r\n]*=[ \t\r\n]*(["'])1\.[0-9]+\1(?:[ \t\r\n]+encoding[ \t\r\n]*=[ \t\r\n]*(["'])[A-Za-z][A-Za-z0-9._-]*\2)?(?:[ \t\r\n]+standalone[ \t\r\n]*=[ \t\r\n]*(["'])(?:yes|no)\3)?[ \t\r\n]*$/;

/** Check the body of an XML declaration — what follows `<?xml` and its white space. A problem, or `undefined`. */
export function checkXmlDeclaration(body: string): string | undefined {
  return XML_DECLARATION.test(body)
    ? undefined
    : "malformed XML declaration: expected version, then optionally encoding and standalone";
}
