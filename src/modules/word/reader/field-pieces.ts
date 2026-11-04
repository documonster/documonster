/**
 * DOCX reader — complex fields read field character by field character.
 *
 * A complex field is a run of field characters: `begin`, the instruction,
 * `separate`, the result, `end`. Read whole, it becomes one
 * {@link FieldContent} whose result is a string — which is faithful only
 * when the field lies in one paragraph — directly, or wholly inside one
 * hyperlink, smart tag or custom XML element — nests no other field, and its
 * field characters share no run with other content. Any other field is read
 * in pieces: each of its field characters becomes a {@link FieldContent} with
 * `fieldChar` set, at the place it was — inside a revision or hyperlink when
 * it was in one — and everything between them, in whichever paragraphs and
 * wrappers it lies, is read as the content it is. Characters inside markup
 * the reader keeps verbatim (an inline content control) stay in it.
 *
 * {@link splitFieldPieces} does this one paragraph at a time, so a field
 * open at the end of a paragraph carries over to the next: the body is read
 * in order, by the buffered and the streaming reader alike. It hands the
 * paragraph back with each piece's runs replaced by one marker run
 * ({@link fieldPieceOf}); the fields it leaves whole stay for the reader's
 * field state machine.
 */

import { isRangeMarkerElement } from "@word/reader/range-markers";
import type { XmlElement, XmlNode } from "@xml/types";

/** A field read in pieces that a paragraph opened and has not yet ended. */
export interface OpenFieldPiece {
  /** Instruction text met since the field's last piece, written before its next. */
  pending: string;
}

/** What a marker run from {@link splitFieldPieces} stands for. */
export interface FieldPiece {
  readonly fieldChar: "begin" | "separate" | "end";
  /** The instruction text written beside the character. */
  instruction: string;
  /** The `w:fldChar` element (its `w:ffData` holds a legacy form field). */
  readonly fldChar: XmlElement;
  /** The run properties of the run that held the character. */
  readonly rPr: XmlElement | undefined;
}

const pieces = new WeakMap<XmlElement, FieldPiece>();

/** The field piece `run` stands for, when it is a marker run. */
export function fieldPieceOf(run: XmlElement): FieldPiece | undefined {
  return pieces.get(run);
}

const codeTexts = new WeakMap<XmlElement, string>();

/**
 * The instruction text `run` stands for, when it is a marker run holding a
 * field's instruction text in place — where neither the field's piece before
 * it nor the one after it is its neighbour (a field nested in the
 * instruction lies between).
 */
export function fieldCodeTextOf(run: XmlElement): string | undefined {
  return codeTexts.get(run);
}

function localName(el: XmlElement): string {
  const i = el.name.indexOf(":");
  return i < 0 ? el.name : el.name.slice(i + 1);
}

/**
 * Whether qualified `name` has local name `local`, without allocating: asked
 * of every element of every run read.
 */
function hasLocal(name: string, local: string): boolean {
  const offset = name.length - local.length;
  if (offset < 0 || (offset > 0 && name.charCodeAt(offset - 1) !== 58)) {
    return false;
  }
  return name.startsWith(local, offset);
}

/** What an element of a run is to a field: a field character, instruction text, or neither. */
type FieldEventKind = "begin" | "separate" | "end" | "instr" | undefined;

/**
 * A field met in the paragraph, or carried into it: the record is what
 * `open` carries to the next paragraph.
 */
interface FieldRecord extends OpenFieldPiece {
  /** Read in pieces. */
  pieces: boolean;
  /** Its separate character has been met. */
  separated: boolean;
  /** The element its characters lie in; `null` once that is more than one. */
  container: XmlElement | null | undefined;
}

/**
 * Markup the reader drops wherever it is, so a field holding it is still
 * read whole: proofing marks, and Word's record of where it last broke a page.
 */
const IGNORED = new Set(["proofErr", "lastRenderedPageBreak"]);

/**
 * Whether `el` is result content a field read whole holds in its
 * `cachedValue`: text, and a line break (`\n`, as the writer writes one).
 */
export function isResultText(el: XmlElement): boolean {
  if (hasLocal(el.name, "t")) {
    return true;
  }
  if (!hasLocal(el.name, "br")) {
    return false;
  }
  const type = el.attributes["w:type"] ?? el.attributes["type"];
  return (type === undefined || type === "textWrapping") && !el.attributes["w:clear"];
}

/** What a child of a run is to the first pass of {@link splitFieldPieces}. */
type RunChildKind = NonNullable<FieldEventKind> | "props" | "ignored" | "text" | "other";

/**
 * {@link fieldEvent}, or what else the run child is. Asked of every child of
 * every run read, so the names Word writes are matched as they are, without
 * taking the local name apart.
 */
function runChildKind(el: XmlElement): RunChildKind {
  switch (el.name) {
    case "w:t":
      return "text";
    case "w:rPr":
      return "props";
    case "w:fldChar":
    case "w:instrText":
      return fieldEvent(el) ?? "other";
    case "w:proofErr":
    case "w:lastRenderedPageBreak":
      return "ignored";
    default:
      break;
  }
  const kind = fieldEvent(el);
  if (kind !== undefined) {
    return kind;
  }
  const name = el.name;
  if (hasLocal(name, "rPr")) {
    return "props";
  }
  if (hasLocal(name, "proofErr") || hasLocal(name, "lastRenderedPageBreak")) {
    return "ignored";
  }
  return isResultText(el) ? "text" : "other";
}

function fieldEvent(el: XmlElement): FieldEventKind {
  const name = el.name;
  if (!hasLocal(name, "fldChar")) {
    return hasLocal(name, "instrText") || hasLocal(name, "delInstrText") ? "instr" : undefined;
  }
  const type = el.attributes["w:fldCharType"] ?? el.attributes["fldCharType"];
  return type === "begin" || type === "separate" || type === "end" ? type : undefined;
}

function textOf(el: XmlElement): string {
  let text = "";
  for (const c of el.children) {
    if (c.type === "text" || c.type === "cdata") {
      text += c.value;
    }
  }
  return text;
}

/**
 * Wrappers the reader reads the content of, in place: their runs are
 * paragraph content, so a field's characters inside one are replaced as
 * they are anywhere else.
 */
const TRANSPARENT = new Set([
  "hyperlink",
  "ins",
  "del",
  "moveTo",
  "moveFrom",
  "smartTag",
  "customXml"
]);

/**
 * The containers a field lying wholly in is read whole: the reader's field
 * state machine reads those runs. A revision's runs are read without it, so a
 * field inside `w:ins`, `w:del` or a move is always read in pieces.
 */
const WHOLE = new Set(["p", "hyperlink", "smartTag", "customXml"]);

/**
 * Wrappers kept as markup (an inline `w:sdt`, a bidirectional `w:dir` /
 * `w:bdo`): the field characters inside one are written back with it, but are
 * followed, so a field crossing its boundary is read in pieces outside it.
 */
const OPAQUE = new Set(["sdt", "dir", "bdo"]);

/** Wrapper properties, read with the wrapper. */
const WRAPPER_PROPERTIES = new Set(["pPr", "smartTagPr", "customXmlPr"]);

/** The state of the first pass of {@link splitFieldPieces} over one paragraph. */
interface Scan {
  /** The fields open at the current point, innermost last. */
  readonly stack: FieldRecord[];
  /** Each field character and instruction met outside kept markup, with its field. */
  readonly ownerEls: XmlElement[];
  readonly ownerRecords: FieldRecord[];
  forced: ReadonlySet<XmlElement> | undefined;
}

const SCAN: Scan = { stack: [], ownerEls: [], ownerRecords: [], forced: undefined };

/** Content met inside the open fields: text-only result content leaves them whole. */
function holdContent(scan: Scan, textOnly: boolean): void {
  for (const record of scan.stack) {
    if (!textOnly || !record.separated) {
      record.pieces = true;
    }
  }
}

/**
 * The first pass over one run. `whole` says whether `container` is one a
 * field lying wholly in is read whole in ({@link WHOLE}); `kept`, that the
 * run is inside markup kept as it is.
 */
function scanRun(
  scan: Scan,
  run: XmlElement,
  container: XmlElement,
  whole: boolean,
  kept: boolean
): void {
  const stack = scan.stack;
  // The fields this run holds a character of: nearly always none or one.
  let first: FieldRecord | undefined;
  let more: FieldRecord[] | undefined;
  let other = false;
  let nonText = false;
  for (const rc of run.children) {
    if (rc.type !== "element") {
      continue;
    }
    const kind = runChildKind(rc);
    if (kind === "text" || kind === "other") {
      other = true;
      nonText ||= kind === "other";
      continue;
    }
    if (kind === "props" || kind === "ignored") {
      continue;
    }
    let record: FieldRecord | undefined;
    if (kind === "begin") {
      for (const outer of stack) {
        outer.pieces = true;
      }
      record = {
        pieces: scan.forced?.has(rc) === true,
        separated: false,
        container: undefined,
        pending: ""
      };
      stack.push(record);
    } else {
      record = stack[stack.length - 1];
      if (kind === "end") {
        stack.pop();
      } else if (kind === "separate" && record) {
        record.separated = true;
      }
      if (!record && kind !== "instr") {
        // A separate or end character no field opened: kept where it is.
        record = { pieces: true, separated: true, container: null, pending: "" };
      }
    }
    if (!record) {
      continue;
    }
    if (record.container === undefined) {
      record.container = container;
    } else if (record.container !== container) {
      record.pieces = true;
    }
    if (kept || !whole) {
      record.pieces = true;
    }
    if (!kept) {
      scan.ownerEls.push(rc);
      scan.ownerRecords.push(record);
    }
    if (!first) {
      first = record;
    } else {
      (more ??= []).push(record);
    }
  }
  if (first) {
    if (other) {
      first.pieces = true;
      for (const record of more ?? []) {
        record.pieces = true;
      }
    }
  } else if (other && stack.length > 0) {
    holdContent(scan, !nonText);
  }
}

/** Follow the runs inside markup kept as it is. */
function scanKept(scan: Scan, el: XmlElement): void {
  for (const c of el.children) {
    if (c.type !== "element") {
      continue;
    }
    const name = localName(c);
    if (name === "r") {
      scanRun(scan, c, el, false, true);
    } else if (name === "sdtContent" || OPAQUE.has(name) || TRANSPARENT.has(name)) {
      scanKept(scan, c);
    }
  }
}

function scanContainer(scan: Scan, container: XmlElement, whole: boolean): void {
  for (const child of container.children) {
    if (child.type !== "element") {
      continue;
    }
    if (child.name === "w:r" || hasLocal(child.name, "r")) {
      scanRun(scan, child, container, whole, false);
      continue;
    }
    if (child.name === "w:pPr") {
      continue;
    }
    const name = localName(child);
    if (WRAPPER_PROPERTIES.has(name) || IGNORED.has(name) || isRangeMarkerElement(child)) {
      continue;
    }
    if (scan.stack.length > 0) {
      holdContent(scan, false);
    }
    if (TRANSPARENT.has(name)) {
      scanContainer(scan, child, WHOLE.has(name));
    } else if (OPAQUE.has(name)) {
      scanKept(scan, child);
    }
  }
}

/**
 * `paragraph` with the field characters of every field read in pieces
 * replaced by marker runs — in the paragraph and in the wrappers whose
 * content the reader reads in place — and `open` — the fields read in pieces
 * still open — updated past it. The fields `forced` begins are read in pieces
 * whatever their shape. Returns `paragraph` itself when it holds no piece.
 */
export function splitFieldPieces(
  paragraph: XmlElement,
  open: OpenFieldPiece[],
  forced?: ReadonlySet<XmlElement>
): XmlElement {
  // Pass 1: which fields are read in pieces. A field is when it spans a
  // paragraph or wrapper boundary, lies in a revision, nests another field,
  // shares a run with content, or holds content its whole reading would
  // drop — anything in its instruction, or anything but text in its result.
  // One scan state serves every paragraph (the pass does not re-enter):
  // nearly every paragraph read comes through here.
  const scan = SCAN;
  scan.stack.length = 0;
  scan.ownerEls.length = 0;
  scan.ownerRecords.length = 0;
  scan.forced = forced;
  for (const state of open) {
    const record = state as FieldRecord;
    record.pieces = true;
    record.separated = true;
    record.container = null;
    scan.stack.push(record);
  }
  scanContainer(scan, paragraph, true);
  const stack = scan.stack;
  for (const record of stack) {
    record.pieces = true;
  }
  if (open.length > 0 || stack.length > 0) {
    open.length = 0;
    for (const record of stack) {
      open.push(record);
    }
  }
  let anyPieces = false;
  for (const record of scan.ownerRecords) {
    if (record.pieces) {
      anyPieces = true;
      break;
    }
  }
  scan.forced = undefined;
  if (!anyPieces) {
    return paragraph;
  }
  const owner = new Map<XmlElement, FieldRecord>();
  scan.ownerEls.forEach((el, i) => owner.set(el, scan.ownerRecords[i]!));

  // Pass 2: replace the pieces' field characters with marker runs. An
  // instruction directly after its begin is written with the begin; one
  // directly before the field's next character, with that character; any
  // other — with a nested field, or other content, before the next
  // character — stays where it is, as a marker run of its own.
  let lastBegin: { record: FieldRecord; piece: FieldPiece } | undefined;
  /** Instruction text met and not yet placed, and the list it was met in. */
  let deferred:
    | {
        record: FieldRecord;
        text: string;
        run: XmlElement;
        rPr: XmlElement | undefined;
        out: XmlNode[];
      }
    | undefined;
  /** Place deferred instruction text before what comes next, unless that is its field's next piece. */
  const settle = (next?: FieldRecord): void => {
    if (!deferred) {
      return;
    }
    if (deferred.record === next) {
      next.pending += deferred.text;
    } else {
      const { run, rPr, text } = deferred;
      const el: XmlElement = { ...run, children: rPr ? [rPr] : [] };
      codeTexts.set(el, text);
      deferred.out.push(el);
    }
    deferred = undefined;
  };
  /** Whether `el` — a run, or a wrapper read in place — holds a piece's character. */
  const holdsPiece = (el: XmlElement): boolean =>
    el.children.some(
      c =>
        c.type === "element" &&
        (owner.get(c)?.pieces === true ||
          ((localName(c) === "r" || TRANSPARENT.has(localName(c))) && holdsPiece(c)))
    );
  const splitRun = (child: XmlElement, out: XmlNode[]): void => {
    const rPr = child.children.find(
      (rc): rc is XmlElement => rc.type === "element" && localName(rc) === "rPr"
    );
    let residual: XmlNode[] = [];
    const flush = (): void => {
      if (residual.some(n => n.type === "element")) {
        settle();
        out.push({ ...child, children: rPr ? [rPr, ...residual] : residual });
        lastBegin = undefined;
      }
      residual = [];
    };
    for (const rc of child.children) {
      const record = rc.type === "element" ? owner.get(rc) : undefined;
      if (rc.type !== "element" || rc === rPr || !record?.pieces) {
        if (rc !== rPr) {
          residual.push(rc);
        }
        continue;
      }
      const kind = fieldEvent(rc)!;
      if (kind === "instr") {
        if (lastBegin?.record === record && !residual.some(n => n.type === "element")) {
          lastBegin.piece.instruction += textOf(rc);
        } else {
          flush();
          if (deferred && deferred.record !== record) {
            settle();
          }
          if (deferred) {
            deferred.text += textOf(rc);
          } else {
            deferred = { record, text: textOf(rc), run: child, rPr, out };
          }
          lastBegin = undefined;
        }
        continue;
      }
      flush();
      settle(record);
      const piece: FieldPiece = {
        fieldChar: kind,
        instruction: kind === "begin" ? "" : record.pending,
        fldChar: rc,
        rPr
      };
      record.pending = "";
      const el: XmlElement = { ...child, children: rPr ? [rPr] : [] };
      pieces.set(el, piece);
      out.push(el);
      lastBegin = kind === "begin" ? { record, piece } : undefined;
    }
    flush();
  };
  const rebuild = (container: XmlElement): XmlNode[] => {
    const out: XmlNode[] = [];
    for (const child of container.children) {
      if (child.type !== "element") {
        out.push(child);
        continue;
      }
      if (localName(child) === "r" && holdsPiece(child)) {
        splitRun(child, out);
        continue;
      }
      if (TRANSPARENT.has(localName(child)) && holdsPiece(child)) {
        settle();
        lastBegin = undefined;
        out.push({ ...child, children: rebuild(child) });
      } else {
        // Markup dropped or kept as an anchor does not part instruction text
        // from its field's next piece.
        if (!IGNORED.has(localName(child)) && !isRangeMarkerElement(child)) {
          settle();
        }
        out.push(child);
      }
      lastBegin = undefined;
    }
    if (deferred?.out === out) {
      settle();
    }
    return out;
  };
  return { ...paragraph, children: rebuild(paragraph) };
}

function element(
  name: string,
  attributes: Record<string, string>,
  children: XmlNode[] = []
): XmlElement {
  return { type: "element", name, attributes, children };
}

function fieldCharRun(type: "begin" | "separate" | "end"): XmlElement {
  return element("w:r", {}, [element("w:fldChar", { "w:fldCharType": type })]);
}

/**
 * `children` with each block-level `w:fldSimple` — a simple field around
 * whole paragraphs, which the schema has no place for — replaced by its
 * content, the field written as the complex field it stands for: its begin,
 * instruction and separate characters open its first paragraph, its end
 * closes its last. Returns `children` itself when it holds none.
 */
export function expandBlockSimpleFields(children: XmlNode[]): XmlNode[] {
  if (!children.some(c => c.type === "element" && localName(c) === "fldSimple")) {
    return children;
  }
  const out: XmlNode[] = [];
  for (const child of children) {
    if (child.type !== "element" || localName(child) !== "fldSimple") {
      out.push(child);
      continue;
    }
    const content = expandBlockSimpleFields(child.children);
    const paragraphs = content.filter(
      (c): c is XmlElement => c.type === "element" && localName(c) === "p"
    );
    if (paragraphs.length === 0) {
      const empty = element("w:p", {});
      content.push(empty);
      paragraphs.push(empty);
    }
    const instruction = child.attributes["w:instr"] ?? child.attributes["instr"] ?? "";
    const opening = [
      fieldCharRun("begin"),
      element("w:r", {}, [
        element("w:instrText", { "xml:space": "preserve" }, [{ type: "text", value: instruction }])
      ]),
      fieldCharRun("separate")
    ];
    const first = paragraphs[0]!;
    const last = paragraphs[paragraphs.length - 1]!;
    const replaced = new Map<XmlElement, XmlElement>();
    const withOpening = (p: XmlElement): XmlElement => {
      const at = p.children.findIndex(c => c.type === "element" && localName(c) === "pPr") + 1;
      return { ...p, children: [...p.children.slice(0, at), ...opening, ...p.children.slice(at)] };
    };
    const opened = withOpening(first);
    replaced.set(first, opened);
    const closed = last === first ? opened : last;
    replaced.set(last, { ...closed, children: [...closed.children, fieldCharRun("end")] });
    for (const c of content) {
      out.push(c.type === "element" ? (replaced.get(c) ?? c) : c);
    }
  }
  return out;
}
