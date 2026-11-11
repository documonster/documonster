/**
 * Records: elements read as objects while the document streams past, and a file read in byte
 * ranges that agree with reading it whole.
 *
 * `RecordScanner` sits on a byte-fed {@link SaxParser} and reports each record — an element whose
 * name is asked for and that is not inside another — with its byte range, building its object only
 * when asked to. `Xml.readRecords` runs one per chunk, and checks the chunk boundaries it reads past
 * with more of them.
 */

import { throwIfAborted } from "@utils/errors";
import { check } from "@utils/schema";
import {
  checkByteSource,
  checkReadRecordsOptions,
  partitionResultsSchema,
  snapshotPartitionResults,
  snapshotReadRecordsOptions,
  xmlArgument
} from "@xml/arguments";
import { ByteReader } from "@xml/byte-source";
import {
  XmlArgumentError,
  XmlChunkAmbiguityError,
  XmlLimitExceededError,
  XmlParseError,
  XmlTruncatedError
} from "@xml/errors";
import { SaxParser } from "@xml/sax";
import { ObjectBuilder } from "@xml/to-object";
import { isWhitespaceOnly, resolveOptions } from "@xml/to-object-shared";
import type { ResolvedOptions } from "@xml/to-object-shared";
import type {
  XmlByteSource,
  XmlReadRecordsOptions,
  XmlReadRecordsResult,
  XmlPartitionStep,
  XmlRecord,
  XmlSaxOptions,
  XmlSaxTag
} from "@xml/types";

// =============================================================================
// RecordScanner
// =============================================================================

/** Where a record lies: what two readings of one chunk are compared on. */
export interface RecordSpan {
  readonly name: string;
  readonly startByte: number;
  readonly endByte: number;
}

interface OpenRecord {
  readonly name: string;
  readonly startByte: number;
  /** The parser's depth outside the record: the record ends when its depth returns here. */
  readonly depth: number;
  /** Builds its object, when it is to be given out. */
  readonly builder: ObjectBuilder | undefined;
}

interface RecordScannerConfig {
  readonly sax: XmlSaxOptions;
  readonly tags: ReadonlySet<string>;
  /** Conversion options; without them no object is built. */
  readonly convert: ResolvedOptions | undefined;
  readonly rawXml: boolean;
  /** Whether a record starting at this offset is to be built and given out. */
  readonly owns: (startByte: number) => boolean;
  /** Called as each record ends; `record` is set when it was owned. */
  readonly onRecord: (span: RecordSpan, record: XmlRecord | undefined) => void;
  /**
   * Called with every tag while {@link RecordScanner.reportTags} is set: `<name` as it starts,
   * `/name` as it ends — an element opened before the start included.
   */
  readonly onTag?: (span: RecordSpan) => void;
  /** Accounts for the structure of the chunk, for the reading that reports it. */
  readonly structure?: StructureRecorder;
}

/**
 * The structure of a chunk — what `Xml.verifyPartition` needs to join it to the chunks around it —
 * gathered from the events whose markup begins in the chunk: in order, the end tags of elements
 * opened before it, and what lies outside any element it opened — a run of elements counts as two at
 * most, a run of text as one, so what it holds grows with the nesting it crosses, not with the
 * records it reads — and, taken once at its end, the elements it opened and left open.
 *
 * The elements open are the parser's own stack, taken when the reading passes the end of the chunk,
 * so a tag costs the account a test or two rather than a stack of its own: whether it is the last of
 * the chunk, and whether it leaves or returns to the chunk's top level.
 */
export class StructureRecorder {
  readonly steps: XmlPartitionStep[] = [];
  /** The elements the chunk opened and left open, outermost first: set once the chunk is read. */
  open: readonly string[] | undefined;
  /**
   * Whether the end of the chunk fell inside markup: the next tag tells whether it was the chunk's —
   * it began before the end — or the first past it.
   */
  awaiting = false;
  /**
   * Whether non-white-space text begins past the end of the chunk, before the first markup there: the
   * byte where it does, `null` when none does, `undefined` until the reading has read that far. The
   * next chunk, read from its start, may begin inside markup and read the end of it as text — an `l`
   * step — and this says what really lies there.
   */
  follow: number | null | undefined = undefined;
  /** Index in `steps` where the run since the last `/` step begins. */
  private run = 0;

  constructor(
    readonly end: number,
    /** Where a reading from inside the file starts; `-1` for one from the start. */
    private readonly start: number
  ) {}

  /** An event past the chunk's end is markup: nothing but white space can come before it there. */
  markupPast(): void {
    if (this.follow === undefined) {
      this.follow = null;
    }
  }

  /** An end tag closing what the chunk did not open. */
  outerClose(name: string, byte: number): void {
    this.steps.push(["/", byte, name]);
    this.run = this.steps.length;
  }

  /**
   * Whether the reading, from a start inside the file, met no markup in the whole chunk: it took all
   * of it for text, and what it reads past the end — its `follow` — is no surer than that.
   */
  textOnly = false;

  /** The reading has read every byte of the chunk: take its open elements, or wait for the markup it is inside. */
  reachEnd(parser: SaxParser): void {
    this.textOnly = parser.readNoMarkup;
    if (this.open === undefined) {
      if (parser.openMarkupByte === undefined) {
        this.open = parser.openTagNames;
      } else {
        this.awaiting = true;
      }
    }
  }

  /**
   * The reading is over: the elements open now. It stopped at the end of the input, inside markup
   * past the chunk — so all text before that markup has been seen — or, read as text throughout, at
   * the chunk's own end.
   */
  finish(parser: SaxParser): void {
    this.open ??= parser.openTagNames;
    if (this.textOnly) {
      // The text is still held, unreported: it is the chunk's, from its start.
      if (!parser.holdsBlankText) {
        this.mark("l", this.start);
      }
      // What lies past the chunk is no surer than the chunk itself: the join asks the chunk before.
      this.follow = undefined;
    } else {
      this.follow ??= null;
    }
  }

  /**
   * A CDATA section outside any element the chunk opened: markup, never the end of markup cut by the
   * chunk's start, and content wherever it is — outside the root, whatever it holds, it is not allowed.
   */
  cdata(byte: number): void {
    if (byte >= this.end) {
      this.markupPast();
    } else {
      this.mark("t", byte);
    }
  }

  text(text: string, byte: number): void {
    if (byte >= this.end) {
      if (this.follow === undefined) {
        this.follow = isWhitespaceOnly(text) ? null : byte;
      }
    } else if (!isWhitespaceOnly(text)) {
      // Text from the very start of a reading inside the file may be the end of markup before it.
      this.mark(byte === this.start ? "l" : "t", byte);
    }
  }

  /**
   * Record an element or text outside any element the chunk opened. Between two `/` steps the
   * elements open around the chunk are the same, and `Xml.verifyPartition` needs no more of a run
   * than this: its `l` — which the join may find was markup, and drop — then its first other entry,
   * its first element, and the first entry after that element: whether a root begins there, and
   * whether anything follows a root. So no run holds more than four.
   */
  mark(kind: "e" | "t" | "l", byte: number): void {
    const steps = this.steps;
    if (kind === "l") {
      steps.push([kind, byte]);
      return;
    }
    let first = -1;
    let element = -1;
    for (let i = this.run; i < steps.length; i++) {
      const step = steps[i][0];
      if (step === "l") {
        continue;
      }
      if (first < 0) {
        first = i;
      }
      if (step === "e") {
        element = i;
        break;
      }
    }
    if (first < 0 || (element < 0 ? kind === "e" : element === steps.length - 1)) {
      steps.push([kind, byte]);
    }
  }
}

/**
 * Reads records off a byte-fed SAX parser. Its handlers are the parser's: an error is thrown out
 * of {@link write}, and text is read only inside a record being built.
 */
export class RecordScanner {
  readonly parser: SaxParser;
  private current: OpenRecord | undefined;
  /**
   * Whether tags go to `onTag`: only while another reading is compared with this one, so a reading
   * with nothing to compare builds no span per tag.
   */
  reportTags: boolean;
  /** Whether text and CDATA are pointed at the structure now: see {@link watchTop}. */
  private watchingTop = false;

  constructor(private readonly config: RecordScannerConfig) {
    const parser = (this.parser = SaxParser.withCheckedOptions(config.sax));
    this.reportTags = config.onTag !== undefined;
    parser.on("error", err => {
      throw err;
    });
    parser.on("opentag", tag => this.openTag(tag));
    parser.on("closetag", tag => this.closeTag(tag));
    const { structure, onTag } = config;
    if (structure !== undefined || onTag !== undefined) {
      parser.onOuterClose = name => {
        const startByte = parser.startByte!;
        if (structure !== undefined) {
          if (structure.open === undefined) {
            if (startByte < structure.end) {
              structure.outerClose(name, startByte);
            }
            if (structure.awaiting) {
              // Nothing is open here: the end tag closed what the chunk did not open.
              structure.open = [];
            }
          }
          if (startByte >= structure.end) {
            structure.markupPast();
          }
        }
        if (this.reportTags) {
          onTag?.({ name: "/" + name, startByte, endByte: parser.endByte! });
        }
      };
    }
    this.watchTop(true);
    if (config.rawXml) {
      parser.keepMarkupText(config.tags);
    }
    // Attributes outside a record are read for nothing; inside one they become the object's keys as written.
    const convert = config.convert;
    if (convert !== undefined && !convert.ignoreAttributes) {
      parser.prefixAttributeNames(convert.attrPrefix);
    }
  }

  write(bytes: Uint8Array): void {
    this.parser.write(bytes);
  }

  close(): void {
    this.parser.close();
  }

  /** Where the record being read began, when one is open. */
  get openRecordStart(): number | undefined {
    return this.current?.startByte;
  }

  /**
   * Where the earliest record that may still end in this reading began: the open record, or the
   * markup being read — any markup, not only what is plainly a start tag: a record can begin only
   * where markup does, and taking all of it costs at most a little longer reading, where leaving a
   * kind out let a record go unseen.
   */
  get pendingStart(): number | undefined {
    return this.current?.startByte ?? this.parser.openMarkupByte;
  }

  /** The name of the record being read, when one is open. */
  get openRecordName(): string | undefined {
    return this.current?.name;
  }

  /**
   * The parser's lexical state with the record being read: two scanners that agree on it read the
   * same records from here on.
   */
  getLexicalState(): string {
    const parserState = this.parser.getLexicalState();
    const current = this.current;
    return current === undefined
      ? parserState
      : `${parserState}|${current.name}@${current.startByte}/${this.parser.depth - current.depth}`;
  }

  private readonly onText = (text: string): void => {
    this.current!.builder!.text(text);
  };

  private readonly onCData = (text: string): void => {
    this.current!.builder!.cdata(text);
  };

  /** Text outside any element the chunk opened is the structure's to see, when no record reads it. */
  private readonly onTopText = (text: string): void => {
    this.config.structure!.text(text, this.parser.startByte!);
  };

  private readonly onTopCData = (): void => {
    this.config.structure!.cdata(this.parser.startByte!);
  };

  /**
   * Point text and CDATA at the structure while the reading is at the chunk's top level — outside any
   * element it opened — and no record reads them. Switched only as the reading leaves or returns
   * there, not at every tag.
   */
  private watchTop(top: boolean): void {
    if (
      this.config.structure !== undefined &&
      this.current === undefined &&
      top !== this.watchingTop
    ) {
      this.watchingTop = top;
      this.parser.setHandler("text", top ? this.onTopText : undefined);
      this.parser.setHandler("cdata", top ? this.onTopCData : undefined);
    }
  }

  private openTag(tag: XmlSaxTag): void {
    const { structure, onTag } = this.config;
    if (this.reportTags) {
      onTag!({
        name: "<" + tag.name,
        startByte: this.parser.startByte!,
        endByte: this.parser.endByte!
      });
    }
    if (structure !== undefined && structure.open !== undefined) {
      // Past the chunk: the first tag there is markup, and nothing but white space came before it.
      if (structure.follow === undefined) {
        structure.markupPast();
      }
    } else if (structure !== undefined) {
      const parser = this.parser;
      const top = parser.depth === 0;
      if (structure.awaiting) {
        // The end of the chunk fell inside markup: this tag is the chunk's last, or the first past it.
        if (parser.startByte! >= structure.end) {
          structure.markupPast();
          structure.open = parser.openTagNames;
        } else {
          if (top && tag.isSelfClosing) {
            structure.mark("e", parser.startByte!);
          }
          structure.open = tag.isSelfClosing
            ? parser.openTagNames
            : [...parser.openTagNames, tag.name];
        }
      } else if (top) {
        if (tag.isSelfClosing) {
          structure.mark("e", parser.startByte!);
        } else {
          this.watchTop(false);
        }
      }
    }
    const current = this.current;
    if (current !== undefined) {
      current.builder?.openTag(tag);
      return;
    }
    if (!this.config.tags.has(tag.name)) {
      return;
    }
    const { config, parser } = this;
    const startByte = parser.startByte!;
    const convert = config.owns(startByte) ? config.convert : undefined;
    const builder =
      convert === undefined ? undefined : new ObjectBuilder(convert, !convert.ignoreAttributes);
    builder?.openTag(tag);
    if (builder !== undefined && config.rawXml) {
      parser.beginCapture();
    }
    if (tag.isSelfClosing) {
      this.emit(tag.name, startByte, builder);
      return;
    }
    this.current = { name: tag.name, startByte, depth: parser.depth, builder };
    if (builder !== undefined) {
      parser.setHandler("text", this.onText);
      parser.setHandler("cdata", this.onCData);
    }
  }

  private closeTag(tag: XmlSaxTag): void {
    const { structure, onTag } = this.config;
    if (this.reportTags) {
      onTag!({
        name: "/" + tag.name,
        startByte: this.parser.startByte!,
        endByte: this.parser.endByte!
      });
    }
    // The parser has taken the element off its stack already: its depth is the one after it.
    const top = this.parser.depth === 0;
    if (structure !== undefined && structure.open !== undefined) {
      // A self-closing tag's end is its start: one that began in the chunk is no markup past it.
      if (structure.follow === undefined && this.parser.startByte! >= structure.end) {
        structure.markupPast();
      }
    } else if (structure !== undefined && !tag.isSelfClosing) {
      const parser = this.parser;
      if (structure.awaiting) {
        if (parser.startByte! >= structure.end) {
          structure.markupPast();
          structure.open = [...parser.openTagNames, tag.name];
        } else {
          if (top) {
            structure.mark("e", parser.startByte!);
          }
          structure.open = parser.openTagNames;
        }
      } else if (top) {
        structure.mark("e", parser.startByte!);
      }
    }
    this.closeRecord(tag);
    if (top && !tag.isSelfClosing) {
      this.watchTop(true);
    }
  }

  private closeRecord(tag: XmlSaxTag): void {
    const current = this.current;
    if (current === undefined || tag.isSelfClosing) {
      return;
    }
    current.builder?.closeTag(tag);
    if (this.parser.depth === current.depth) {
      this.current = undefined;
      if (current.builder !== undefined) {
        this.parser.setHandler("text", undefined);
        this.parser.setHandler("cdata", undefined);
      }
      this.emit(current.name, current.startByte, current.builder);
    }
  }

  private emit(name: string, startByte: number, builder: ObjectBuilder | undefined): void {
    const span: RecordSpan = { name, startByte, endByte: this.parser.endByte! };
    let record: XmlRecord | undefined;
    if (builder !== undefined) {
      record = this.config.rawXml
        ? { ...span, object: builder.result, xml: this.parser.endCapture() }
        : { ...span, object: builder.result };
    }
    this.config.onRecord(span, record);
  }
}

// =============================================================================
// Xml.readRecords
// =============================================================================

/**
 * How much is fed at once while a boundary is being placed: a reading of the next chunk is checked
 * against this one at the end of each such slice, so it is how far past agreement either reads.
 */
const SLICE_BYTES = 1024;

/**
 * A second reading of a later chunk, begun at its start as that chunk's own read begins, and
 * checked against the reading that came through the boundary.
 */
interface ChunkCheck {
  /** The chunk: `[start, end)`. */
  readonly start: number;
  readonly end: number;
  readonly scanner: RecordScanner;
  /** Records of the chunk as the reading through the boundary found them. */
  readonly ours: RecordSpan[];
  /** Records of the chunk as the reading from its start found them. */
  readonly theirs: RecordSpan[];
  /** How many of each have been compared. */
  compared: number;
  /** Whether the reading from its start can add nothing more: it has left the chunk with no record open in it. */
  finished: boolean;
}

/** What a record's callback threw, carried out of the parser it ran inside. */
class CallbackFailure {
  constructor(readonly error: unknown) {}
}

/**
 * What `value` settles to, as `await` would take it — or `undefined` when it is no thenable, and a
 * synchronous callback costs no wait. A promise of this realm is taken as it is. Of anything else,
 * `then` is read once and that method is called, in a microtask as the language calls it: reading it
 * again could find another, and a thenable then went unwaited for.
 */
function awaitable(value: unknown): Promise<unknown> | undefined {
  if (value instanceof Promise && Object.getPrototypeOf(value) === Promise.prototype) {
    return value;
  }
  if ((typeof value !== "object" && typeof value !== "function") || value === null) {
    return undefined;
  }
  const then: unknown = (value as { then?: unknown }).then;
  if (typeof then !== "function") {
    return undefined;
  }
  return new Promise((resolve, reject) => {
    queueMicrotask(() => {
      try {
        then.call(value, resolve, reject);
      } catch (err) {
        reject(err);
      }
    });
  });
}

function sameSpan(a: RecordSpan, b: RecordSpan): boolean {
  return a.name === b.name && a.startByte === b.startByte && a.endByte === b.endByte;
}

function ambiguity(
  check: ChunkCheck,
  message: string,
  atByte: number,
  cause?: unknown
): XmlChunkAmbiguityError {
  return new XmlChunkAmbiguityError(
    `Xml.readRecords: ${message}`,
    check.start,
    atByte,
    cause === undefined ? undefined : { cause }
  );
}

/** Compare what both readings of `check`'s chunk have found, as far as both have got. */
function compare(check: ChunkCheck): void {
  const { ours, theirs } = check;
  const both = Math.min(ours.length, theirs.length);
  for (; check.compared < both; check.compared++) {
    const a = ours[check.compared];
    const b = theirs[check.compared];
    if (!sameSpan(a, b)) {
      throw ambiguity(
        check,
        `read from its start, the chunk has <${b.name}> at byte ${b.startByte}–${b.endByte}; read through from before it, <${a.name}> at byte ${a.startByte}–${a.endByte}`,
        Math.min(a.startByte, b.startByte)
      );
    }
  }
}

/** Neither reading will find another record of the chunk: they must have found the same ones. */
function settle(check: ChunkCheck): void {
  compare(check);
  const { ours, theirs } = check;
  if (ours.length !== theirs.length) {
    const extra = ours.length > theirs.length ? ours[theirs.length] : theirs[ours.length];
    const reading = ours.length > theirs.length ? "through from before it" : "from its start";
    throw ambiguity(
      check,
      `only the reading ${reading} finds <${extra.name}> at byte ${extra.startByte}`,
      extra.startByte
    );
  }
}

/**
 * Read the records of an XML file, or of one byte range of it, as objects — one pass over the
 * bytes, building each record's object as `Xml.parseToObject` would from its text, and holding no
 * more than one record at a time.
 *
 * A record is an element named in `tags` that is not inside another record. Each is given to
 * `onRecord` in file order, with its byte range; a returned promise is awaited first.
 *
 * **Chunks.** With `offset` and `size`, only the records whose start tag begins in
 * `[offset, offset + size)` are read, so a file can be split into ranges read independently — by
 * separate workers, say — and the records of all of them, in order, are those of the whole file.
 * `source` gives the bytes from `offset` on, and is read as far past the range as its last record
 * runs, and a little further. A range that starts after 0 is read without its prolog: pass the
 * file's `context` (`Xml.readRoot(...).context`).
 *
 * A range read from its own start must begin in the content of some element, and is read from
 * there as if it did. That assumption is checked rather than trusted: where the read of one range
 * continues past the start of the next, it reads the next one from its start as well, and the two
 * readings must find the same records until they reach the same state. If they cannot, the call
 * throws {@link XmlChunkAmbiguityError} — the file has records nested in records of the same name,
 * or tags inside comments or CDATA sections, around the boundary. The partition then cannot be
 * trusted; read that range together with the one before it.
 *
 * **What a partition promises.** If every range of a partition of one version of the file reads
 * without an error, and `Xml.verifyPartition` accepts their results, their records, in order, are
 * exactly the whole file's. No range alone promises that: a range stops once its own records are
 * read, and none sees the elements open around it — a file cut short, with a second root, or with an
 * end tag that closes the wrong element can let every range succeed. Each result carries its range's
 * `structure` for the join. Stage each range's records, and publish them only when every range has
 * succeeded and the join has passed; if anything fails, discard the whole batch — the ranges that
 * succeeded included — and read again (the failed range with the one before it, or the file whole).
 * `onRecord` is called as records are read, so a call that fails may already have given some.
 *
 * Errors are thrown, never recovered from. Input that ends inside a record the call was asked for
 * is an {@link XmlTruncatedError} carrying `recordStartByte`, and a range read to the end of the
 * input checks that the document ends there as one does. The range from offset 0 reads the prolog
 * through to the root's start tag, wherever the range ends: only a reading from the start can check
 * it.
 *
 * DOCTYPE internal subsets are read: their general entities expand, and `context` carries them to
 * later ranges. External entities and DTDs are never read — a reference to one is an undefined
 * entity — and nothing is fetched. `maxDepth`, `maxEntityExpansions` and `maxEntityExpansionLength` bound what one call spends,
 * so they count within it: from a range's start, nesting outside it is not seen.
 *
 * @example
 * ```ts
 * const { context } = await Xml.readRoot(await readBytes(path, 0, 65536));
 * const { records } = await Xml.readRecords(readStream(path, { start: offset }), {
 *   tags: ["Track"],
 *   offset,
 *   size: 64 * 1024 * 1024,
 *   context,
 *   onRecord: record => tracks.push(record.object)
 * });
 * ```
 */
async function readRecords(
  source: XmlByteSource,
  given: XmlReadRecordsOptions
): Promise<XmlReadRecordsResult> {
  const where = "Xml.readRecords";
  const checkedSource = checkByteSource(where, source);
  // Read once: what is checked is what every parser of the call is built from, and a getter on the
  // caller's object cannot hand a later one something else.
  const options = snapshotReadRecordsOptions(given);
  checkReadRecordsOptions(where, options);
  const resolved = resolveOptions(options, where);
  const isArray = resolved.isArray;
  // `isArray` is the caller's code run inside the parser, as `onRecord` is: what it throws is theirs.
  const convert: ResolvedOptions =
    isArray === null
      ? resolved
      : {
          ...resolved,
          isArray: name => {
            try {
              const result = isArray(name);
              // A signal it fired stops the read here, not at the end of the record.
              throwIfAborted(options.signal);
              return result;
            } catch (err) {
              throw new CallbackFailure(err);
            }
          }
        };
  const offset = options.offset ?? 0;
  const size = options.size ?? Infinity;
  const end = offset + size;
  const resume = offset > 0;
  const tags = new Set(options.tags);
  const onRecord = options.onRecord;
  const rawXml = options.rawXml ?? false;
  /**
   * The context a reading that starts later needs: the caller's, or — reading from the start — the
   * one the main reading has just read, taken once it has read the root's start tag.
   */
  let laterContext = options.context;
  const sax = (byteOffset: number, resumed: boolean): XmlSaxOptions => ({
    position: false,
    byteOffset,
    resume: resumed,
    internalEntities: true,
    context: resumed ? laterContext : undefined,
    maxDepth: options.maxDepth,
    maxEntityExpansions: options.maxEntityExpansions,
    maxEntityExpansionLength: options.maxEntityExpansionLength,
    invalidCharHandling: options.invalidCharHandling,
    maxContextBytes: options.maxContextBytes,
    // Only a read from the file's start has a document start to look for.
    skipLeadingText: !resumed && options.skipLeadingText
  });

  /** Checks of later chunks still running, earliest first. */
  const checks: ChunkCheck[] = [];
  /** Whether the last check to end ended by the two readings agreeing. */
  let agreed = false;
  /** How many records have been given to `onRecord` (and, for a promise, it has settled). */
  let delivered = 0;
  /** What `onRecord` returned for the last record, while the main reading waits for it. */
  let waitingFor: Promise<unknown> | undefined;
  const signal = options.signal;

  // Read whole, from the start to the end, the parser checks the document itself: there is nothing to
  // join, and no tag is worth accounting for.
  const structure =
    resume || size !== Infinity ? new StructureRecorder(end, resume ? offset : -1) : undefined;
  const main = new RecordScanner({
    sax: sax(offset, resume),
    tags,
    convert,
    rawXml,
    structure,
    // While a later chunk is read from its start as well, every tag is compared, not only records:
    // the structure that chunk reports is built from the tags it reads, so they must be the file's.
    onTag: span => {
      for (const check of checks) {
        if (span.startByte >= check.start && span.startByte < check.end) {
          check.ours.push(span);
          compare(check);
        }
      }
    },
    owns: start => start < end,
    onRecord: (span, record) => {
      for (const check of checks) {
        if (span.startByte >= check.start && span.startByte < check.end) {
          check.ours.push(span);
          compare(check);
        }
      }
      if (record === undefined) {
        return;
      }
      if (signal?.aborted) {
        // Fired by the caller's code since the last check — `isArray`, say: nothing is given out
        // after it; the pause ends the reading with an `AbortError`.
        main.parser.pause();
        return;
      }
      // Given out as it ends, so no more than one record is held. A promise pauses the reading here
      // until it settles; a signal fired from inside the callback stops it here too.
      let pending: Promise<unknown> | undefined;
      try {
        // Even looking at what it returned runs the caller's code: a `then` getter may throw.
        pending = awaitable(onRecord(record));
      } catch (err) {
        // The callback runs inside the parser: mark what it throws, so it passes out as it is
        // rather than being taken for a parse error.
        throw new CallbackFailure(err);
      }
      if (pending !== undefined) {
        // Waited for below — unless an abort comes first, and then its failure is nobody's to report.
        pending.catch(() => undefined);
        waitingFor = pending;
        main.parser.pause();
      } else {
        delivered++;
        if (signal?.aborted) {
          main.parser.pause();
        }
      }
    }
  });

  const startCheck = (start: number): void => {
    laterContext ??= main.parser.getPrologContext();
    const end = start + size;
    const ours: RecordSpan[] = [];
    const theirs: RecordSpan[] = [];
    const scanner = new RecordScanner({
      sax: sax(start, true),
      tags,
      convert: undefined,
      rawXml: false,
      owns: () => false,
      onRecord: span => {
        if (span.startByte < end) {
          theirs.push(span);
          compare(check);
        }
      },
      onTag: span => {
        if (span.startByte < end) {
          theirs.push(span);
          compare(check);
        }
      }
    });
    const check: ChunkCheck = { start, end, scanner, ours, theirs, compared: 0, finished: false };
    if (scanner.getLexicalState() === main.getLexicalState()) {
      agreed = true;
    } else {
      checks.push(check);
      main.reportTags = true;
    }
  };

  /** Settle what checks can be settled at `pos`, the main reading's position. */
  const advanceChecks = (pos: number): void => {
    const mainState = main.getLexicalState();
    const mainOpen = main.pendingStart;
    for (let i = 0; i < checks.length; i++) {
      const check = checks[i];
      let done = false;
      const theirOpen = check.scanner.openRecordStart;
      if (
        !check.finished &&
        theirOpen !== undefined &&
        theirOpen >= check.start &&
        theirOpen < check.end &&
        main.openRecordStart !== theirOpen
      ) {
        // Both readings have read the start tag: the reading through the boundary took it as no
        // record's start, and cannot take it as one later, while the reading from the chunk's start
        // has opened a record there. Waiting for that record to end — as far as the end of the file,
        // and once for every boundary it spans — would only find the same difference.
        throw ambiguity(
          check,
          `read from its start, the chunk opens <${check.scanner.openRecordName}> at byte ${theirOpen} as a record; read through from before it, that tag begins none`,
          theirOpen
        );
      }
      if (!check.finished && check.scanner.getLexicalState() === mainState) {
        // From here both readings are the same reading: what each found so far must match.
        settle(check);
        done = true;
        agreed = true;
      } else {
        if (!check.finished && pos >= check.end) {
          const open = check.scanner.pendingStart;
          check.finished = open === undefined || open >= check.end;
        }
        if (
          check.finished &&
          (mainOpen === undefined || mainOpen < check.start || mainOpen >= check.end)
        ) {
          settle(check);
          done = true;
          agreed = false;
        }
      }
      if (done) {
        checks.splice(i, 1);
        i--;
      }
    }
    main.reportTags = checks.length > 0;
  };

  main.reportTags = false;

  const fail = (err: unknown, check: ChunkCheck | undefined): never => {
    if (err instanceof CallbackFailure) {
      throw err.error;
    }
    if (check !== undefined) {
      if (err instanceof XmlParseError) {
        throw ambiguity(
          check,
          `read from its start, the chunk fails: ${err.message}`,
          err.byte ?? check.start,
          err
        );
      }
      throw err;
    }
    if (
      resume &&
      err instanceof XmlParseError &&
      !(err instanceof XmlTruncatedError) &&
      !(err instanceof XmlLimitExceededError)
    ) {
      // From a start inside the file, an error can mean the start was misplaced as well as a malformed file.
      throw new XmlChunkAmbiguityError(
        `${where}: the chunk cannot be read from its start: ${err.message}`,
        offset,
        err.byte ?? offset,
        { cause: err }
      );
    }
    throw err;
  };

  const reader = new ByteReader(checkedSource, where, signal);
  let pos = offset;
  /** Run a step of the main reading — `write` a piece, or `close` it. */
  const stepMain = (step: () => void): void => {
    try {
      step();
    } catch (err) {
      fail(err, undefined);
    }
  };
  /**
   * Go on through every pause a record's callback made: wait for what it returned, then go on from
   * the record's end. Only a paused reading waits, so a synchronous callback costs no `await`.
   */
  const drainPauses = async (): Promise<void> => {
    while (main.parser.paused) {
      reader.checkAborted();
      const pending = waitingFor;
      if (pending !== undefined) {
        waitingFor = undefined;
        await reader.abortable(() => pending);
        delivered++;
        // The signal may have fired while it was pending: no record is given out after it has.
        reader.checkAborted();
      }
      stepMain(() => main.parser.resumeParsing());
    }
  };
  /** The next chunk start after `at`, or `Infinity`. */
  const nextBoundary = (at: number): number =>
    size === Infinity ? Infinity : offset + (Math.floor((at - offset) / size) + 1) * size;
  const canStop = (): boolean => {
    if (pos < end || checks.length > 0) {
      return false;
    }
    // Read from a start inside the file, the whole chunk was text: it holds no `<`, so the reading of
    // an earlier chunk — the last one that met markup — reads through all of it to place its own end,
    // and checks the next chunk's start on the way. Reading on would only read that again.
    if (resume && main.parser.readNoMarkup) {
      return true;
    }
    if (!agreed) {
      return false;
    }
    // From the start of the file, the prolog is read through to the root's start tag, wherever the
    // chunk ends: only a reading from the start can check it, and a later chunk that starts inside
    // it reads it as content.
    if (!resume && !main.parser.sawRootStart) {
      return false;
    }
    // Stopped only inside markup that begins past the chunk: the text before it, which the next chunk
    // reads from a start that may be inside markup and so leaves some checks to this reading, has
    // been read through.
    const pending = main.pendingStart;
    return pending !== undefined && pending >= end;
  };

  let boundary = nextBoundary(pos);
  /** Whether the read stopped with bytes of the source's last chunk still unread. */
  let unreadInChunk = false;
  /** Read `chunk`, in pieces that end at each chunk start and, while a check runs, every slice; whether the read is done. */
  const feed = async (chunk: Uint8Array): Promise<boolean> => {
    let at = 0;
    while (at < chunk.length) {
      reader.checkAborted();
      let stop = Math.min(chunk.length, at + (boundary - pos));
      if (checks.length > 0) {
        stop = Math.min(stop, at + SLICE_BYTES);
      }
      const piece = chunk.subarray(at, stop);
      stepMain(() => main.write(piece));
      if (main.parser.paused) {
        await drainPauses();
      }
      for (const check of checks) {
        if (!check.finished) {
          try {
            check.scanner.write(piece);
          } catch (err) {
            fail(err, check);
          }
        }
      }
      pos += piece.length;
      at = stop;
      if (pos === end) {
        structure?.reachEnd(main.parser);
      }
      advanceChecks(pos);
      if (pos === boundary) {
        boundary = nextBoundary(pos);
        if (!canStop()) {
          startCheck(pos);
        }
      }
      if (canStop()) {
        unreadInChunk = at < chunk.length;
        return true;
      }
    }
    return false;
  };

  let succeeded = false;
  try {
    let done = false;
    let atEnd = false;
    while (!done) {
      const chunk = await reader.next();
      // The signal may have fired since the chunk arrived: nothing is read after it has.
      reader.checkAborted();
      if (chunk === undefined) {
        atEnd = true;
        break;
      }
      done = await feed(chunk);
    }
    if (done && !unreadInChunk && main.parser.readNoMarkup && resume) {
      // Stopped at the end of a chunk read as text throughout: whether the file ends there too is one
      // read away, and a chunk that ends the file says so (`final`).
      for (;;) {
        const chunk = await reader.next();
        reader.checkAborted();
        if (chunk === undefined) {
          atEnd = true;
          break;
        }
        if (chunk.length > 0) {
          break;
        }
      }
    }

    if (atEnd) {
      // The end of the file: it must end here as a document does, and the checks still running
      // must agree up to it.
      const open = main.openRecordStart;
      if (open !== undefined && open < end) {
        throw new XmlTruncatedError(`${where}: the input ends inside <${main.openRecordName}>`, {
          byte: pos,
          recordStartByte: open
        });
      }
      // Or inside the start tag of one: its name is read, its `>` is not.
      const tagName = main.parser.openStartTagName;
      const tagStart = main.parser.openMarkupByte;
      if (open === undefined && tagName !== undefined && tags.has(tagName) && tagStart! < end) {
        throw new XmlTruncatedError(
          `${where}: the input ends inside the start tag of <${tagName}>`,
          {
            byte: pos,
            recordStartByte: tagStart
          }
        );
      }
      // Closed first: the input's last text is reported as it closes.
      stepMain(() => main.close());
      await drainPauses();
      for (const check of checks) {
        settle(check);
      }
    }
    structure?.finish(main.parser);
    succeeded = true;
    const skipped = main.parser.skippedLeadingText;
    return {
      records: delivered,
      endByte: pos,
      ...(skipped > 0 && { skippedLeadingBytes: skipped }),
      structure: {
        start: offset,
        end: size === Infinity ? null : end,
        final: atEnd,
        ...(atEnd && { eof: pos }),
        // Read whole, the file is one root element, checked: so much is all a join needs to know.
        steps: structure?.steps ?? [["e", 0]],
        open: structure?.open ?? [],
        ...(structure?.follow !== undefined && { follow: structure.follow }),
        ...(structure?.textOnly === true && { textOnly: true as const })
      }
    };
  } finally {
    await reader.close(!succeeded);
  }
}

/**
 * Check that the chunks of a partition, read with `Xml.readRecords`, join into one whole, well-formed
 * document — the step that makes their records the file's. Each chunk checks what it reads, but none
 * knows the elements open around it: a file cut short, with a second root, or with an end tag that
 * closes the wrong element can let every chunk succeed. This joins their `structure`s in order and
 * finds those, without reading the file again.
 *
 * Pass every chunk's result, in file order, once all have succeeded; publish their records only if
 * this returns. Throws {@link XmlArgumentError} when the results do not cover the file from byte 0 to
 * its end without gaps or overlaps, {@link XmlTruncatedError} when the file ends inside an element,
 * and {@link XmlParseError} for any other way the chunks fail to make one document.
 *
 * @example
 * ```ts
 * const results = await Promise.all(ranges.map(range => readRange(range)));
 * Xml.verifyPartition(results);
 * publish(staged);
 * ```
 */
function verifyPartition(given: readonly Pick<XmlReadRecordsResult, "structure">[]): void {
  const where = "Xml.verifyPartition";
  // Read once, so what is checked is what is joined.
  const results = snapshotPartitionResults(given) as typeof given;
  check(partitionResultsSchema, results, xmlArgument(where, "results"));
  const open: string[] = [];
  let rootSeen = false;
  let expected = 0;
  /** What the last chunk before this one that met markup found past its end: see `follow`. */
  let trustedFollow: number | null | undefined;
  /** Where the file ends, as the chunks that read to it found. */
  let eof: number | undefined;
  for (let i = 0; i < results.length; i++) {
    const structure = results[i].structure;
    const last = i === results.length - 1;
    if (structure.start !== expected) {
      throw new XmlArgumentError(
        "results",
        `${where}: results[${i}] begins at byte ${structure.start}, not ${expected}: pass the results of every chunk of one partition, in file order`
      );
    }
    if (structure.final) {
      // Every chunk that read to the end of the file found it at the same byte.
      if (eof !== undefined && structure.eof !== eof) {
        throw new XmlArgumentError(
          "results",
          `${where}: results[${i}] found the end of the file at byte ${structure.eof}, another at ${eof}: pass the results of one partition of one file`
        );
      }
      eof = structure.eof;
    }
    if (last ? !structure.final : structure.end === null) {
      throw new XmlArgumentError(
        "results",
        last
          ? `${where}: results[${i}], the last, stopped before the end of the file: the partition does not cover it`
          : `${where}: results[${i}] runs to the end of the file, but is not the last`
      );
    }
    for (const step of structure.steps) {
      const byte = step[1];
      if (step[0] === "/") {
        const name = step[2];
        const innermost = open.pop();
        if (innermost !== name) {
          throw new XmlParseError(
            innermost === undefined
              ? `${where}: the end tag </${name}> at byte ${byte} closes no element`
              : `${where}: the end tag </${name}> at byte ${byte} does not close <${innermost}>`,
            { byte }
          );
        }
      } else if (open.length === 0 && !rootSeen && step[0] !== "e") {
        // Before the root element: the prolog, which the chunk from the start of the file read through
        // and checked. A later chunk that starts inside it reads it as content.
        continue;
      } else if (open.length === 0) {
        let kind = step[0];
        let at = byte;
        if (kind === "l") {
          // The chunk read from its start may have begun inside markup: what the chunk before read
          // past its own end says what really lies there.
          // A chunk read all as text read no surer: the one that met markup before it says.
          const follow = trustedFollow;
          if (follow === null) {
            continue;
          }
          if (follow !== undefined) {
            kind = "t";
            at = follow;
          }
        }
        if (kind === "l") {
          throw new XmlChunkAmbiguityError(
            `${where}: the chunk at byte ${structure.start} begins, outside the root element, with text that may be the end of markup begun before it; read it with the chunk before`,
            structure.start,
            byte
          );
        }
        if (kind === "t") {
          throw new XmlParseError(`${where}: text outside the root element at byte ${at}`, {
            byte: at
          });
        }
        if (rootSeen) {
          throw new XmlParseError(`${where}: a second root element at byte ${byte}`, { byte });
        }
        rootSeen = true;
      }
    }
    for (const name of structure.open) {
      if (open.length === 0) {
        if (rootSeen) {
          throw new XmlParseError(
            `${where}: a second root element <${name}> in the chunk at byte ${structure.start}`,
            { byte: structure.start }
          );
        }
        rootSeen = true;
      }
      open.push(name);
    }
    expected = structure.end ?? Infinity;
    if (structure.textOnly !== true) {
      trustedFollow = structure.follow;
    }
  }
  // The chunks cover the file: each begins before its end — a chunk past it reads nothing of it —
  // and the last reaches it.
  const lastStructure = results[results.length - 1].structure;
  if (lastStructure.end !== null && lastStructure.end < eof!) {
    throw new XmlArgumentError(
      "results",
      `${where}: the last chunk ends at byte ${lastStructure.end}, before the end of the file at byte ${eof}: the partition does not cover it`
    );
  }
  for (let i = 0; i < results.length; i++) {
    const start = results[i].structure.start;
    if (start >= eof! && !(start === 0 && eof === 0)) {
      throw new XmlArgumentError(
        "results",
        `${where}: results[${i}] begins at byte ${start}, at or past the end of the file at byte ${eof}`
      );
    }
  }
  if (open.length > 0) {
    throw new XmlTruncatedError(`${where}: the file ends inside <${open[open.length - 1]}>`);
  }
  if (!rootSeen) {
    throw new XmlTruncatedError(`${where}: the file holds no root element`);
  }
}

export { readRecords, verifyPartition };
