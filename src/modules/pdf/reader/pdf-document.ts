/**
 * PDF document parser.
 *
 * Handles the high-level PDF file structure:
 * - Locating startxref
 * - Parsing cross-reference tables (traditional and stream-based)
 * - Reading trailer dictionaries
 * - Resolving indirect object references
 * - Handling incremental updates
 *
 * @see PDF Reference 1.7, §3.4 - File Structure
 */

import { PdfLimitExceededError, PdfStructureError } from "@pdf/errors";
import {
  parseObject,
  isPdfDict,
  isPdfStream,
  isPdfRef,
  isPdfArray,
  dictGetNumber,
  dictGetRef,
  dictGetArray,
  dictGetName,
  decodePdfStringBytes
} from "@pdf/reader/pdf-parser";
import type { PdfObject, PdfDictValue, PdfRef, PdfStream } from "@pdf/reader/pdf-parser";
import { PdfTokenizer, TokenType } from "@pdf/reader/pdf-tokenizer";
import { decodeStreamFilters } from "@pdf/reader/stream-filters";
import { assertLimitOption } from "@utils/limits";

// =============================================================================
// Module-level cached TextEncoder
// =============================================================================

/** Cached TextEncoder instance to avoid repeated allocation in hot paths */
const _encoder = new TextEncoder();

// =============================================================================
// Types
// =============================================================================

/** Cross-reference entry for a single object */
interface XrefEntry {
  /** Byte offset of the object in the file (type 1) or object number of the object stream (type 2) */
  offset: number;
  /** Generation number (type 1) or index within the object stream (type 2) */
  gen: number;
  /** Entry type: 0 = free, 1 = in-use (uncompressed), 2 = in object stream */
  type: number;
}

/** Result of resolving an object with its object/generation numbers for decryption */
interface ResolvedObject {
  /** The resolved PDF object */
  obj: PdfObject | null;
  /** The object number */
  objNum: number;
  /** The generation number */
  gen: number;
}

// =============================================================================
// PDF Document
// =============================================================================

/**
 * Parsed PDF document with lazy object resolution.
 *
 * Reads the cross-reference table and trailer on construction,
 * then resolves individual objects on demand with caching.
 */
/** Which crypt filter applies to a piece of data (ISO 32000-1 §7.6.5). */
export type PdfCryptKind = "string" | "stream";

/** Per-object encryption or decryption of a string or stream. */
export type PdfCryptFn = (
  data: Uint8Array,
  objNum: number,
  gen: number,
  kind: PdfCryptKind
) => Uint8Array;

/** One member of an object stream: its object number and parsed value. */
interface ObjStmMember {
  objNum: number;
  obj: PdfObject | null;
}

export class PdfDocument {
  private tokenizer: PdfTokenizer;
  private xref: Map<number, XrefEntry> = new Map();
  private cache: Map<string, PdfObject> = new Map();
  // Parsed object streams (keyed by the object-stream's object number). Kept
  // separate from `cache` because the value is a Map, not a `PdfObject`.
  // Each entry is the stream's members in index order; a member that failed
  // to parse keeps its slot (as `null`) so later indices stay aligned.
  private objStmCache: Map<number, ObjStmMember[]> = new Map();
  declare readonly trailer: PdfDictValue;

  /**
   * Encryption handler (set externally after decryption is initialized).
   * `kind` selects the crypt filter: `/StrF` for strings, `/StmF` for streams
   * (ISO 32000-1 §7.6.5); an `/Identity` filter returns the data unchanged.
   */
  decryptFn: PdfCryptFn | null = null;
  /**
   * Inverse of {@link decryptFn}, set alongside it: encrypts a string or stream
   * belonging to object `objNum gen` with the document's own security handler,
   * so an incremental update can append objects the original /Encrypt covers.
   */
  encryptFn: PdfCryptFn | null = null;
  /**
   * `false` when the file's `/EncryptMetadata` is false: metadata streams
   * (`/Type /Metadata`) are then stored in the clear and must be neither
   * decrypted nor encrypted.
   */
  encryptMetadata = true;

  /** Most bytes one stream filter may produce; `undefined` for the default. */
  private readonly maxDecodedBytes: number | undefined;

  /**
   * The first resource limit (nesting depth, decoded size, filter chain) hit
   * while reading, if any. Many readers deliberately tolerate a malformed
   * object and carry on, which would turn a limit hit into a silently
   * truncated result; recording it here lets each entry point refuse the
   * whole read once, instead of every tolerant `catch` having to rethrow it.
   *
   * Once set, the document is poisoned: every later parse, decode or object
   * resolution throws it immediately, so a tolerant loop over pages or objects
   * fails at its next operation instead of doing more work on a hostile file.
   */
  limitError: PdfLimitExceededError | null = null;

  /**
   * `true` when the cross-reference data could not be read from the
   * `startxref` chain and was rebuilt by scanning the file for `N G obj`
   * headers. The object map is then a best guess and the file's own
   * `startxref` names a broken section, so nothing may be chained onto it
   * with `/Prev` — an incremental update must become a full rewrite.
   */
  xrefRecovered = false;

  /**
   * The offset the file's last `startxref` names — the section an
   * incremental update's `/Prev` must point at. `null` when it could not be
   * read (see {@link xrefRecovered}).
   */
  startxrefOffset: number | null = null;

  constructor(data: Uint8Array, maxDecodedBytes?: number) {
    assertLimitOption("maxDecodedBytes", maxDecodedBytes);
    this.maxDecodedBytes = maxDecodedBytes;
    this.tokenizer = new PdfTokenizer(data);
    try {
      this.trailer = this.parseFileStructure();
    } catch (err) {
      // A limit hit swallowed during recovery outranks the error it led to.
      throw this.limitError ?? err;
    }
    this.throwIfLimitExceeded();
  }

  /**
   * One more than the highest object number the cross-reference data knows
   * about, or the trailer's /Size if that is larger. A new object numbered at
   * or above this cannot collide with an existing one even when /Size
   * understates the file.
   */
  get objectNumberBound(): number {
    let bound = dictGetNumber(this.trailer, "Size") ?? 1;
    for (const objNum of this.xref.keys()) {
      if (objNum + 1 > bound) {
        bound = objNum + 1;
      }
    }
    return bound;
  }

  /** Throw the recorded {@link limitError}, if any. */
  throwIfLimitExceeded(): void {
    if (this.limitError) {
      throw this.limitError;
    }
  }

  /** Remember the first limit hit, then rethrow whatever was thrown. */
  private recordLimit(err: unknown): never {
    if (err instanceof PdfLimitExceededError) {
      this.limitError ??= err;
    }
    throw err;
  }

  private parse(tokenizer: PdfTokenizer): PdfObject {
    this.throwIfLimitExceeded();
    try {
      return parseObject(tokenizer);
    } catch (err) {
      return this.recordLimit(err);
    }
  }

  private decode(data: Uint8Array, dict: PdfDictValue): Uint8Array {
    this.throwIfLimitExceeded();
    try {
      return decodeStreamFilters(data, dict, this.maxDecodedBytes);
    } catch (err) {
      return this.recordLimit(err);
    }
  }

  /** Get the underlying raw data */
  get data(): Uint8Array {
    return this.tokenizer.bytes;
  }

  // ===========================================================================
  // File Structure Parsing
  // ===========================================================================

  private parseFileStructure(): PdfDictValue {
    try {
      const startxrefOffset = this.findStartxref();
      const trailer = this.parseXrefChain(startxrefOffset);
      this.startxrefOffset = startxrefOffset;
      return trailer;
    } catch {
      this.throwIfLimitExceeded();
      // If normal xref parsing fails, attempt full-file reconstruction
      this.xrefRecovered = true;
      this.startxrefOffset = null;
      return this.reconstructXref();
    }
  }

  /**
   * Find the startxref offset by scanning backward from EOF.
   */
  private findStartxref(): number {
    const data = this.tokenizer.bytes;
    const startxrefKeyword = _encoder.encode("startxref");

    const pos = this.tokenizer.findSequenceBackward(startxrefKeyword);
    if (pos < 0) {
      throw new PdfStructureError("Could not find startxref keyword");
    }

    // Position after "startxref"
    this.tokenizer.pos = pos + startxrefKeyword.length;
    this.tokenizer.skipWhitespaceAndComments();

    // Read the offset number
    let numStr = "";
    while (this.tokenizer.pos < data.length) {
      const b = data[this.tokenizer.pos];
      if (b >= 0x30 && b <= 0x39) {
        numStr += String.fromCharCode(b);
        this.tokenizer.pos++;
      } else {
        break;
      }
    }

    const offset = parseInt(numStr, 10);
    if (isNaN(offset)) {
      throw new PdfStructureError("Invalid startxref offset");
    }

    return offset;
  }

  /**
   * Parse the xref chain starting at the given offset.
   * Follows /Prev links for incremental updates.
   * Returns the merged trailer dictionary.
   */
  private parseXrefChain(startOffset: number): PdfDictValue {
    let trailerDict: PdfDictValue | null = null;
    let offset: number | null = startOffset;
    const visited = new Set<number>();

    while (offset !== null) {
      if (visited.has(offset)) {
        break; // Prevent infinite loops
      }
      visited.add(offset);

      this.tokenizer.pos = offset;
      this.tokenizer.skipWhitespaceAndComments();

      // Check if this is a traditional xref table or an xref stream
      const peekStart = this.tokenizer.pos;
      const firstToken = this.tokenizer.next();

      if (firstToken.type === TokenType.Keyword && firstToken.strValue === "xref") {
        // Traditional xref table
        const trailer = this.parseTraditionalXref();
        // Hybrid-reference file (ISO 32000-1 §7.5.8.4): objects missing from
        // this table — typically those in object streams — are listed in the
        // cross-reference stream /XRefStm names. It is searched after this
        // section's table and before /Prev; first entry wins, so parsing it
        // now gives exactly that order. Its own dictionary is not a trailer.
        const xrefStm = dictGetNumber(trailer, "XRefStm");
        if (xrefStm !== undefined && !visited.has(xrefStm)) {
          visited.add(xrefStm);
          this.parseXrefStream(xrefStm);
        }
        if (!trailerDict) {
          trailerDict = trailer;
        } else {
          // Merge: first trailer wins for Root, Info, Encrypt, ID
          this.mergeTrailer(trailerDict, trailer);
        }
        const prev = dictGetNumber(trailer, "Prev");
        offset = prev ?? null;
      } else if (firstToken.type === TokenType.Number) {
        // Xref stream (PDF 1.5+): starts with `N gen obj`
        this.tokenizer.pos = peekStart;
        const trailer = this.parseXrefStream(offset);
        if (!trailerDict) {
          trailerDict = trailer;
        } else {
          this.mergeTrailer(trailerDict, trailer);
        }
        const prev = dictGetNumber(trailer, "Prev");
        offset = prev ?? null;
      } else {
        throw new PdfStructureError(
          `Invalid xref at offset ${offset}: expected 'xref' keyword or xref stream`
        );
      }
    }

    if (!trailerDict) {
      throw new PdfStructureError("No trailer found");
    }

    return trailerDict;
  }

  /**
   * Parse a traditional xref table and its trailer.
   */
  private parseTraditionalXref(): PdfDictValue {
    // The "xref" keyword has already been consumed
    while (true) {
      this.tokenizer.skipWhitespaceAndComments();

      // Check if we've hit the trailer
      const peekPos = this.tokenizer.pos;
      const token = this.tokenizer.next();

      if (token.type === TokenType.Keyword && token.strValue === "trailer") {
        break;
      }

      // Subsection header: startObj count
      if (token.type !== TokenType.Number) {
        // End of xref sections
        this.tokenizer.pos = peekPos;
        break;
      }

      const startObj = token.numValue!;
      const countToken = this.tokenizer.next();
      if (countToken.type !== TokenType.Number) {
        throw new PdfStructureError("Invalid xref subsection header");
      }
      const count = countToken.numValue!;

      // Parse entries
      for (let i = 0; i < count; i++) {
        const objNum = startObj + i;
        this.tokenizer.skipWhitespaceAndComments();

        // Each entry is exactly "OOOOOOOOOO GGGGG n \n" or "OOOOOOOOOO GGGGG f \n"
        const line = this.tokenizer.readLine();
        const parts = line.trim().split(/\s+/);
        if (parts.length < 3) {
          continue;
        }

        const entryOffset = parseInt(parts[0], 10);
        const gen = parseInt(parts[1], 10);
        const inUse = parts[2] === "n";

        if (inUse && !this.xref.has(objNum)) {
          this.xref.set(objNum, { offset: entryOffset, gen, type: 1 });
        }
      }
    }

    // Parse the trailer dictionary
    this.tokenizer.skipWhitespaceAndComments();
    const trailerObj = this.parse(this.tokenizer);
    if (!isPdfDict(trailerObj)) {
      throw new PdfStructureError("Expected dictionary after 'trailer' keyword");
    }

    return trailerObj;
  }

  /**
   * Parse a cross-reference stream (PDF 1.5+).
   */
  private parseXrefStream(offset: number): PdfDictValue {
    this.tokenizer.pos = offset;
    const obj = this.parse(this.tokenizer);

    if (!isPdfStream(obj)) {
      throw new PdfStructureError("Expected xref stream object");
    }

    const dict = obj.dict;
    const type = dictGetName(dict, "Type");
    if (type !== "XRef") {
      throw new PdfStructureError(`Expected /Type /XRef, got /Type /${type}`);
    }

    // Decode the stream data
    const streamData = this.decode(obj.data, dict);

    // Parse W array: [fieldSizeType, fieldSizeOffset, fieldSizeGen]
    const wArray = dictGetArray(dict, "W");
    if (!wArray || wArray.length < 3) {
      throw new PdfStructureError("Invalid /W array in xref stream");
    }
    const w0 = wArray[0] as number;
    const w1 = wArray[1] as number;
    const w2 = wArray[2] as number;
    const entrySize = w0 + w1 + w2;

    // Parse Index array (default: [0 Size])
    const size = dictGetNumber(dict, "Size") ?? 0;
    let indexArray = dictGetArray(dict, "Index");
    if (!indexArray) {
      indexArray = [0, size];
    }

    // Process entries
    let dataOffset = 0;
    for (let i = 0; i < indexArray.length; i += 2) {
      const startObj = indexArray[i] as number;
      const count = indexArray[i + 1] as number;

      for (let j = 0; j < count; j++) {
        if (dataOffset + entrySize > streamData.length) {
          break;
        }

        const objNum = startObj + j;
        const fieldType = w0 > 0 ? readIntBE(streamData, dataOffset, w0) : 1;
        const field2 = readIntBE(streamData, dataOffset + w0, w1);
        const field3 = w2 > 0 ? readIntBE(streamData, dataOffset + w0 + w1, w2) : 0;
        dataOffset += entrySize;

        if (this.xref.has(objNum)) {
          continue; // First entry wins
        }

        if (fieldType === 0) {
          // Free object — skip
        } else if (fieldType === 1) {
          // Uncompressed object: field2 = byte offset, field3 = generation
          this.xref.set(objNum, { offset: field2, gen: field3, type: 1 });
        } else if (fieldType === 2) {
          // Compressed object in object stream: field2 = objstm number, field3 = index
          this.xref.set(objNum, { offset: field2, gen: field3, type: 2 });
        }
      }
    }

    return dict;
  }

  /**
   * Reconstruct the xref table by scanning the entire file for `N N obj` patterns.
   * This is a fallback for corrupted or broken PDFs where the normal xref parsing fails.
   *
   * @returns A synthetic trailer dictionary
   */
  private reconstructXref(): PdfDictValue {
    const data = this.tokenizer.bytes;
    this.xref.clear();

    // Regex-style scan: look for patterns like "123 0 obj" in the raw bytes
    // We scan byte-by-byte looking for digit sequences followed by spaces and "obj"
    const objKeyword = _encoder.encode("obj");
    let pos = 0;

    while (pos < data.length - 5) {
      // Skip to a potential start of an object definition (digit character)
      if (data[pos] < 0x30 || data[pos] > 0x39) {
        pos++;
        continue;
      }

      // Ensure we're at a line boundary or start of file
      if (pos > 0 && data[pos - 1] !== 0x0a && data[pos - 1] !== 0x0d && data[pos - 1] !== 0x20) {
        pos++;
        continue;
      }

      // Try to read: objNum gen obj
      const savedPos = pos;
      let objNumStr = "";
      while (pos < data.length && data[pos] >= 0x30 && data[pos] <= 0x39) {
        objNumStr += String.fromCharCode(data[pos]);
        pos++;
      }

      if (objNumStr.length === 0 || pos >= data.length || data[pos] !== 0x20) {
        pos = savedPos + 1;
        continue;
      }
      pos++; // skip space

      let genStr = "";
      while (pos < data.length && data[pos] >= 0x30 && data[pos] <= 0x39) {
        genStr += String.fromCharCode(data[pos]);
        pos++;
      }

      if (genStr.length === 0 || pos >= data.length || data[pos] !== 0x20) {
        pos = savedPos + 1;
        continue;
      }
      pos++; // skip space

      // Check for "obj" keyword
      if (
        pos + objKeyword.length <= data.length &&
        data[pos] === objKeyword[0] &&
        data[pos + 1] === objKeyword[1] &&
        data[pos + 2] === objKeyword[2]
      ) {
        // Verify the character after "obj" is whitespace or delimiter
        const afterObj = pos + 3;
        if (
          afterObj >= data.length ||
          data[afterObj] === 0x20 ||
          data[afterObj] === 0x0a ||
          data[afterObj] === 0x0d ||
          data[afterObj] === 0x09 ||
          data[afterObj] === 0x3c // '<' for immediate dict/stream
        ) {
          const objNum = parseInt(objNumStr, 10);
          const gen = parseInt(genStr, 10);

          // Last definition wins: an incremental update appends a newer
          // version of an object after the one it replaces.
          this.xref.set(objNum, { offset: savedPos, gen, type: 1 });
        }
      }

      pos = savedPos + 1;
    }

    if (this.xref.size === 0) {
      throw new PdfStructureError("Could not reconstruct xref: no objects found");
    }

    // Parse every object found directly, once: object streams, cross-reference
    // stream dictionaries (the trailer of a file with no classic one) and a
    // catalog are all needed below. In file order, so "last wins" holds.
    const direct = [...this.xref].sort((a, b) => a[1].offset - b[1].offset);
    const objStms: number[] = [];
    let xrefStmDict: PdfDictValue | null = null;
    let directCatalog: PdfRef | null = null;
    for (const [objNum, entry] of direct) {
      let obj: PdfObject;
      try {
        this.tokenizer.pos = entry.offset;
        obj = this.parse(this.tokenizer);
      } catch {
        this.throwIfLimitExceeded();
        continue; // Skip unparseable objects
      }
      const dict = isPdfStream(obj) ? obj.dict : isPdfDict(obj) ? obj : null;
      const type = dict ? dictGetName(dict, "Type") : undefined;
      if (type === "Catalog") {
        directCatalog ??= { type: "ref", objNum, gen: entry.gen };
      } else if (isPdfStream(obj) && type === "ObjStm") {
        objStms.push(objNum);
      } else if (isPdfStream(obj) && type === "XRef" && obj.dict.has("Root")) {
        xrefStmDict = obj.dict;
      }
    }

    // Prefer a classic trailer, then the last cross-reference stream's
    // dictionary (which plays the trailer's role, §7.5.8.2).
    let trailer = this.findClassicTrailer();
    if (!trailer && xrefStmDict) {
      trailer = new Map();
      for (const key of ["Root", "Info", "ID", "Encrypt"]) {
        const value = xrefStmDict.get(key);
        if (value !== undefined) {
          trailer.set(key, value);
        }
      }
    }

    // Register the members of every object stream as compressed entries, as
    // pdf.js and qpdf do; otherwise a damaged file whose pages live in object
    // streams loses them. Precedence follows qpdf's reconstruction, which
    // never lets a compressed entry replace one found as `N G obj` in the
    // file: a direct object wins over any compressed copy, wherever it sits.
    // Between two object streams holding the same number, the later one in
    // the file wins, matching "last definition wins" above.
    //
    // An encrypted file's object streams cannot be decoded yet (the security
    // handler is installed after construction), and garbage decoded from
    // ciphertext could register bogus numbers, so they are skipped there.
    if (!trailer?.has("Encrypt")) {
      for (const stmNum of objStms) {
        const members = this.getObjectStreamMembers(stmNum);
        members?.forEach(({ objNum }, index) => {
          const existing = this.xref.get(objNum);
          if (!existing || existing.type === 2) {
            this.xref.set(objNum, { offset: stmNum, gen: index, type: 2 });
          }
        });
      }
    }

    if (trailer?.has("Root")) {
      if (!trailer.has("Size")) {
        trailer.set("Size", this.objectNumberBoundOf(this.xref));
      }
      return trailer;
    }

    // Build a synthetic trailer by finding the Root catalog
    const syntheticTrailer: PdfDictValue = trailer ?? new Map();
    syntheticTrailer.set("Size", this.objectNumberBoundOf(this.xref));
    let root = directCatalog;
    if (!root) {
      // The catalog may itself be compressed.
      for (const [objNum, entry] of this.xref) {
        if (entry.type !== 2) {
          continue;
        }
        const obj = this.resolve(objNum, 0);
        if (isPdfDict(obj) && dictGetName(obj, "Type") === "Catalog") {
          root = { type: "ref", objNum, gen: 0 };
          break;
        }
      }
    }
    if (root) {
      syntheticTrailer.set("Root", root);
    }

    return syntheticTrailer;
  }

  /** One more than the highest object number in `xref`. */
  private objectNumberBoundOf(xref: Map<number, XrefEntry>): number {
    let bound = 1;
    for (const objNum of xref.keys()) {
      bound = Math.max(bound, objNum + 1);
    }
    return bound;
  }

  /** The dictionary after the last `trailer` keyword in the file, if any. */
  private findClassicTrailer(): PdfDictValue | null {
    const trailerKeyword = _encoder.encode("trailer");
    const trailerPos = this.tokenizer.findSequenceBackward(trailerKeyword);
    if (trailerPos < 0) {
      return null;
    }
    this.tokenizer.pos = trailerPos + trailerKeyword.length;
    this.tokenizer.skipWhitespaceAndComments();
    try {
      const trailerObj = this.parse(this.tokenizer);
      return isPdfDict(trailerObj) ? trailerObj : null;
    } catch {
      this.throwIfLimitExceeded();
      return null;
    }
  }

  /**
   * Merge trailer entries from an older trailer into the current one.
   * Only adds keys that don't already exist.
   */
  private mergeTrailer(current: PdfDictValue, older: PdfDictValue): void {
    for (const [key, value] of older) {
      if (!current.has(key)) {
        current.set(key, value);
      }
    }
  }

  // ===========================================================================
  // Object Resolution
  // ===========================================================================

  /**
   * Resolve a PDF object by its object number and generation.
   * Returns null if the object doesn't exist.
   */
  resolve(objNum: number, gen = 0): PdfObject | null {
    this.throwIfLimitExceeded();
    const cacheKey = `${objNum}:${gen}`;
    if (this.cache.has(cacheKey)) {
      return this.cache.get(cacheKey)!;
    }

    const entry = this.xref.get(objNum);
    if (!entry) {
      return null;
    }

    let obj: PdfObject | null = null;

    if (entry.type === 1) {
      // Uncompressed object — parse directly at offset
      obj = this.parseObjectAt(entry.offset, objNum, entry.gen);
    } else if (entry.type === 2) {
      // Compressed object in an object stream
      obj = this.parseCompressedObject(entry.offset, entry.gen);
    }

    // Decrypt string values within the resolved object. Only an indirect
    // object's own strings are encrypted: an object stream is encrypted as a
    // whole, and the strings inside it are not encrypted again (§7.5.7).
    if (obj !== null && entry.type === 1 && this.decryptFn) {
      obj = this.decryptObjectStrings(obj, objNum, entry.gen);
    }

    if (obj !== null) {
      this.cache.set(cacheKey, obj);
    }
    return obj;
  }

  /**
   * Resolve a PDF object and return it along with its object/generation numbers.
   * Useful for tracking which object a value came from (for decryption).
   *
   * @param objNum - The object number to resolve
   * @param gen - The generation number (default 0)
   * @returns The resolved object with its objNum and gen for decryption context
   */
  resolveWithObjNum(objNum: number, gen = 0): ResolvedObject {
    const obj = this.resolve(objNum, gen);
    return { obj, objNum, gen };
  }

  /**
   * Dereference a PdfRef to its actual object value.
   * If the input is not a PdfRef, returns it as-is.
   */
  deref(obj: PdfObject | null | undefined): PdfObject | null {
    this.throwIfLimitExceeded();
    if (obj === null || obj === undefined) {
      return null;
    }
    if (isPdfRef(obj)) {
      return this.resolve(obj.objNum, obj.gen);
    }
    return obj;
  }

  /**
   * Dereference a PdfRef and assert it's a dictionary.
   */
  derefDict(obj: PdfObject | null | undefined): PdfDictValue | null {
    const resolved = this.deref(obj);
    if (resolved === null) {
      return null;
    }
    if (isPdfDict(resolved)) {
      return resolved;
    }
    if (isPdfStream(resolved)) {
      return resolved.dict;
    }
    return null;
  }

  /**
   * Dereference a PdfRef and get the stream, along with the objNum/gen
   * needed for correct per-object decryption.
   */
  derefStream(obj: PdfObject | null | undefined): PdfStream | null {
    const resolved = this.deref(obj);
    if (resolved === null) {
      return null;
    }
    if (isPdfStream(resolved)) {
      return resolved;
    }
    return null;
  }

  /**
   * Dereference a PdfRef and get the stream with its object number and generation.
   * Returns null if the object is not a stream.
   * The objNum/gen are needed for correct per-object decryption (V1-V4).
   */
  derefStreamWithObjNum(
    obj: PdfObject | null | undefined
  ): { stream: PdfStream; objNum: number; gen: number } | null {
    if (obj === null || obj === undefined) {
      return null;
    }
    let objNum = 0;
    let gen = 0;
    if (isPdfRef(obj)) {
      objNum = obj.objNum;
      gen = obj.gen;
    }
    const resolved = this.deref(obj);
    if (resolved === null) {
      return null;
    }
    if (isPdfStream(resolved)) {
      return { stream: resolved, objNum, gen };
    }
    return null;
  }

  /**
   * Get decoded stream data from a stream object.
   * Applies filter chain decoding and decryption.
   *
   * When objNum/gen are not provided (default 0), decryption may not
   * produce correct results. Use {@link resolveWithObjNum} to obtain
   * the correct objNum/gen for the stream's containing object.
   */
  getStreamData(stream: PdfStream, objNum = 0, gen = 0): Uint8Array {
    this.throwIfLimitExceeded();
    let data = stream.data;

    // Decrypt stream data if encryption is active
    if (this.decryptFn && this.isStreamEncrypted(stream.dict)) {
      data = this.decryptFn(data, objNum, gen, "stream");
    }

    return this.decode(data, stream.dict);
  }

  /**
   * Decrypt a string value (bytes) if encryption is active.
   */
  decryptString(bytes: Uint8Array, objNum: number, gen: number): Uint8Array {
    if (this.decryptFn) {
      return this.decryptFn(bytes, objNum, gen, "string");
    }
    return bytes;
  }

  /**
   * Whether a stream's data went through the security handler: not a
   * metadata stream when `/EncryptMetadata` is false (§7.6.3.2), and not a
   * stream whose own `/Crypt` filter selects `/Identity` (§7.4.10).
   */
  private isStreamEncrypted(dict: PdfDictValue): boolean {
    if (!this.encryptMetadata && dictGetName(dict, "Type") === "Metadata") {
      return false;
    }
    const filter = dict.get("Filter");
    const first = Array.isArray(filter) ? filter[0] : filter;
    if (first === "Crypt") {
      const parms = dict.get("DecodeParms");
      const firstParms = Array.isArray(parms) ? parms[0] : parms;
      const name = firstParms instanceof Map ? firstParms.get("Name") : undefined;
      return name !== undefined && name !== "Identity";
    }
    return true;
  }

  /**
   * Decode a PDF string to a JS string, with optional decryption.
   */
  decodeString(bytes: Uint8Array, objNum = 0, gen = 0): string {
    const decrypted = this.decryptString(bytes, objNum, gen);
    return decodePdfStringBytes(decrypted);
  }

  /**
   * Recursively decrypt all string values (Uint8Array) within a parsed PDF object.
   * PDF spec requires all strings in an encrypted document to be decrypted using
   * the per-object key derived from the containing object's objNum/gen.
   * Streams are NOT decrypted here — they are decrypted in getStreamData().
   */
  private decryptObjectStrings(obj: PdfObject, objNum: number, gen: number): PdfObject {
    if (obj === null || typeof obj !== "object") {
      return obj;
    }

    // Decrypt Uint8Array string values
    if (obj instanceof Uint8Array) {
      return this.decryptFn!(obj, objNum, gen, "string");
    }

    // Recurse into dictionaries
    if (isPdfDict(obj)) {
      const decrypted: PdfDictValue = new Map();
      for (const [key, value] of obj) {
        decrypted.set(key, this.decryptObjectStrings(value, objNum, gen));
      }
      return decrypted;
    }

    // Recurse into arrays
    if (isPdfArray(obj)) {
      return obj.map(item => this.decryptObjectStrings(item, objNum, gen));
    }

    // Decrypt strings inside stream dicts (but NOT the stream data itself)
    if (isPdfStream(obj)) {
      const decryptedDict = this.decryptObjectStrings(obj.dict, objNum, gen) as PdfDictValue;
      return { type: "stream" as const, dict: decryptedDict, data: obj.data };
    }

    return obj;
  }

  /**
   * Get the catalog dictionary (the root of the document structure).
   */
  getCatalog(): PdfDictValue {
    const rootRef = dictGetRef(this.trailer, "Root");
    if (!rootRef) {
      throw new PdfStructureError("No /Root in trailer");
    }
    const catalog = this.derefDict(rootRef);
    if (!catalog) {
      throw new PdfStructureError("Could not resolve catalog");
    }
    return catalog;
  }

  /**
   * Get the pages array from the page tree.
   * Returns an array of page dictionaries in order.
   */
  getPages(): PdfDictValue[] {
    return this.getPagesWithObjInfo().map(p => p.dict);
  }

  /**
   * Get pages with their object numbers (needed for correct decryption of
   * inline streams within page objects).
   */
  getPagesWithObjInfo(): Array<{ dict: PdfDictValue; objNum: number; gen: number }> {
    const catalog = this.getCatalog();
    const pagesRef = catalog.get("Pages");
    const pagesDict = this.derefDict(pagesRef);
    if (!pagesDict) {
      throw new PdfStructureError("Could not resolve /Pages");
    }
    const pages: Array<{ dict: PdfDictValue; objNum: number; gen: number }> = [];
    const visited = new Set<PdfDictValue>();
    this.collectPages(pagesDict, pages, visited);
    return pages;
  }

  /**
   * Recursively collect page dictionaries from the page tree.
   * Uses a visited set to prevent infinite recursion on cyclic page trees.
   */
  private collectPages(
    node: PdfDictValue,
    pages: Array<{ dict: PdfDictValue; objNum: number; gen: number }>,
    visited: Set<PdfDictValue>
  ): void {
    if (visited.has(node)) {
      return; // Cycle guard
    }
    visited.add(node);

    const type = dictGetName(node, "Type");

    if (type === "Page") {
      // We don't know the objNum from here — it was lost during deref.
      // Use 0 as fallback; callers that need objNum should use getPagesWithObjInfo().
      pages.push({ dict: node, objNum: 0, gen: 0 });
      return;
    }

    // Pages node — recurse into Kids
    const kids = dictGetArray(node, "Kids");
    if (!kids) {
      return;
    }

    for (const kid of kids) {
      let objNum = 0;
      let gen = 0;
      if (isPdfRef(kid)) {
        objNum = kid.objNum;
        gen = kid.gen;
      }
      const childDict = this.derefDict(kid);
      if (childDict) {
        const childType = dictGetName(childDict, "Type");
        if (childType === "Page") {
          pages.push({ dict: childDict, objNum, gen });
        } else {
          this.collectPages(childDict, pages, visited);
        }
      }
    }
  }

  /**
   * Get the object number for a given object reference.
   * Useful for tracking which object a value came from (for decryption).
   */
  getObjNumForRef(ref: PdfRef): number {
    return ref.objNum;
  }

  // ===========================================================================
  // Low-level Object Parsing
  // ===========================================================================

  /**
   * Parse an object definition at the given byte offset.
   */
  private parseObjectAt(offset: number, objNum: number, _gen: number): PdfObject | null {
    this.tokenizer.pos = offset;
    try {
      const obj = this.parse(this.tokenizer);
      return obj;
    } catch {
      this.throwIfLimitExceeded();
      return null;
    }
  }

  /**
   * Parse a compressed object from an object stream.
   * @param objStmNum - The object number of the object stream
   * @param index - The index of the object within the stream
   */
  private parseCompressedObject(objStmNum: number, index: number): PdfObject | null {
    return this.getObjectStreamMembers(objStmNum)?.[index]?.obj ?? null;
  }

  /** The members of object stream `objStmNum`, parsed once and cached. */
  private getObjectStreamMembers(objStmNum: number): ObjStmMember[] | null {
    let members = this.objStmCache.get(objStmNum);
    if (!members) {
      members = this.parseObjectStream(objStmNum) ?? undefined;
      if (!members) {
        return null;
      }
      this.objStmCache.set(objStmNum, members);
    }
    return members;
  }

  /**
   * Parse all objects from an object stream, in index order.
   *
   * The header is untrusted: /N and /First are clamped to the decoded data,
   * reading stops at the first pair that is not two non-negative integers,
   * and a member whose offset falls outside the data is kept as `null` so the
   * indices of the members after it stay correct.
   */
  private parseObjectStream(objStmNum: number): ObjStmMember[] | null {
    const entry = this.xref.get(objStmNum);
    if (!entry || entry.type !== 1) {
      return null;
    }

    this.tokenizer.pos = entry.offset;
    let stmObj: PdfObject;
    try {
      stmObj = this.parse(this.tokenizer);
    } catch {
      this.throwIfLimitExceeded();
      return null;
    }
    if (!isPdfStream(stmObj)) {
      return null;
    }

    const dict = stmObj.dict;
    const n = dictGetNumber(dict, "N") ?? 0;
    const first = dictGetNumber(dict, "First") ?? 0;

    // Decode stream data (pass objStmNum/gen for correct decryption)
    const streamData = this.getStreamData(stmObj, objStmNum, entry.gen);
    if (!Number.isInteger(first) || first < 0 || first > streamData.length) {
      return null;
    }

    // Parse the N pairs of (objNum offset) before 'first'. Each pair takes at
    // least four bytes ("1 0 "), which bounds a bogus /N by the data itself.
    const count = Number.isInteger(n) && n > 0 ? Math.min(n, Math.ceil(first / 4)) : 0;
    const headerTokenizer = new PdfTokenizer(streamData.subarray(0, first));
    const members: ObjStmMember[] = [];
    for (let i = 0; i < count; i++) {
      const numTok = headerTokenizer.next();
      const offTok = headerTokenizer.next();
      if (numTok.type !== TokenType.Number || offTok.type !== TokenType.Number) {
        break;
      }
      const objectNumber = numTok.numValue!;
      const relOffset = offTok.numValue!;
      if (!Number.isInteger(objectNumber) || objectNumber <= 0 || !Number.isInteger(relOffset)) {
        break;
      }
      let obj: PdfObject | null = null;
      if (relOffset >= 0 && first + relOffset < streamData.length) {
        try {
          obj = this.parse(new PdfTokenizer(streamData, first + relOffset));
        } catch {
          this.throwIfLimitExceeded();
          // Keep the slot: an unparseable member must not shift the others.
        }
      }
      members.push({ objNum: objectNumber, obj });
    }

    return members;
  }

  /**
   * Resolve a page's bounding box (MediaBox/CropBox) with indirect ref resolution
   * and parent inheritance. Returns `{ width, height }` or null if no box found.
   *
   * This is a shared helper so callers don't duplicate box resolution logic.
   */
  resolvePageBox(
    pageDict: PdfDictValue,
    visited?: Set<PdfDictValue>
  ): { width: number; height: number } | null {
    const seen = visited ?? new Set<PdfDictValue>();
    if (seen.has(pageDict)) {
      return null; // Cycle guard
    }
    seen.add(pageDict);

    for (const key of ["MediaBox", "CropBox"]) {
      const raw = pageDict.get(key);
      if (!raw) {
        continue;
      }
      // Dereference in case the box is an indirect reference
      const resolved = this.deref(raw);
      if (Array.isArray(resolved) && resolved.length === 4) {
        const width = Math.abs((resolved[2] as number) - (resolved[0] as number));
        const height = Math.abs((resolved[3] as number) - (resolved[1] as number));
        if (width > 0 && height > 0) {
          return { width, height };
        }
      }
    }

    // Inherit from parent
    const parent = pageDict.get("Parent");
    if (parent) {
      const parentDict = this.derefDict(parent);
      if (parentDict) {
        return this.resolvePageBox(parentDict, seen);
      }
    }

    return null;
  }

  /**
   * Resolve a page's Resources dictionary, inheriting from parent pages if needed.
   * Protected against cyclic parent chains.
   */
  resolvePageResources(pageDict: PdfDictValue, visited?: Set<PdfDictValue>): PdfDictValue {
    const seen = visited ?? new Set<PdfDictValue>();
    if (seen.has(pageDict)) {
      return new Map(); // Cycle guard
    }
    seen.add(pageDict);

    const resources = pageDict.get("Resources");
    if (resources) {
      const resolved = this.derefDict(resources);
      if (resolved) {
        return resolved;
      }
    }

    const parent = pageDict.get("Parent");
    if (parent) {
      const parentDict = this.derefDict(parent);
      if (parentDict) {
        return this.resolvePageResources(parentDict, seen);
      }
    }

    return new Map();
  }
}

// =============================================================================
// Helpers
// =============================================================================

/**
 * Read a big-endian integer of the given byte width.
 * Uses multiplication instead of bitwise shift to avoid signed 32-bit overflow
 * for values that exceed 2^31 (e.g. large file offsets).
 */
function readIntBE(data: Uint8Array, offset: number, width: number): number {
  let value = 0;
  for (let i = 0; i < width; i++) {
    value = value * 256 + (data[offset + i] ?? 0);
  }
  return value;
}
