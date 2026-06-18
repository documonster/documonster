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

import { PdfStructureError } from "@pdf/errors";
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
export class PdfDocument {
  private tokenizer: PdfTokenizer;
  private xref: Map<number, XrefEntry> = new Map();
  private cache: Map<string, PdfObject> = new Map();
  declare readonly trailer: PdfDictValue;

  /** Encryption handler (set externally after decryption is initialized) */
  decryptFn: ((data: Uint8Array, objNum: number, gen: number) => Uint8Array) | null = null;

  constructor(data: Uint8Array) {
    this.tokenizer = new PdfTokenizer(data);
    this.trailer = this.parseFileStructure();
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
      return this.parseXrefChain(startxrefOffset);
    } catch {
      // If normal xref parsing fails, attempt full-file reconstruction
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
    const trailerObj = parseObject(this.tokenizer);
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
    const obj = parseObject(this.tokenizer);

    if (!isPdfStream(obj)) {
      throw new PdfStructureError("Expected xref stream object");
    }

    const dict = obj.dict;
    const type = dictGetName(dict, "Type");
    if (type !== "XRef") {
      throw new PdfStructureError(`Expected /Type /XRef, got /Type /${type}`);
    }

    // Decode the stream data
    const streamData = decodeStreamFilters(obj.data, dict);

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

          if (!this.xref.has(objNum)) {
            this.xref.set(objNum, { offset: savedPos, gen, type: 1 });
          }
        }
      }

      pos = savedPos + 1;
    }

    if (this.xref.size === 0) {
      throw new PdfStructureError("Could not reconstruct xref: no objects found");
    }

    // Try to find a trailer dictionary by scanning for "trailer" keyword
    const trailerKeyword = _encoder.encode("trailer");
    const trailerPos = this.tokenizer.findSequenceBackward(trailerKeyword);

    if (trailerPos >= 0) {
      this.tokenizer.pos = trailerPos + trailerKeyword.length;
      this.tokenizer.skipWhitespaceAndComments();
      try {
        const trailerObj = parseObject(this.tokenizer);
        if (isPdfDict(trailerObj)) {
          return trailerObj;
        }
      } catch {
        // Fall through to synthetic trailer
      }
    }

    // Build a synthetic trailer by finding the Root catalog
    const syntheticTrailer: PdfDictValue = new Map();
    syntheticTrailer.set("Size", this.xref.size);

    // Scan resolved objects to find the catalog (the one with /Type /Catalog)
    for (const [objNum, entry] of this.xref) {
      if (entry.type !== 1) {
        continue;
      }
      try {
        this.tokenizer.pos = entry.offset;
        const obj = parseObject(this.tokenizer);
        if (isPdfDict(obj)) {
          const typeVal = dictGetName(obj, "Type");
          if (typeVal === "Catalog") {
            syntheticTrailer.set("Root", { type: "ref", objNum, gen: entry.gen } as PdfRef);
            break;
          }
        } else if (isPdfStream(obj)) {
          const typeVal = dictGetName(obj.dict, "Type");
          if (typeVal === "Catalog") {
            syntheticTrailer.set("Root", { type: "ref", objNum, gen: entry.gen } as PdfRef);
            break;
          }
        }
      } catch {
        // Skip unparseable objects
      }
    }

    return syntheticTrailer;
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

    // Decrypt string values within the resolved object (V1-V4 per-object encryption)
    if (obj !== null && this.decryptFn) {
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
    let data = stream.data;

    // Decrypt stream data if encryption is active
    if (this.decryptFn) {
      data = this.decryptFn(data, objNum, gen);
    }

    return decodeStreamFilters(data, stream.dict);
  }

  /**
   * Decrypt a string value (bytes) if encryption is active.
   */
  decryptString(bytes: Uint8Array, objNum: number, gen: number): Uint8Array {
    if (this.decryptFn) {
      return this.decryptFn(bytes, objNum, gen);
    }
    return bytes;
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
      return this.decryptFn!(obj, objNum, gen);
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
      const obj = parseObject(this.tokenizer);
      return obj;
    } catch {
      return null;
    }
  }

  /**
   * Parse a compressed object from an object stream.
   * @param objStmNum - The object number of the object stream
   * @param index - The index of the object within the stream
   */
  private parseCompressedObject(objStmNum: number, index: number): PdfObject | null {
    // Resolve the object stream itself (must be type 1 — not recursive)
    const stmCacheKey = `objstm:${objStmNum}`;
    let stmObjects: Map<number, PdfObject> | undefined;

    if (this.cache.has(stmCacheKey)) {
      stmObjects = this.cache.get(stmCacheKey) as unknown as Map<number, PdfObject>;
    } else {
      stmObjects = this.parseObjectStream(objStmNum) ?? undefined;
      if (stmObjects) {
        this.cache.set(stmCacheKey, stmObjects as unknown as PdfObject);
      }
    }

    if (!stmObjects) {
      return null;
    }

    // The index field in the xref is the index within the object stream
    // We need to find the object by its index position
    let i = 0;
    for (const [, value] of stmObjects) {
      if (i === index) {
        return value;
      }
      i++;
    }
    return null;
  }

  /**
   * Parse all objects from an object stream.
   * @returns Map of object number → object value
   */
  private parseObjectStream(objStmNum: number): Map<number, PdfObject> | null {
    const entry = this.xref.get(objStmNum);
    if (!entry || entry.type !== 1) {
      return null;
    }

    this.tokenizer.pos = entry.offset;
    const stmObj = parseObject(this.tokenizer);
    if (!isPdfStream(stmObj)) {
      return null;
    }

    const dict = stmObj.dict;
    const n = dictGetNumber(dict, "N") ?? 0;
    const first = dictGetNumber(dict, "First") ?? 0;

    // Decode stream data (pass objStmNum/gen for correct decryption)
    const streamData = this.getStreamData(stmObj, objStmNum, entry.gen);

    // Parse the N pairs of (objNum offset) before 'first'
    const headerTokenizer = new PdfTokenizer(streamData);
    const pairs: Array<[number, number]> = [];
    for (let i = 0; i < n; i++) {
      const numTok = headerTokenizer.next();
      const offTok = headerTokenizer.next();
      if (numTok.type === TokenType.Number && offTok.type === TokenType.Number) {
        pairs.push([numTok.numValue!, offTok.numValue!]);
      }
    }

    // Parse each object
    const result = new Map<number, PdfObject>();
    for (const [objectNumber, relOffset] of pairs) {
      const objTokenizer = new PdfTokenizer(streamData, first + relOffset);
      try {
        const obj = parseObject(objTokenizer);
        result.set(objectNumber, obj);
      } catch {
        // Skip unparseable objects
      }
    }

    return result;
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
