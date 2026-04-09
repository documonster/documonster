/**
 * PDF file writer.
 *
 * Assembles a complete PDF 2.0 document from indirect objects.
 * Handles the four sections of a PDF file:
 * 1. Header (%PDF-2.0)
 * 2. Body (indirect objects)
 * 3. Cross-reference table
 * 4. Trailer (with document catalog reference)
 *
 * Encryption uses AES-256 (V=5, R=5) per ISO 32000-2:2020.
 *
 * @see ISO 32000-2:2020, Chapter 7.5 — File Structure
 */

import { PdfDict, pdfRef, pdfString, pdfHexString, pdfDate, pdfNumber } from "./pdf-object";
import type { PdfContentStream } from "./pdf-stream";
import { PdfStructureError } from "../errors";
import { concatUint8Arrays } from "@utils/binary";
import { zlibSync } from "@archive/compression/compress";
import type { EncryptionState } from "./encryption";
import { encryptData } from "./encryption";

// =============================================================================
// Types
// =============================================================================

interface PdfIndirectObject {
  objectNumber: number;
  /** Byte offset of the object in the output buffer */
  offset: number;
  /** Serialized content (between `N 0 obj` and `endobj`) */
  content: string;
}

interface PdfStreamObject extends PdfIndirectObject {
  streamData: Uint8Array;
}

// =============================================================================
// PDF Writer
// =============================================================================

/**
 * Constructs a valid PDF 2.0 file from a set of indirect objects.
 *
 * Usage:
 * 1. Allocate object numbers with allocObject()
 * 2. Add objects with addObject() or addStreamObject()
 * 3. Call build() to produce the final Uint8Array
 */
export class PdfWriter {
  private nextObjectNumber = 1;
  private objects: Array<PdfIndirectObject | PdfStreamObject> = [];
  private catalogRef = 0;
  private infoRef = 0;
  private encryption: EncryptionState | null = null;

  /**
   * Enable encryption for this document.
   */
  setEncryption(state: EncryptionState): void {
    this.encryption = state;
  }

  // ===========================================================================
  // Object Management
  // ===========================================================================

  /**
   * Allocate the next object number without adding content yet.
   * Use this for forward references (e.g., page references in the page tree).
   */
  allocObject(): number {
    return this.nextObjectNumber++;
  }

  /**
   * Add a dictionary object to the PDF.
   * @param objectNumber - Previously allocated object number
   * @param dict - The dictionary content
   */
  addObject(objectNumber: number, dict: PdfDict | string): void {
    this.objects.push({
      objectNumber,
      offset: 0,
      content: typeof dict === "string" ? dict : dict.toString()
    });
  }

  /**
   * Add a stream object (dictionary + binary stream data) to the PDF.
   * The /Length key is automatically added to the dictionary.
   */
  addStreamObject(objectNumber: number, dict: PdfDict, stream: PdfContentStream): void;
  addStreamObject(objectNumber: number, dict: PdfDict, data: Uint8Array): void;
  addStreamObject(objectNumber: number, dict: PdfDict, data: PdfContentStream | Uint8Array): void {
    let streamData = data instanceof Uint8Array ? data : data.toUint8Array();

    // Compress with zlib (RFC 1950) for PDF /FlateDecode
    if (streamData.length > 256 && !dict.toString().includes("/Filter")) {
      const compressed = zlibSync(streamData, { level: 6 });
      if (compressed.length < streamData.length) {
        dict.set("Filter", "/FlateDecode");
        streamData = compressed;
      }
    }

    // NOTE: Stream encryption is deferred to build() time so that
    // setEncryption() can be called after all objects are added.
    // We store a "needsEncryption" flag; the /Length is set in build()
    // after potential encryption changes the size.
    dict.set("Length", pdfNumber(streamData.length));
    (this.objects as PdfStreamObject[]).push({
      objectNumber,
      offset: 0,
      content: dict.toString(),
      streamData
    });
  }

  /**
   * Set the document catalog object number.
   * This is required and references the root of the document structure.
   */
  setCatalog(objectNumber: number): void {
    this.catalogRef = objectNumber;
  }

  // ===========================================================================
  // Document Info Dictionary
  // ===========================================================================

  /**
   * Create and add the document information dictionary.
   */
  addInfoDict(options: {
    title?: string;
    author?: string;
    subject?: string;
    creator?: string;
  }): number {
    const objNum = this.allocObject();
    const dict = new PdfDict();

    if (options.title) {
      dict.set("Title", pdfString(options.title));
    }
    if (options.author) {
      dict.set("Author", pdfString(options.author));
    }
    if (options.subject) {
      dict.set("Subject", pdfString(options.subject));
    }
    if (options.creator) {
      dict.set("Creator", pdfString(options.creator));
    }
    dict.set("Producer", pdfString("excelts"));
    dict.set("CreationDate", pdfDate(new Date()));

    this.addObject(objNum, dict);
    this.infoRef = objNum;
    return objNum;
  }

  // ===========================================================================
  // Page Tree Helpers
  // ===========================================================================

  /**
   * Create and add a Page dictionary.
   */
  addPage(options: {
    parentRef: number;
    width: number;
    height: number;
    contentsRef: number | string;
    resourcesRef: number;
    annotRefs?: number[];
  }): number {
    const objNum = this.allocObject();
    const mediaBox = `[0 0 ${pdfNumber(options.width)} ${pdfNumber(options.height)}]`;
    const contentsValue =
      typeof options.contentsRef === "string" ? options.contentsRef : pdfRef(options.contentsRef);
    const dict = new PdfDict()
      .set("Type", "/Page")
      .set("Parent", pdfRef(options.parentRef))
      .set("MediaBox", mediaBox)
      .set("Contents", contentsValue)
      .set("Resources", pdfRef(options.resourcesRef));
    if (options.annotRefs && options.annotRefs.length > 0) {
      dict.set("Annots", "[" + options.annotRefs.map(r => pdfRef(r)).join(" ") + "]");
    }
    this.addObject(objNum, dict);
    return objNum;
  }

  /**
   * Create and add the Catalog dictionary.
   */
  addCatalog(pagesRef: number, outlinesRef?: number): number {
    const objNum = this.allocObject();
    const dict = new PdfDict().set("Type", "/Catalog").set("Pages", pdfRef(pagesRef));
    if (outlinesRef) {
      dict.set("Outlines", pdfRef(outlinesRef));
      dict.set("PageMode", "/UseOutlines");
    }
    this.addObject(objNum, dict);
    this.setCatalog(objNum);
    return objNum;
  }

  // ===========================================================================
  // Build
  // ===========================================================================

  /**
   * Build the complete PDF file as a Uint8Array.
   */
  build(): Uint8Array {
    if (this.catalogRef === 0) {
      throw new PdfStructureError("No catalog object set. Call addCatalog() before build().");
    }

    const encoder = new TextEncoder();
    const chunks: Uint8Array[] = [];
    let byteOffset = 0;

    // --- Header ---
    // Include a comment with high bytes to signal binary content per PDF spec §3.4.1
    const headerStr = "%PDF-2.0\n";
    const headerStrBytes = encoder.encode(headerStr);
    chunks.push(headerStrBytes);
    byteOffset += headerStrBytes.length;
    // Binary comment: raw high-bit bytes (not via TextEncoder to avoid UTF-8 multibyte)
    const binaryComment = new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]);
    chunks.push(binaryComment);
    byteOffset += binaryComment.length;

    // Sort objects by object number for deterministic output
    this.objects.sort((a, b) => a.objectNumber - b.objectNumber);

    // Pre-allocate encrypt object number (content written later, after body)
    const encryptObjNum = this.encryption ? this.allocObject() : 0;

    // --- Body (indirect objects) ---
    for (const obj of this.objects) {
      obj.offset = byteOffset;

      const objHeader = `${obj.objectNumber} 0 obj\n`;
      const objHeaderBytes = encoder.encode(objHeader);
      chunks.push(objHeaderBytes);
      byteOffset += objHeaderBytes.length;

      // Encrypt string literals in object content (if encryption enabled)
      let content = obj.content;
      if (this.encryption && obj.objectNumber !== encryptObjNum) {
        content = encryptStringsInContent(
          content,
          obj.objectNumber,
          0,
          this.encryption.encryptionKey
        );
      }

      // If this is a stream object, encrypt stream data and update /Length
      let streamBytes: Uint8Array | null = null;
      if ("streamData" in obj && obj.streamData) {
        streamBytes = obj.streamData;
        if (this.encryption && obj.objectNumber !== encryptObjNum) {
          streamBytes = encryptData(
            streamBytes,
            obj.objectNumber,
            0,
            this.encryption.encryptionKey
          );
          // Update /Length in content to reflect encrypted size
          content = content.replace(/\/Length \d+/, `/Length ${streamBytes.length}`);
        }
      }

      const contentBytes = encoder.encode(content + "\n");
      chunks.push(contentBytes);
      byteOffset += contentBytes.length;

      // Write stream data if present
      if (streamBytes) {
        const streamStart = encoder.encode("stream\n");
        chunks.push(streamStart);
        byteOffset += streamStart.length;

        chunks.push(streamBytes);
        byteOffset += streamBytes.length;

        const streamEnd = encoder.encode("\nendstream\n");
        chunks.push(streamEnd);
        byteOffset += streamEnd.length;
      }

      const objFooter = encoder.encode("endobj\n");
      chunks.push(objFooter);
      byteOffset += objFooter.length;
    }

    // --- Encrypt dictionary (V=5, R=5, AES-256) ---
    if (this.encryption) {
      const encDict = new PdfDict()
        .set("Filter", "/Standard")
        .set("V", "5")
        .set("R", "5")
        .set("Length", "256")
        .set("P", String(this.encryption.permissions))
        .set("O", pdfHexString(this.encryption.oValue))
        .set("U", pdfHexString(this.encryption.uValue))
        .set("OE", pdfHexString(this.encryption.oeValue))
        .set("UE", pdfHexString(this.encryption.ueValue))
        .set("Perms", pdfHexString(this.encryption.permsValue))
        .set("EncryptMetadata", "true")
        .set(
          "CF",
          "<< /StdCF << /Type /CryptFilter /CFM /AESV3 /AuthEvent /DocOpen /Length 32 >> >>"
        )
        .set("StmF", "/StdCF")
        .set("StrF", "/StdCF");
      const encContent = encDict.toString();
      const encObj: PdfIndirectObject = {
        objectNumber: encryptObjNum,
        offset: byteOffset,
        content: encContent
      };
      this.objects.push(encObj);

      const encHeader = encoder.encode(`${encryptObjNum} 0 obj\n`);
      chunks.push(encHeader);
      byteOffset += encHeader.length;
      const encBody = encoder.encode(encContent + "\n");
      chunks.push(encBody);
      byteOffset += encBody.length;
      const encFooter = encoder.encode("endobj\n");
      chunks.push(encFooter);
      byteOffset += encFooter.length;
    }

    // --- Cross-Reference Table ---
    const xrefOffset = byteOffset;

    const objMap = new Map<number, PdfIndirectObject>();
    for (const obj of this.objects) {
      objMap.set(obj.objectNumber, obj);
    }
    const maxObjNum = this.objects.reduce((max, o) => Math.max(max, o.objectNumber), 0);

    const xrefLines: string[] = [];
    xrefLines.push("xref");
    xrefLines.push(`0 ${maxObjNum + 1}`);

    // Build a proper free list per PDF spec §3.4.3:
    // Object 0 is always free and points to the next free object.
    // Each free entry's first field is the next free object number.
    // Generation number is 65535 for never-used entries.
    // The last free entry points back to 0.
    const freeObjects: number[] = [0];
    for (let i = 1; i <= maxObjNum; i++) {
      if (!objMap.has(i)) {
        freeObjects.push(i);
      }
    }

    // Build next-free linked list
    const nextFree = new Map<number, number>();
    for (let i = 0; i < freeObjects.length - 1; i++) {
      nextFree.set(freeObjects[i], freeObjects[i + 1]);
    }
    nextFree.set(freeObjects[freeObjects.length - 1], 0); // last points to 0

    // Object 0: head of free list
    const obj0Next = (nextFree.get(0) ?? 0).toString().padStart(10, "0");
    xrefLines.push(`${obj0Next} 65535 f `);

    for (let i = 1; i <= maxObjNum; i++) {
      const obj = objMap.get(i);
      if (obj) {
        const offsetStr = obj.offset.toString().padStart(10, "0");
        xrefLines.push(`${offsetStr} 00000 n `);
      } else {
        const nextObj = (nextFree.get(i) ?? 0).toString().padStart(10, "0");
        xrefLines.push(`${nextObj} 65535 f `);
      }
    }

    const xrefStr = xrefLines.join("\n") + "\n";
    const xrefBytes = encoder.encode(xrefStr);
    chunks.push(xrefBytes);

    // --- Trailer ---
    let trailerStr = "trailer\n<<\n";
    trailerStr += `/Size ${maxObjNum + 1}\n`;
    trailerStr += `/Root ${pdfRef(this.catalogRef)}\n`;
    if (this.infoRef > 0) {
      trailerStr += `/Info ${pdfRef(this.infoRef)}\n`;
    }
    if (encryptObjNum > 0) {
      trailerStr += `/Encrypt ${pdfRef(encryptObjNum)}\n`;
      trailerStr += `/ID [${pdfHexString(this.encryption!.fileId)} ${pdfHexString(this.encryption!.fileId)}]\n`;
    }
    trailerStr += ">>\n";
    trailerStr += "startxref\n";
    trailerStr += `${xrefOffset}\n`;
    trailerStr += "%%EOF\n";

    const trailerBytes = encoder.encode(trailerStr);
    chunks.push(trailerBytes);

    // --- Concatenate all chunks ---
    return concatUint8Arrays(chunks);
  }
}

// =============================================================================
// String Encryption Helper
// =============================================================================

/**
 * Encrypt all PDF string literals `(...)` in an object's content.
 * PDF spec §3.5 requires all strings (except those in the Encrypt dict) to be encrypted.
 * After encryption, strings are replaced with hex strings `<...>`.
 *
 * Uses a character-by-character parser instead of regex to correctly handle
 * nested escaped parentheses and backslashes like `(a\\(b\\)c)`.
 */
function encryptStringsInContent(
  content: string,
  objectNumber: number,
  generation: number,
  encryptionKey: Uint8Array
): string {
  const encoder = new TextEncoder();
  const result: string[] = [];
  let i = 0;

  while (i < content.length) {
    if (content[i] === "(") {
      // Parse a balanced parenthesized string, respecting \ escapes
      const start = i;
      i++; // skip opening '('
      let inner = "";
      let depth = 1;
      while (i < content.length && depth > 0) {
        const ch = content[i];
        if (ch === "\\") {
          // Escaped character — consume both chars
          inner += content[i] + (content[i + 1] ?? "");
          i += 2;
        } else if (ch === "(") {
          depth++;
          inner += ch;
          i++;
        } else if (ch === ")") {
          depth--;
          if (depth > 0) {
            inner += ch;
          }
          i++;
        } else {
          inner += ch;
          i++;
        }
      }

      if (depth !== 0) {
        // Malformed — emit as-is
        result.push(content.slice(start, i));
        continue;
      }

      // Unescape the PDF string to get raw bytes
      const unescaped = unescapePdfString(inner);
      const rawBytes = encoder.encode(unescaped);
      const encrypted = encryptData(rawBytes, objectNumber, generation, encryptionKey);
      result.push(pdfHexString(encrypted));
    } else if (content[i] === "<" && content[i + 1] === "<") {
      // Dict delimiter << — emit both chars and skip
      result.push("<<");
      i += 2;
    } else if (content[i] === ">" && content[i + 1] === ">") {
      // Dict delimiter >> — emit both chars and skip
      result.push(">>");
      i += 2;
    } else if (content[i] === "<") {
      // Parse a hex string <...>
      i++; // skip opening '<'
      let hex = "";
      while (i < content.length && content[i] !== ">") {
        hex += content[i];
        i++;
      }
      if (i < content.length) {
        i++; // skip closing '>'
      }

      // Decode hex to raw bytes
      const cleanHex = hex.replace(/\s/g, "");
      if (cleanHex.length === 0) {
        result.push("<>");
        continue;
      }
      const rawBytes = new Uint8Array(Math.ceil(cleanHex.length / 2));
      for (let h = 0; h < rawBytes.length; h++) {
        rawBytes[h] = parseInt(cleanHex.substring(h * 2, h * 2 + 2).padEnd(2, "0"), 16);
      }
      const encrypted = encryptData(rawBytes, objectNumber, generation, encryptionKey);
      result.push(pdfHexString(encrypted));
    } else {
      result.push(content[i]);
      i++;
    }
  }

  return result.join("");
}

function unescapePdfString(value: string): string {
  let result = "";

  for (let i = 0; i < value.length; i++) {
    const ch = value[i];
    if (ch !== "\\") {
      result += ch;
      continue;
    }

    const next = value[i + 1];
    if (next === undefined) {
      result += "\\";
      break;
    }

    switch (next) {
      case "n":
        result += "\n";
        break;
      case "r":
        result += "\r";
        break;
      case "t":
        result += "\t";
        break;
      case "(":
        result += "(";
        break;
      case ")":
        result += ")";
        break;
      case "\\":
        result += "\\";
        break;
      default:
        result += next;
        break;
    }
    i++;
  }

  return result;
}
