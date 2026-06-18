/**
 * PDF/A-1b compliance utilities.
 *
 * Provides XMP metadata stream writing and OutputIntent creation for
 * PDF/A-1b (ISO 19005-1, Level B) conformance.
 *
 * **Limitations:**
 * - Type1 base fonts (Helvetica, Times-Roman, Courier, etc.) are NOT embedded.
 *   PDF/A-1b strictly requires all fonts to be embedded. Documents using only
 *   CIDFonts (embedded TrueType via `embedFont()`) are fully compliant.
 *   Documents using base Type1 fonts will pass structural validation but may
 *   fail strict PDF/A font-embedding checks.
 *
 * @see ISO 19005-1:2005 — Document management — Electronic document file
 *      format for long-term preservation — Part 1: Use of PDF 1.4 (PDF/A-1)
 */

import { PdfDict, pdfNumber, pdfRef } from "@pdf/core/pdf-object";
import type { PdfWriter } from "@pdf/core/pdf-writer";

// =============================================================================
// sRGB ICC Profile
// =============================================================================

/**
 * Minimal sRGB ICC profile (v2.1.0).
 *
 * This is a valid ICC profile with the correct header structure, profile
 * signature, and a minimal tag table. It identifies the color space as RGB
 * with the sRGB rendering intent. The profile is intentionally minimal
 * (~128 bytes) — enough to satisfy PDF/A-1b OutputIntent requirements.
 *
 * Structure:
 * - 128-byte header (profile size, preferred CMM, version, device class,
 *   color space, PCS, creation date, signature, platform, flags, etc.)
 * - Tag table with 0 tags (profile is header-only for minimal compliance)
 *
 * @see ICC.1:2001-04 — File Format for Color Profiles (v2)
 */
export const sRGB_ICC_PROFILE: Uint8Array = buildMinimalSrgbProfile();

function buildMinimalSrgbProfile(): Uint8Array {
  // ICC profile header is exactly 128 bytes
  // We add a tag table with 3 required tags: desc, wtpt, cprt
  // Each tag table entry is 12 bytes
  const TAG_COUNT = 3;
  const TAG_TABLE_SIZE = 4 + TAG_COUNT * 12; // 4 (count) + 3 * 12 (entries)

  // Tag data — description, white point, copyright
  const descData = buildTextDescriptionTag("sRGB IEC61966-2.1");
  const wtptData = buildXYZTag(0.9505, 1.0, 1.089); // D50 white point
  const cprtData = buildTextTag("No copyright");

  const descOffset = 128 + TAG_TABLE_SIZE;
  const wtptOffset = descOffset + descData.length;
  const cprtOffset = wtptOffset + wtptData.length;
  const profileSize = cprtOffset + cprtData.length;

  const buf = new Uint8Array(profileSize);
  const view = new DataView(buf.buffer);

  // --- Header (128 bytes) ---
  view.setUint32(0, profileSize); // Profile size
  // Preferred CMM: none (0)
  view.setUint8(8, 2); // Major version 2
  view.setUint8(9, 0x10); // Minor version 1.0
  // Profile/Device class: 'mntr' (monitor)
  writeAscii(buf, 12, "mntr");
  // Color space: 'RGB '
  writeAscii(buf, 16, "RGB ");
  // Profile Connection Space: 'XYZ '
  writeAscii(buf, 20, "XYZ ");
  // Creation date/time (2000-01-01 00:00:00)
  view.setUint16(24, 2000); // year
  view.setUint16(26, 1); // month
  view.setUint16(28, 1); // day
  // hour, minute, second = 0 (already zero)
  // File signature: 'acsp'
  writeAscii(buf, 36, "acsp");
  // Primary platform: 'APPL'
  writeAscii(buf, 40, "APPL");
  // Rendering intent: 0 = Perceptual
  view.setUint32(64, 0);
  // PCS illuminant (D50 XYZ): X=0.9642, Y=1.0, Z=0.8249
  writeS15Fixed16(view, 68, 0.9642);
  writeS15Fixed16(view, 72, 1.0);
  writeS15Fixed16(view, 76, 0.8249);

  // --- Tag table ---
  const tagTableOffset = 128;
  view.setUint32(tagTableOffset, TAG_COUNT);

  // Tag 1: 'desc' (profile description)
  writeAscii(buf, tagTableOffset + 4, "desc");
  view.setUint32(tagTableOffset + 8, descOffset);
  view.setUint32(tagTableOffset + 12, descData.length);

  // Tag 2: 'wtpt' (media white point)
  writeAscii(buf, tagTableOffset + 16, "wtpt");
  view.setUint32(tagTableOffset + 20, wtptOffset);
  view.setUint32(tagTableOffset + 24, wtptData.length);

  // Tag 3: 'cprt' (copyright)
  writeAscii(buf, tagTableOffset + 28, "cprt");
  view.setUint32(tagTableOffset + 32, cprtOffset);
  view.setUint32(tagTableOffset + 36, cprtData.length);

  // --- Tag data ---
  buf.set(descData, descOffset);
  buf.set(wtptData, wtptOffset);
  buf.set(cprtData, cprtOffset);

  return buf;
}

/** Write a 4-character ASCII string at the given offset. */
function writeAscii(buf: Uint8Array, offset: number, str: string): void {
  for (let i = 0; i < str.length; i++) {
    buf[offset + i] = str.charCodeAt(i);
  }
}

/** Write an s15Fixed16Number (ICC fixed-point) at the given offset. */
function writeS15Fixed16(view: DataView, offset: number, value: number): void {
  const fixed = Math.round(value * 65536);
  view.setInt32(offset, fixed);
}

/**
 * Build an ICC 'desc' (textDescription) tag.
 * Type signature: 'desc', followed by ASCII string.
 */
function buildTextDescriptionTag(text: string): Uint8Array {
  // desc type: 4 (sig) + 4 (reserved) + 4 (ASCII count) + N (ASCII) + padding
  const asciiLen = text.length + 1; // include null terminator
  // Total: sig(4) + reserved(4) + count(4) + ascii(asciiLen) + unicode_count(4) + scriptcode_count(2) + scriptcode(67)
  // Simplified: just ASCII portion for minimal profile
  const totalLen = 4 + 4 + 4 + asciiLen + 4 + 4 + 2 + 67;
  const buf = new Uint8Array(totalLen);
  const view = new DataView(buf.buffer);

  writeAscii(buf, 0, "desc");
  // reserved: 0 (already zero)
  view.setUint32(8, asciiLen);
  for (let i = 0; i < text.length; i++) {
    buf[12 + i] = text.charCodeAt(i);
  }
  // null terminator at 12 + text.length is already 0
  // Unicode count = 0, scriptcode count = 0 — all zeros

  return buf;
}

/** Build an ICC 'XYZ ' tag for a single XYZ triplet. */
function buildXYZTag(x: number, y: number, z: number): Uint8Array {
  const buf = new Uint8Array(20); // sig(4) + reserved(4) + X(4) + Y(4) + Z(4)
  const view = new DataView(buf.buffer);

  writeAscii(buf, 0, "XYZ ");
  // reserved: 0
  writeS15Fixed16(view, 8, x);
  writeS15Fixed16(view, 12, y);
  writeS15Fixed16(view, 16, z);

  return buf;
}

/** Build an ICC 'text' tag. */
function buildTextTag(text: string): Uint8Array {
  const asciiLen = text.length + 1; // include null terminator
  const buf = new Uint8Array(4 + 4 + asciiLen); // sig(4) + reserved(4) + text
  writeAscii(buf, 0, "text");
  for (let i = 0; i < text.length; i++) {
    buf[8 + i] = text.charCodeAt(i);
  }
  return buf;
}

// =============================================================================
// XMP Metadata Writer
// =============================================================================

/**
 * Write a PDF/A-1b XMP metadata stream as an indirect object.
 *
 * The XMP packet contains:
 * - `dc:title` — document title
 * - `dc:creator` — document author
 * - `xmp:CreatorTool` — creating application
 * - `pdf:Producer` — PDF producer
 * - `pdfaid:part` — PDF/A part (1)
 * - `pdfaid:conformance` — PDF/A conformance level (B)
 *
 * @returns The object number of the XMP metadata stream.
 */
export function writePdfAMetadata(
  writer: PdfWriter,
  metadata: {
    title?: string;
    author?: string;
    subject?: string;
    creator?: string;
  }
): number {
  const now = new Date();
  const isoDate = now.toISOString().replace(/\.\d{3}Z$/, "Z");

  const title = escapeXml(metadata.title ?? "");
  const author = escapeXml(metadata.author ?? "");
  const subject = escapeXml(metadata.subject ?? "");
  const creator = escapeXml(metadata.creator ?? "excelts");
  const producer = "excelts";

  const xmp = [
    '<?xpacket begin="\uFEFF" id="W5M0MpCehiHzreSzNTczkc9d"?>',
    '<x:xmpmeta xmlns:x="adobe:ns:meta/">',
    '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">',
    '<rdf:Description rdf:about=""',
    '  xmlns:dc="http://purl.org/dc/elements/1.1/"',
    '  xmlns:xmp="http://ns.adobe.com/xap/1.0/"',
    '  xmlns:pdf="http://ns.adobe.com/pdf/1.3/"',
    '  xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/">',
    "  <dc:title>",
    '    <rdf:Alt><rdf:li xml:lang="x-default">' + title + "</rdf:li></rdf:Alt>",
    "  </dc:title>",
    "  <dc:creator>",
    "    <rdf:Seq><rdf:li>" + author + "</rdf:li></rdf:Seq>",
    "  </dc:creator>",
    "  <dc:description>",
    '    <rdf:Alt><rdf:li xml:lang="x-default">' + subject + "</rdf:li></rdf:Alt>",
    "  </dc:description>",
    "  <xmp:CreatorTool>" + creator + "</xmp:CreatorTool>",
    "  <xmp:CreateDate>" + isoDate + "</xmp:CreateDate>",
    "  <xmp:ModifyDate>" + isoDate + "</xmp:ModifyDate>",
    "  <pdf:Producer>" + producer + "</pdf:Producer>",
    "  <pdfaid:part>1</pdfaid:part>",
    "  <pdfaid:conformance>B</pdfaid:conformance>",
    "</rdf:Description>",
    "</rdf:RDF>",
    "</x:xmpmeta>",
    '<?xpacket end="w"?>'
  ].join("\n");

  const encoder = new TextEncoder();
  const xmpBytes = encoder.encode(xmp);

  const objNum = writer.allocObject();
  const dict = new PdfDict()
    .set("Type", "/Metadata")
    .set("Subtype", "/XML")
    .set("Length", pdfNumber(xmpBytes.length));

  // XMP metadata must NOT be compressed for PDF/A compliance
  // (and for general discoverability by search tools)
  writer.addStreamObject(objNum, dict, xmpBytes, { compress: false });
  return objNum;
}

// =============================================================================
// OutputIntent Writer
// =============================================================================

/**
 * Write a PDF/A-1b OutputIntent with an embedded sRGB ICC profile.
 *
 * Creates two objects:
 * 1. The ICC profile stream
 * 2. The OutputIntent dictionary referencing the profile
 *
 * @returns The object number of the OutputIntent dictionary.
 */
export function writePdfAOutputIntent(writer: PdfWriter): number {
  // Write ICC profile stream
  const iccObjNum = writer.allocObject();
  const iccDict = new PdfDict()
    .set("N", pdfNumber(3)) // 3 components (RGB)
    .set("Length", pdfNumber(sRGB_ICC_PROFILE.length));

  writer.addStreamObject(iccObjNum, iccDict, sRGB_ICC_PROFILE);

  // Write OutputIntent dictionary
  const intentObjNum = writer.allocObject();
  const intentDict = new PdfDict()
    .set("Type", "/OutputIntent")
    .set("S", "/GTS_PDFA1")
    .set("OutputConditionIdentifier", "(sRGB IEC61966-2.1)")
    .set("RegistryName", "(http://www.color.org)")
    .set("Info", "(sRGB IEC61966-2.1)")
    .set("DestOutputProfile", pdfRef(iccObjNum));

  writer.addObject(intentObjNum, intentDict);
  return intentObjNum;
}

// =============================================================================
// Helpers
// =============================================================================

/** Escape XML special characters for XMP content. */
function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
