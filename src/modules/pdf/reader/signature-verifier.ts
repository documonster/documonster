/**
 * Find and verify every digital signature in a PDF.
 *
 * `verifyPdfSignature` checks one signature whose `/Contents` and `/ByteRange` the
 * caller has already dug out of the file. This does the digging: it walks the AcroForm
 * field tree through the reader, so a signature is found by what the document says it
 * is — a `/FT /Sig` field whose value is a signature dictionary — rather than by
 * searching the bytes for something that looks like one.
 *
 * What it establishes is integrity, not trust: see {@link PdfSignatureReport.intact}.
 */

import { at, isBlob, pdfBytesOrBlob } from "@pdf/core/arguments";
import {
  certificateCommonName,
  certificateValidity,
  parseCmsSignedData,
  verifyPdfSignatureBytes
} from "@pdf/core/digital-signature";
import { PdfPasswordError } from "@pdf/errors";
import { parsePdfDate } from "@pdf/reader/metadata-reader";
import { initDecryption, isEncrypted } from "@pdf/reader/pdf-decrypt";
import { PdfDocument } from "@pdf/reader/pdf-document";
import { decodePdfStringBytes, dictGetName, isPdfArray, isPdfRef } from "@pdf/reader/pdf-parser";
import type { PdfDictValue, PdfObject } from "@pdf/reader/pdf-parser";
import { asPlainBytes } from "@utils/binary";
import { check, s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

/** Options for `Pdf.verifySignatures`. */
export interface VerifySignaturesOptions {
  /** Password of an encrypted document. */
  password?: string;
}

/**
 * One signature found in a PDF, and what can be established about it from the file
 * alone.
 *
 * **No trust is evaluated.** {@link intact} says the signed bytes are unchanged and the
 * signature verifies with the public key of the certificate the signature itself
 * carries — which anyone can generate under any name. Nothing here builds a chain to a
 * trusted root, checks revocation (CRL/OCSP), honours a timestamp token, or checks the
 * certificate's key usage. To decide whether to trust the signer, evaluate
 * {@link certificate} against your own trust store; {@link claimedSigner} and the
 * dictionary entries below are unauthenticated claims until you have.
 *
 * **Later revisions.** A signature covers the file as it was when it was signed. Bytes
 * appended after it — an incremental update — are not covered, and are reported only
 * through {@link coversWholeFile}. Whether such an update was _permitted_ (by `/DocMDP`
 * or `/FieldMDP` transform parameters) is not evaluated: an `intact` signature on an
 * earlier revision says nothing about what the update changed.
 */
export interface PdfSignatureReport {
  /** Fully qualified name of the signature field (`parent.child`). */
  fieldName: string;
  /**
   * The signed byte ranges, `[offset1, length1, offset2, length2]`, as the signature
   * dictionary declares them.
   */
  byteRange: [number, number, number, number];
  /**
   * Whether `/ByteRange` is well formed: it starts at byte 0, its two ranges lie inside
   * the file, and the gap between them holds exactly this signature's `/Contents` hex
   * string — so no byte of the revision is left unsigned. A signature whose ranges fail
   * this is reported not {@link intact} even if its digest matches.
   */
  byteRangeValid: boolean;
  /**
   * Whether the signature is cryptographically intact: `/ByteRange` is well formed, the
   * digest of the signed bytes matches the one the signature carries, and the signature
   * over it verifies with the public key of the embedded certificate the CMS SignerInfo
   * names. This is integrity only — see the type's description for what it does not
   * establish.
   */
  intact: boolean;
  /** Whether the signed byte ranges reach the end of the file (no later revision). */
  coversWholeFile: boolean;
  /** Digest algorithm OID the signature uses; empty when it could not be read. */
  digestAlgorithm: string;
  /** Why the signature is not {@link intact}, when it is not. */
  failure?: string;
  /**
   * The signer's certificate (DER), as embedded in the signature — the one its SignerInfo
   * names. Unverified: evaluate it against your trust store to authenticate the signer.
   */
  certificate?: Uint8Array<ArrayBuffer>;
  /**
   * Common name from {@link certificate}. A claim, not an identity: a self-signed
   * certificate can carry any name.
   */
  claimedSigner?: string;
  /**
   * The validity period {@link certificate} declares. Reported, not judged: whether it
   * should be compared with {@link signingTime} (itself a claim), a timestamp or the
   * current time is a trust-policy decision.
   */
  certificateValidity?: { notBefore: Date; notAfter: Date };
  /** `/Name` from the signature dictionary. */
  name?: string;
  /** `/Reason` from the signature dictionary. */
  reason?: string;
  /** `/Location` from the signature dictionary. */
  location?: string;
  /** `/ContactInfo` from the signature dictionary. */
  contactInfo?: string;
  /** `/M` from the signature dictionary — the signing time the signer claims. */
  signingTime?: Date;
  /** `/SubFilter`, e.g. `adbe.pkcs7.detached`. */
  subFilter?: string;
}

const verifySignaturesOptionsSchema = s.object(
  { password: s.string().optional() },
  { expected: "an options object" }
);
const _verifySignaturesOptionsMatch: SchemaMatches<
  typeof verifySignaturesOptionsSchema,
  VerifySignaturesOptions
> = true;
void _verifySignaturesOptionsMatch;

/**
 * Find every signature field in a PDF and check each one's integrity.
 *
 * Returns one report per signature, in field order; an unsigned document yields `[]`.
 * A signature that is intact but does not cover the whole file (`coversWholeFile:
 * false`) signed an earlier revision, and the bytes after it are an unsigned update.
 * No trust is evaluated — see {@link PdfSignatureReport}.
 *
 * `data` is the file's bytes: a `Uint8Array`, another `ArrayBufferView` or an `ArrayBuffer`
 * (viewed, not copied) — or a `Blob` (a browser `File`), whose bytes are read first.
 *
 * @throws {PdfArgumentError} `data` is none of these, or `options` is not an options object.
 * @throws {PdfPasswordError} when the document is encrypted and the password is wrong (a
 *   `PdfStructureError`).
 */
export async function verifyPdfSignatures(
  data: Uint8Array | ArrayBuffer | ArrayBufferView | Blob,
  options: VerifySignaturesOptions = {}
): Promise<PdfSignatureReport[]> {
  const source = pdfBytesOrBlob("Pdf.verifySignatures", "data", data);
  check(verifySignaturesOptionsSchema, options, at("Pdf.verifySignatures", "options"));
  const bytes = isBlob(source) ? new Uint8Array(await source.arrayBuffer()) : source;
  const doc = new PdfDocument(bytes);
  if (isEncrypted(doc) && !initDecryption(doc, options.password ?? "")) {
    throw new PdfPasswordError("Failed to decrypt PDF: incorrect password");
  }

  const reports: PdfSignatureReport[] = [];
  for (const { fieldName, value } of findSignatureValues(doc)) {
    reports.push(await verifyOne(bytes, fieldName, value));
  }
  return reports;
}

interface SignatureValue {
  fieldName: string;
  value: PdfDictValue;
}

/** Walk the AcroForm tree, inheriting /FT and building qualified names. */
function findSignatureValues(doc: PdfDocument): SignatureValue[] {
  const found: SignatureValue[] = [];
  let acroForm: PdfDictValue | null;
  try {
    acroForm = doc.derefDict(doc.getCatalog().get("AcroForm") ?? null);
  } catch {
    return found;
  }
  if (!acroForm) {
    return found;
  }
  const fields = doc.deref(acroForm.get("Fields") ?? null);
  if (!isPdfArray(fields)) {
    return found;
  }

  const visited = new Set<string>();
  const seenValues = new Set<PdfDictValue>();
  // Depth-first with an explicit stack, so a field tree as deep as the file cares to
  // make it cannot exhaust the call stack. Kids are pushed in reverse to keep field order.
  const stack: Array<{ item: PdfObject; inheritedFt: string | undefined; prefix: string }> = [];
  for (let i = fields.length - 1; i >= 0; i--) {
    stack.push({ item: fields[i], inheritedFt: undefined, prefix: "" });
  }
  while (stack.length > 0) {
    const { item, inheritedFt, prefix } = stack.pop()!;
    if (isPdfRef(item)) {
      const key = `${item.objNum}:${item.gen}`;
      if (visited.has(key)) {
        continue;
      }
      visited.add(key);
    }
    const field = doc.derefDict(item);
    if (!field) {
      continue;
    }
    const partial = readText(doc, field.get("T"));
    const name = partial === undefined ? prefix : prefix ? `${prefix}.${partial}` : partial;
    const ft = dictGetName(field, "FT") ?? inheritedFt;
    if (ft === "Sig") {
      const value = doc.derefDict(field.get("V") ?? null);
      // A field and its widget can both point at one dictionary; it is one signature.
      if (value && value.has("ByteRange") && value.has("Contents") && !seenValues.has(value)) {
        seenValues.add(value);
        found.push({ fieldName: name, value });
      }
    }
    const kids = doc.deref(field.get("Kids") ?? null);
    if (isPdfArray(kids)) {
      for (let i = kids.length - 1; i >= 0; i--) {
        stack.push({ item: kids[i], inheritedFt: ft, prefix: name });
      }
    }
  }
  return found;
}

async function verifyOne(
  data: Uint8Array,
  fieldName: string,
  value: PdfDictValue
): Promise<PdfSignatureReport> {
  const rawRange = value.get("ByteRange");
  const numbers = isPdfArray(rawRange)
    ? rawRange.filter((n): n is number => typeof n === "number")
    : [];
  const byteRange: [number, number, number, number] =
    numbers.length === 4 ? [numbers[0], numbers[1], numbers[2], numbers[3]] : [0, 0, 0, 0];

  const details = {
    fieldName,
    byteRange,
    ...describe(value)
  };
  const broken = (failure: string): PdfSignatureReport => ({
    ...details,
    byteRangeValid: false,
    intact: false,
    coversWholeFile: false,
    digestAlgorithm: "",
    failure
  });

  const gap = numbers.length === 4 ? signatureGap(data, byteRange) : undefined;
  if (gap === undefined) {
    return broken("Malformed /ByteRange — the signed ranges do not enclose the /Contents value");
  }
  if (gap === TOO_LARGE) {
    return broken(
      `/ByteRange gap is larger than ${MAX_SIGNATURE_HEX_DIGITS} hex digits — not a signature`
    );
  }

  // The gap is what the signature was computed around; it must be this dictionary's
  // /Contents, or the ranges were pointed at some other string's bytes. A signature's
  // /Contents is never encrypted (ISO 32000-1 §7.6.1), and the reader leaves it as
  // written, so this holds for an encrypted file too.
  const contents = value.get("Contents");
  if (!(contents instanceof Uint8Array) || !sameSignature(contents, gap)) {
    return broken("/ByteRange does not enclose this signature's /Contents");
  }

  const result = await verifyPdfSignatureBytes(data, gap, byteRange);
  let certificate: Uint8Array<ArrayBuffer> | undefined;
  try {
    certificate = asPlainBytes(parseCmsSignedData(gap).certificate);
  } catch {
    certificate = undefined;
  }
  const claimedSigner = certificate && certificateCommonName(certificate);
  const validity = certificate && certificateValidity(certificate);
  return {
    ...details,
    byteRangeValid: true,
    intact: result.valid,
    coversWholeFile: result.coversWholeFile,
    digestAlgorithm: result.digestAlgorithm,
    ...(result.reason === undefined ? {} : { failure: result.reason }),
    ...(certificate ? { certificate } : {}),
    ...(claimedSigner === undefined ? {} : { claimedSigner }),
    ...(validity ? { certificateValidity: validity } : {})
  };
}

/**
 * Most hex digits a signature gap may hold. A CMS signature with embedded revocation
 * data runs to tens of kilobytes; this leaves two orders of magnitude of room while
 * refusing to decode a "signature" that is most of a large file.
 */
const MAX_SIGNATURE_HEX_DIGITS = 8 * 1024 * 1024;
const TOO_LARGE = Symbol("too large");

/** PDF white space (ISO 32000-1 Table 1), which a hex string may contain anywhere. */
function isPdfWhiteSpace(c: number): boolean {
  return c === 0x00 || c === 0x09 || c === 0x0a || c === 0x0c || c === 0x0d || c === 0x20;
}

/**
 * The hex digits between the two signed ranges, white space removed, or `undefined`
 * when the ranges are not the shape a signature needs: starting at 0, inside the file,
 * in order, and separated by exactly one `<…>` hex string.
 */
function signatureGap(
  data: Uint8Array,
  [off1, len1, off2, len2]: [number, number, number, number]
): Uint8Array | typeof TOO_LARGE | undefined {
  const integers = [off1, len1, off2, len2].every(n => Number.isSafeInteger(n) && n >= 0);
  if (!integers || off1 !== 0 || off1 + len1 >= off2 || off2 + len2 > data.length) {
    return undefined;
  }
  const start = off1 + len1;
  const end = off2;
  if (data[start] !== 0x3c || data[end - 1] !== 0x3e) {
    return undefined;
  }
  if (end - start - 2 > MAX_SIGNATURE_HEX_DIGITS * 2) {
    return TOO_LARGE;
  }
  // Decoded straight into bytes: the gap can be megabytes of hex, and a string per digit
  // then a join and a re-parse cost several times the work.
  const bytes = new Uint8Array(Math.ceil((end - start - 2) / 2));
  let digits = 0;
  for (let i = start + 1; i < end - 1; i++) {
    const c = data[i];
    if (isPdfWhiteSpace(c)) {
      continue;
    }
    const nibble = hexNibble(c);
    if (nibble < 0) {
      return undefined;
    }
    if (digits % 2 === 0) {
      bytes[digits >> 1] = nibble << 4;
    } else {
      bytes[digits >> 1] |= nibble;
    }
    digits++;
  }
  if (digits > MAX_SIGNATURE_HEX_DIGITS) {
    return TOO_LARGE;
  }
  // An odd final digit is followed by an implied 0 (§7.3.4.3), which the shift supplied.
  return bytes.subarray(0, Math.ceil(digits / 2));
}

function hexNibble(c: number): number {
  if (c >= 0x30 && c <= 0x39) {
    return c - 0x30;
  }
  if (c >= 0x41 && c <= 0x46) {
    return c - 0x37;
  }
  if (c >= 0x61 && c <= 0x66) {
    return c - 0x57;
  }
  return -1;
}

/** Compare ignoring trailing zero padding, which a placeholder leaves behind. */
function sameSignature(a: Uint8Array, b: Uint8Array): boolean {
  let lengthA = a.length;
  let lengthB = b.length;
  while (lengthA > 0 && a[lengthA - 1] === 0) {
    lengthA--;
  }
  while (lengthB > 0 && b[lengthB - 1] === 0) {
    lengthB--;
  }
  if (lengthA !== lengthB) {
    return false;
  }
  for (let i = 0; i < lengthA; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

function describe(
  value: PdfDictValue
): Pick<
  PdfSignatureReport,
  "name" | "reason" | "location" | "contactInfo" | "signingTime" | "subFilter"
> {
  const text = (key: string): string | undefined => {
    const raw = value.get(key);
    return raw instanceof Uint8Array ? decodePdfStringBytes(raw) : undefined;
  };
  const out: ReturnType<typeof describe> = {};
  const name = text("Name");
  const reason = text("Reason");
  const location = text("Location");
  const contactInfo = text("ContactInfo");
  const m = text("M");
  const signingTime = m === undefined ? null : parsePdfDate(m);
  const subFilter = dictGetName(value, "SubFilter");
  if (name !== undefined) {
    out.name = name;
  }
  if (reason !== undefined) {
    out.reason = reason;
  }
  if (location !== undefined) {
    out.location = location;
  }
  if (contactInfo !== undefined) {
    out.contactInfo = contactInfo;
  }
  if (signingTime) {
    out.signingTime = signingTime;
  }
  if (subFilter !== undefined) {
    out.subFilter = subFilter;
  }
  return out;
}

function readText(doc: PdfDocument, raw: PdfObject | undefined): string | undefined {
  const value = doc.deref(raw ?? null);
  if (value instanceof Uint8Array) {
    return decodePdfStringBytes(value);
  }
  return typeof value === "string" ? value : undefined;
}
