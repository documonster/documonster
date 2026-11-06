/**
 * `Pdf.Signature` — the low-level pieces of PDF signing and verification, for a caller
 * assembling or checking a signed file by hand.
 *
 * Most callers want `Pdf.Builder#sign` or `Pdf.Editor#sign`, which write a signature
 * in one step, and `Pdf.verifySignatures`, which finds and checks every signature in a
 * file. These are the building blocks those use: the signature dictionary with its
 * `/Contents` placeholder, filling that placeholder in, checking one signature whose
 * bytes the caller has located, and the DER decoder used to read a CMS structure back.
 */

import { at, isByteArray, pdfBytes as readPdfBytes, readerBytes } from "@pdf/core/arguments";
import { checkSigningMaterial } from "@pdf/core/asn1";
import type { Asn1Node } from "@pdf/core/asn1";
import {
  asn1Parse,
  buildSignatureDictPlaceholder,
  signPdf,
  verifyPdfSignature
} from "@pdf/core/digital-signature";
import type { PdfSignatureCheck, SignPdfOptions } from "@pdf/core/digital-signature";
import { PdfSignaturePlaceholderMissingError } from "@pdf/errors";
import { asPlainBytes } from "@utils/binary";
import { check, s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

const derBytes = () => s.custom<Uint8Array>(isByteArray, "a Uint8Array");

const signerSchema = s.object(
  {
    name: s.string().optional(),
    reason: s.string().optional(),
    location: s.string().optional(),
    contactInfo: s.string().optional()
  },
  { expected: "an options object" }
);
const _signerMatch: SchemaMatches<
  typeof signerSchema,
  NonNullable<Parameters<typeof buildDictPlaceholder>[0]>
> = true;
void _signerMatch;

const signPdfOptionsSchema = s.object(
  { certificate: derBytes(), privateKey: derBytes() },
  { expected: "an options object" }
);
const _signPdfOptionsMatch: SchemaMatches<typeof signPdfOptionsSchema, SignPdfOptions> = true;
void _signPdfOptionsMatch;

const SIGNATURE_HEX = "the hex string of /Contents, without its angle brackets";
const signatureHexSchema = s
  .string({ expected: SIGNATURE_HEX })
  .refine(v => /^[0-9a-f]*$/i.test(v), SIGNATURE_HEX);

// A `NaN` offset used to be reported as "PDF content was modified after signing".
const byteRangeSchema = s.custom<readonly [number, number, number, number]>(
  v =>
    Array.isArray(v) && v.length === 4 && v.every(value => Number.isInteger(value) && value >= 0),
  "four integers >= 0, [offset1, length1, offset2, length2]"
);

/**
 * The signature dictionary — `/Contents` zero-filled placeholder and a `/ByteRange` to patch —
 * as a string to splice into a file, with the placeholder it contains.
 *
 * @throws {PdfArgumentError} `options` is not an object, or a signer field is not a string.
 */
export function buildDictPlaceholder(options?: {
  name?: string;
  reason?: string;
  location?: string;
  contactInfo?: string;
}): { dictString: string; placeholder: string } {
  const where = "Pdf.Signature.buildDictPlaceholder";
  check(signerSchema.optional(), options, at(where, "options"));
  return buildSignatureDictPlaceholder(options);
}

/**
 * Decode one DER element — the codec the signature code reads a CMS structure with.
 *
 * `data` is a `Uint8Array`, another `ArrayBufferView` or an `ArrayBuffer` (viewed, not copied).
 *
 * @throws {PdfArgumentError} `data` is not bytes.
 * @throws {PdfStructureError} the bytes are not DER: truncated, or an indefinite length.
 */
export function parseAsn1(data: Uint8Array | ArrayBuffer | ArrayBufferView): Asn1Node {
  return asn1Parse(readerBytes("Pdf.Signature.parseAsn1", "data", data, "DER bytes"));
}

/**
 * Check one signature whose bytes the caller has located: the hex `/Contents` and the
 * `/ByteRange` it was computed around.
 *
 * The result uses the vocabulary of `Pdf.verifySignatures`' reports — `intact`,
 * `failure` — so a caller moving between the two reads the same field names. `pdfData` is
 * the file's bytes: a `Uint8Array`, another `ArrayBufferView` or an `ArrayBuffer`.
 *
 * @throws {PdfArgumentError} `pdfData` is not bytes, `signatureHex` is not hex, or `byteRange`
 *   is not four integers >= 0.
 */
export async function verify(
  pdfData: Uint8Array | ArrayBuffer | ArrayBufferView,
  signatureHex: string,
  byteRange: readonly [number, number, number, number]
): Promise<PdfSignatureCheck> {
  const where = "Pdf.Signature.verify";
  const data = readPdfBytes(where, "pdfData", pdfData);
  check(signatureHexSchema, signatureHex, at(where, "signatureHex"));
  check(byteRangeSchema, byteRange, at(where, "byteRange"));
  const result = await verifyPdfSignature(
    data,
    signatureHex,
    byteRange as [number, number, number, number]
  );
  return {
    intact: result.valid,
    coversWholeFile: result.coversWholeFile,
    digestAlgorithm: result.digestAlgorithm,
    ...(result.reason === undefined ? {} : { failure: result.reason })
  };
}

/**
 * Fill in the CMS signature of a PDF whose signature dictionary — `/ByteRange` and a
 * zero-filled `/Contents` placeholder, as {@link buildDictPlaceholder} writes it — is
 * already in the file.
 *
 * This is only the last step of signing. It does not add a signature field: a file
 * without a placeholder is rejected. `Pdf.Builder#sign` and `Pdf.Editor#sign` write the
 * placeholder and fill it in one call. `pdfBytes` is a `Uint8Array`, another
 * `ArrayBufferView` or an `ArrayBuffer`.
 *
 * @throws {PdfArgumentError} the bytes or options are not what they must be, or the
 *   certificate or key is not DER X.509 / PKCS#8 RSA.
 * @throws {PdfSignaturePlaceholderMissingError} when the file holds no signature placeholder
 *   (a `PdfStructureError`).
 */
export async function fillPlaceholder(
  pdfBytes: Uint8Array | ArrayBuffer | ArrayBufferView,
  options: SignPdfOptions
): Promise<Uint8Array<ArrayBuffer>> {
  const where = "Pdf.Signature.fillPlaceholder";
  const data = readPdfBytes(where, "pdfBytes", pdfBytes);
  check(signPdfOptionsSchema, options, at(where, "options"));
  checkSigningMaterial(where, options.certificate, options.privateKey);
  try {
    return asPlainBytes(await signPdf(data, options.certificate, options.privateKey));
  } catch (err) {
    if (err instanceof PdfSignaturePlaceholderMissingError) {
      throw new PdfSignaturePlaceholderMissingError(
        "Pdf.Signature.fillPlaceholder: no signature placeholder (/ByteRange with a /Contents " +
          "hex string) found. Write one first — Pdf.Builder#sign and Pdf.Editor#sign do both " +
          "steps; Pdf.Signature.buildDictPlaceholder builds the dictionary.",
        { cause: err }
      );
    }
    throw err;
  }
}
