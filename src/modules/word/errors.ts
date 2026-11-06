/**
 * DOCX module error types.
 *
 * All DOCX-related errors extend DocxError.
 */

import type { BaseErrorOptions } from "@utils/errors";
import { BaseError } from "@utils/errors";

/**
 * Base class for all DOCX-related errors.
 */
export class DocxError extends BaseError {
  constructor(message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.name = "DocxError";
  }
}

/**
 * Check if an error is a DOCX error.
 */
export function isDocxError(err: unknown): err is DocxError {
  return err instanceof DocxError;
}

/**
 * Error thrown when DOCX parsing fails.
 */
export class DocxParseError extends DocxError {
  override name = "DocxParseError";
}

/**
 * Error thrown when DOCX generation/writing fails.
 */
export class DocxWriteError extends DocxError {
  override name = "DocxWriteError";
}

/**
 * Error thrown when a builder or editor (`Document.*`, `Build.*`) is given a
 * value that has no valid form in the package — a font size that is `NaN`, a
 * heading level of 12, a colour written `"#FF0000"`, an underline style Word
 * does not define, `null` where text belongs, an index past the end. Thrown
 * by the call that received the value, before it changes anything, with a
 * message naming the function and the field.
 *
 * Also thrown while a document model is written (`Document.toBuffer`, `Document.writeFile`,
 * `Package.*`, the streaming writer) for a value the writer cannot write. That
 * check is shared by every writing entry point, so its message names no call.
 */
export class DocxArgumentError extends DocxError {
  override name = "DocxArgumentError";
  /**
   * Where the value at fault is, as a property path from the parameter name in
   * the public signature (`level`, `options.fontSize`, `properties.run.size`) —
   * or, for a value refused while a document is written, from the document root
   * (`body[0].properties.alignment`, `headers["rId7"].content.children[0]…`,
   * `glossary.blocks[0].content[0]…`) or from the streaming writer's `element`.
   * Following the path on the object you passed leads to the value. A number with
   * no XML form in an attribute the writer emits directly may be named by the
   * nearest object that holds it (`body[0].properties`), which is still a true
   * path, one step shallower; the message then names the XML attribute.
   */
  readonly field: string;

  constructor(field: string, message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.field = field;
  }
}

/**
 * Error thrown when the file system will not let `Document.readFile` / `Document.writeFile`
 * read or write a file — it does not exist, permission is denied, the directory
 * is missing. The system's own error is the `cause` (its `code`, `ENOENT` or
 * `EACCES`, is `error.cause.code`), as `ExcelFileError` keeps it. A file that was
 * read but is not a valid document is a {@link DocxParseError}, not this.
 */
export class DocxFileError extends DocxError {
  override name = "DocxFileError";

  constructor(
    public readonly path: string,
    public readonly operation: "read" | "write",
    details?: string,
    options?: BaseErrorOptions
  ) {
    super(
      details
        ? `Failed to ${operation} Word file "${path}": ${details}`
        : `Failed to ${operation} Word file "${path}"`,
      options
    );
  }
}

/**
 * Error thrown when an image cannot become a valid package part: its media
 * type is not an `ImageMediaType`, or its file name has no extension that
 * `[Content_Types].xml` can type. Thrown where the image is supplied and,
 * for a model built by hand, by the writers before any part is emitted.
 * `field` names what is at fault (`mediaType`, `images[0].fileName`,
 * `thumbnail.contentType`).
 */
export class DocxInvalidImageError extends DocxArgumentError {
  override name = "DocxInvalidImageError";
}

/**
 * Error thrown when a drawing — a shape, a picture, a text box or a chart
 * frame — is given a size, position, wrap distance or other numeric
 * property that has no valid form in the package: not a finite number, or
 * negative where the schema allows only zero and up. Thrown where the value
 * is supplied and, for a model built by hand, by the writers before the
 * value reaches a part.
 */
export class DocxInvalidDrawingError extends DocxArgumentError {
  override name = "DocxInvalidDrawingError";
}

/**
 * A collection whose ids content refers to: `w:footnoteReference`,
 * `w:endnoteReference`, `w:commentReference`, `w:abstractNumId`, `w:numId`
 * and an image's `r:embed`.
 */
export type DocxIdCollection =
  | "footnotes"
  | "endnotes"
  | "comments"
  | "abstractNumberings"
  | "numberingInstances"
  | "images";

/**
 * Error thrown when a document about to be written gives two entries of a
 * referenced collection the same id. Content refers to an entry by id alone,
 * so the writer cannot tell which entry a reference meant and refuses to
 * guess; writing it would produce a package with two `w:footnote w:id="3"`.
 *
 * The usual cause is an entry edited in place after a `Document.*` helper had
 * read the collection — see `DocxDocument` for the edits the helpers track.
 */
export class DocxDuplicateIdError extends DocxWriteError {
  override name = "DocxDuplicateIdError";
  /** The collection holding the duplicate. */
  readonly collection: DocxIdCollection;
  /** The id two entries share. An image's id is its `rId`. */
  readonly id: number | string;

  constructor(collection: DocxIdCollection, id: number | string, detail?: string) {
    super(
      `Duplicate id ${typeof id === "string" ? `"${id}"` : id} in ${collection}` +
        (detail ? ` (${detail})` : "") +
        ": references to it cannot be told apart. Give each entry its own id, or replace " +
        "the collection with a new array rather than editing an entry in place, so the " +
        "Document helpers can see the change."
    );
    this.collection = collection;
    this.id = id;
  }
}

/**
 * Error thrown when a required DOCX part is missing.
 */
export class DocxMissingPartError extends DocxParseError {
  override name = "DocxMissingPartError";

  constructor(partPath: string) {
    super(`Required DOCX part not found: ${partPath}`);
  }
}

/**
 * Error thrown when document structure is invalid.
 */
export class DocxInvalidStructureError extends DocxParseError {
  override name = "DocxInvalidStructureError";
}

/**
 * Error thrown when an unsupported feature is encountered during parsing.
 */
export class DocxUnsupportedFeatureError extends DocxError {
  override name = "DocxUnsupportedFeatureError";

  constructor(feature: string) {
    super(`Unsupported DOCX feature: ${feature}`);
  }
}

/**
 * Error thrown when a DOCX file is encrypted (CFB format) and no password
 * was provided. Pass `{ password }` in the read options — of `Document.read()`, a
 * `Package` function or `new Stream.DocumentReader()` — to decrypt automatically,
 * or call `WordCrypto.decrypt()` from `documonster/word/crypto` directly if you
 * need lower-level access to the decrypted ZIP bytes.
 */
export class DocxEncryptedError extends DocxError {
  override name = "DocxEncryptedError";

  constructor(options?: BaseErrorOptions) {
    super(
      "The document is encrypted (password-protected). " +
        "Pass { password } in the read options (Document.read(), Package.*, new Stream.DocumentReader()), " +
        'or use WordCrypto.decrypt() from "documonster/word/crypto" for lower-level decryption.',
      options
    );
  }
}

/**
 * Error thrown when an encrypted DOCX cannot be decrypted with the provided
 * password (wrong password) or when the encryption metadata is malformed.
 */
export class DocxDecryptionError extends DocxError {
  override name = "DocxDecryptionError";
  /**
   * Why, when the decryption knows: `"password"` — the password is wrong (ask again); `"unsupported"` — an
   * encryption other than Agile (Standard, RC4); `"malformed"` — the encryption metadata or package is damaged;
   * `"integrity"` — the package fails its data-integrity check (altered or damaged); `"limit"` — the package is
   * larger than the read allows. The same values `ExcelDecryptionError.reason` takes.
   */
  readonly reason?: "password" | "unsupported" | "malformed" | "integrity" | "limit";

  constructor(
    message: string,
    options?: BaseErrorOptions & {
      readonly reason?: "password" | "unsupported" | "malformed" | "integrity" | "limit";
    }
  ) {
    super(message, options);
    if (options?.reason !== undefined) {
      this.reason = options.reason;
    }
  }
}

/**
 * Error thrown when an input package exceeds a declared resource limit (e.g.
 * total package size, single part size, number of parts, XML nesting depth). Used to defend
 * against ZIP bombs and runaway memory usage.
 */
export class DocxLimitExceededError extends DocxParseError {
  override name = "DocxLimitExceededError";
  /**
   * What was exceeded: `"packageSize"` (`securityPolicy.maxPackageSize`), `"partSize"` (`maxPartSize`),
   * `"partCount"` (`maxPartCount`) or `"xmlDepth"` (`maxXmlDepth`) — options of the `securityPolicy`.
   */
  readonly limit: "packageSize" | "partSize" | "partCount" | "xmlDepth";
  /** The configured maximum — named as the other modules' limit errors name it. */
  readonly allowed: number;
  /** @deprecated Use {@link allowed}. Removed in 1.0. */
  readonly maximum: number;
  /** Actual measured value (or the value that would have been reached). */
  readonly actual: number;

  constructor(
    limit: "packageSize" | "partSize" | "partCount" | "xmlDepth",
    maximum: number,
    actual: number,
    detail?: string
  ) {
    super(
      `DOCX ${limit} limit exceeded: actual ${actual} > maximum ${maximum}` +
        (detail ? ` (${detail})` : "")
    );
    this.limit = limit;
    this.allowed = maximum;
    this.maximum = maximum;
    this.actual = actual;
  }
}

/**
 * Error thrown when the active `DocxSecurityPolicy.rawXmlPolicy` is `"reject"`
 * and the writer encounters opaque/preserved rawXml that would otherwise have
 * been emitted verbatim. Callers can either drop the offending model field
 * before serialisation or relax the policy to `"strip"` / `"preserve"`.
 */
export class DocxRawXmlPolicyError extends DocxWriteError {
  override name = "DocxRawXmlPolicyError";
  /** Short label identifying the offending writer site. */
  readonly site: string;

  constructor(site: string) {
    super(
      `Raw XML rejected by security policy at "${site}". ` +
        "Either remove the rawXml/opaque field from the model or set " +
        "DocxSecurityPolicy.rawXmlPolicy to 'preserve' or 'strip'."
    );
    this.site = site;
  }
}

/** Error thrown when template processing fails. */
export class DocxTemplateError extends DocxError {
  override readonly name = "DocxTemplateError";
  /** The placeholder that caused the error. */
  readonly placeholder: string;
  /** Location hint (e.g. "body paragraph 3", "header"). */
  readonly location: string;
  /** The tag name/expression that caused the error. */
  readonly tagName?: string;
  /** The paragraph index within the current block where the error occurred. */
  readonly paragraphIndex?: number;
  /** The section path (e.g. "body", "header:default", "footer:first"). */
  readonly sectionPath?: string;

  constructor(
    message: string,
    placeholder: string,
    location: string,
    options?: BaseErrorOptions & {
      tagName?: string;
      paragraphIndex?: number;
      sectionPath?: string;
    }
  ) {
    super(message, options);
    this.placeholder = placeholder;
    this.location = location;
    this.tagName = options?.tagName;
    this.paragraphIndex = options?.paragraphIndex;
    this.sectionPath = options?.sectionPath;
  }
}

/** @deprecated Use {@link DocxTemplateError}. Removed in 1.0. */
export const TemplateError = DocxTemplateError;
/** @deprecated Use {@link DocxTemplateError}. Removed in 1.0. */
export type TemplateError = DocxTemplateError;
