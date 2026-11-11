/**
 * DOCX Incremental Edit API
 *
 * Supports efficient "open → modify → save" workflows by editing a DOCX
 * file at the ZIP entry level.
 *
 * Use cases:
 * - Replace just the document.xml in an existing template
 * - Update headers/footers without touching styles or images
 * - Patch specific parts (e.g. core properties metadata)
 *
 * **Limitations and caveats — read before using `replacePart`/`deletePart`:**
 *
 * This API is a thin wrapper around ZIP entry mutation. It deliberately
 * does not understand DOCX semantics:
 *
 * - `[Content_Types].xml` is **not** given the real type of a new part
 *   with a previously-unseen content type. If your replacement part
 *   introduces a new MIME type, callers must also patch `[Content_Types].xml` (or
 *   round-trip through `readDocx` + `packageDocx` instead). A new part the
 *   final `[Content_Types].xml` does not type is given the generic
 *   `application/octet-stream` (a `Default` for its extension), except a
 *   media part (`word/media/…`), which is rejected with a `DocxWriteError`
 *   because a picture typed generically would not render.
 * - Relationship files (`*.rels`) are **not** rebuilt. Adding or deleting
 *   a part with relationships will leave the corresponding `_rels/*.rels`
 *   stale.
 * - `replaceBody`, `replaceHeader`, `replaceFooter` only rewrite their
 *   target XML part; they do **not** modify the corresponding `.rels` file
 *   or copy in new media binaries. As a result they reject content that
 *   would introduce new relationship IDs (images, hyperlinks, charts,
 *   altChunks, opaque drawings carrying rIds) — use `patchDocument` /
 *   `packageDocx` for those cases. Existing references that survive
 *   unchanged are fine because their `.rels` entries already exist.
 * - A comment that a `replaceBody`, `replaceHeader` or `replaceFooter` takes
 *   the last anchor of — it was anchored in the replaced part and is anchored
 *   nowhere afterwards — is removed from `comments.xml`, together with its
 *   `commentsExtended` / `commentsIds` / `commentsExtensible` entries. A
 *   comment that had no anchor before the edit is left as it is.
 *
 * For DOCX-aware edits (changes that may add/remove relationships or
 * content types) use `patchDocument` or `compileTemplate`, or round-trip
 * through `readDocx` + `packageDocx`.
 *
 * Note: This API operates on raw ZIP entries; it does not parse the DOCX
 * model. For higher-level edits use `patchDocument` or `compileTemplate`.
 */

import { createZip } from "@archive/create-archive";
import { unzipZip } from "@archive/read-archive";
import { asPlainBytes } from "@utils/binary";
import type { SchemaMatches } from "@utils/schema";
import { xmlEncodeAttr } from "@utils/xml-encode";
import { PartPath } from "@word/constants";
import { isFlatOpc } from "@word/convert/flat-opc";
import {
  checkArray,
  checkBytes,
  checkNodeArray,
  checkOneOf,
  checkRecord,
  checkString,
  packageOptionsShape
} from "@word/core/argument-checks";
import { utf8Decoder, utf8Encoder } from "@word/core/internal-utils";
import { checkSchema, w } from "@word/core/option-schema";
import { walkBlocks } from "@word/core/walker";
import {
  DocxError,
  DocxParseError,
  DocxUnsupportedFeatureError,
  DocxWriteError
} from "@word/errors";
import { checkReadOptions } from "@word/reader/docx-reader";
import type { ReadDocxOptions } from "@word/reader/docx-reader";
import {
  checkPackageSize,
  createPartBudget,
  flatOpcEntries,
  readOptionsShape
} from "@word/reader/package-limits";
import { asPartLimitError } from "@word/reader/part-xml";
import { decryptIfEncrypted } from "@word/security/package-encryption";
import type { DocxSecurityPolicy } from "@word/security/policy";
import type { BodyContent, DocxDocument, Paragraph, Table } from "@word/types";
import { assertValidPartName } from "@word/writer/content-types";
import { renderBodyContent, renderDocument } from "@word/writer/document-writer";
import { renderHeader, renderFooter } from "@word/writer/header-footer-writer";
import type { IdGenerators } from "@word/writer/render-context";
import {
  closeStoryPart,
  collectBookmarkStartIds,
  collectDocPrIds,
  collectRangeStartIds,
  collectRevisionIds,
  createIdGenerators,
  createRenderContext,
  enterRevisionIds,
  exitRevisionIds,
  renderHelpersFor
} from "@word/writer/render-context";
import { SaxParser } from "@xml/sax";
import { XmlWriter } from "@xml/writer";

// =============================================================================
// Public API
// =============================================================================

/** A single incremental edit operation. */
export type IncrementalEdit =
  | {
      /**
       * Replace a specific part by raw bytes.
       *
       * Caveat: does not update `[Content_Types].xml` or any `.rels` file.
       * Adding a part with new relationships produces a malformed package. A
       * new part `[Content_Types].xml` does not type gets
       * `application/octet-stream`, and a new media part is rejected. See the
       * file header for details.
       */
      readonly type: "replacePart";
      readonly path: string;
      readonly data: Uint8Array;
    }
  | {
      /** Replace a part with a string (UTF-8 encoded). Same caveats as `replacePart`. */
      readonly type: "replacePartText";
      readonly path: string;
      readonly text: string;
    }
  | {
      /**
       * Delete a specific part.
       *
       * Caveat: does not update `[Content_Types].xml` or any `.rels` file.
       * Deleting a part that other parts reference (via rels or other XML)
       * will leave dangling references — round-trip through readDocx +
       * packageDocx if you need a consistent package.
       */
      readonly type: "deletePart";
      readonly path: string;
    }
  | {
      /** Replace the body content of `word/document.xml`. */
      readonly type: "replaceBody";
      readonly body: readonly BodyContent[];
    }
  | {
      /** Replace a header part identified by path (e.g. `word/header1.xml`). */
      readonly type: "replaceHeader";
      readonly path: string;
      readonly children: readonly (Paragraph | Table)[];
    }
  | {
      /** Replace a footer part identified by path (e.g. `word/footer1.xml`). */
      readonly type: "replaceFooter";
      readonly path: string;
      readonly children: readonly (Paragraph | Table)[];
    };

/**
 * Options for `Package.edit`: how the file is read, as `Document.read` reads it
 * (`password`, `securityPolicy`, `checkCrc32`), and how the result is written.
 * The result of an encrypted file is not encrypted: `WordCrypto.encrypt` it again.
 */
export interface IncrementalEditOptions extends ReadDocxOptions {
  /** Compression level for replaced entries (0-9). Default: 6. */
  readonly compressionLevel?: number;
  /** Timestamp for every entry of the written package, as in `Document.toBuffer`. Default: now. */
  readonly modTime?: Date;
}

const incrementalEditOptionsSchema = w.record({
  ...readOptionsShape,
  ...packageOptionsShape
});
const _incrementalEditOptionsMatch: SchemaMatches<
  typeof incrementalEditOptionsSchema,
  IncrementalEditOptions
> = true;
void _incrementalEditOptionsMatch;

/**
 * Apply incremental edits to an existing DOCX file.
 *
 * Reads all parts from the original ZIP, applies the edits, and writes a
 * new ZIP. Unchanged parts are passed through (decompressed and recompressed,
 * but their content is not parsed).
 *
 * @param buffer - The original DOCX file as a Uint8Array.
 * @param edits - Array of incremental edits to apply.
 * @param options - Optional settings.
 * @returns A new DOCX file as a Uint8Array.
 *
 * @example
 * ```ts
 * const original = await fs.readFile("template.docx");
 * const edited = await Package.edit(original, [
 *   { type: "replaceBody", body: [...newParagraphs] }
 * ]);
 * ```
 */
export async function editDocxIncremental(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  edits: readonly IncrementalEdit[],
  options?: IncrementalEditOptions
): Promise<Uint8Array<ArrayBuffer>> {
  const where = "Package.edit";
  const bytes = checkPackageBytes(buffer, where);
  checkIncrementalEdits(edits, where);
  checkSchema(incrementalEditOptionsSchema.optional(), options, where);
  const policy = checkReadOptions(options, where);
  // Read all entries from the original ZIP. Some third-party ZIP writers emit
  // entry paths with a leading "/" — strip it so lookups by canonical path
  // (e.g. "word/document.xml") match consistently with the rest of the
  // codebase (see readDocxPart below).
  const parts = new Map<string, Uint8Array>();
  await readPackageEntries(bytes, options, policy, async (path, entry) => {
    parts.set(path, await entry.bytes());
    return false;
  });
  const originalPaths = new Set(parts.keys());

  // Apply edits. The comments the replaced story parts anchored are noted
  // first: those this edit leaves without an anchor are removed below.
  const replacedAnchors = new Set<string>();
  for (const edit of edits) {
    if (
      edit.type === "replaceBody" ||
      edit.type === "replaceHeader" ||
      edit.type === "replaceFooter"
    ) {
      const replaced = parts.get(
        edit.type === "replaceBody" ? PartPath.Document : edit.path.replace(/^\/+/, "")
      );
      if (replaced) {
        collectCommentAnchors(utf8Decoder.decode(replaced), replacedAnchors);
      }
    }
    switch (edit.type) {
      case "replacePart":
        parts.set(edit.path.replace(/^\/+/, ""), checkBytes(edit.data, "data", where));
        break;
      case "replacePartText":
        parts.set(edit.path.replace(/^\/+/, ""), utf8Encoder.encode(edit.text));
        break;
      case "deletePart":
        parts.delete(edit.path.replace(/^\/+/, ""));
        break;
      case "replaceBody": {
        // Preserve the original document's section properties, background,
        // and document-level wrappers — only swap the children inside w:body.
        // Reject new external references for the same reason as
        // replaceHeader/Footer: this editor doesn't rewrite document.xml.rels.
        assertNoExternalReferences(edit.body, "body");
        const original = parts.get(PartPath.Document);
        const reserved = reservedBookmarkIds(parts, PartPath.Document);
        parts.set(
          PartPath.Document,
          original
            ? replaceBodyChildrenInDocumentXml(original, edit.body, reserved)
            : renderBodyOnly(edit.body, reserved)
        );
        break;
      }
      case "replaceHeader":
      case "replaceFooter": {
        const path = edit.path.replace(/^\/+/, "");
        const kind = edit.type === "replaceHeader" ? "header" : "footer";
        parts.set(path, renderHeaderFooter(kind, edit.children, reservedBookmarkIds(parts, path)));
        break;
      }
    }
  }

  if (replacedAnchors.size > 0) {
    dropOrphanedComments(parts, replacedAnchors);
  }

  typeNewParts(parts, originalPaths, policy.maxXmlDepth);

  // Write back to a new ZIP
  const archive = createZip({
    level: options?.compressionLevel ?? 6,
    modTime: options?.modTime
  });
  for (const [path, data] of parts) {
    archive.add(path, data);
  }
  return archive.toBuffer();
}

/**
 * List all part paths in a DOCX file without parsing the model.
 *
 * @param buffer - The DOCX file (ZIP or Flat OPC).
 * @param options - How to read it, as `Document.read` takes them (`password`, `securityPolicy`, `checkCrc32`).
 * @returns Array of part paths.
 */
export async function listDocxParts(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  options?: ReadDocxOptions
): Promise<string[]> {
  const where = "Package.listParts";
  const bytes = checkPackageBytes(buffer, where);
  const policy = checkReadOptions(options, where);
  const paths: string[] = [];
  await readPackageEntries(bytes, options, policy, path => {
    paths.push(path);
    return false;
  });
  return paths;
}

/**
 * Read a single part's raw bytes from a DOCX file.
 *
 * @param buffer - The DOCX file (ZIP or Flat OPC).
 * @param path - The part path (e.g. `"word/document.xml"`).
 * @param options - How to read it, as `Document.read` takes them (`password`, `securityPolicy`, `checkCrc32`).
 * @returns The part's raw bytes, or undefined if the part doesn't exist.
 */
export async function readDocxPart(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  path: string,
  options?: ReadDocxOptions
): Promise<Uint8Array<ArrayBuffer> | undefined> {
  const where = "Package.readPart";
  const bytes = checkPackageBytes(buffer, where);
  checkString(path, "path", where);
  const policy = checkReadOptions(options, where);
  // A part name is written with a leading "/" in `[Content_Types].xml` and
  // without one as a ZIP entry; either names the same part.
  const wanted = path.replace(/^\/+/, "");
  let found: Uint8Array | undefined;
  await readPackageEntries(bytes, options, policy, async (entryPath, entry) => {
    if (entryPath !== wanted) {
      return false;
    }
    found = await entry.bytes();
    return true;
  });
  return found && asPlainBytes(found);
}

function checkPackageBytes(buffer: unknown, where: string): Uint8Array {
  return checkBytes(buffer, "buffer", where);
}

/**
 * Visit every part of the package — folder entries skipped, a leading "/"
 * some ZIP writers emit stripped — until `visit` returns `true`, reading it as
 * `Document.read` does: under the options' security policy, an encrypted file
 * decrypted with `password` (a `DocxEncryptedError` without one), a Flat OPC
 * file as the package it holds. Bytes that are not a package are a
 * `DocxParseError`, as `Document.read` reports them.
 *
 * A part's size counts against the policy when its bytes are read, so
 * `listParts` decompresses nothing and `readPart` only the part it returns.
 */
async function readPackageEntries(
  input: Uint8Array,
  options: ReadDocxOptions | undefined,
  policy: Required<DocxSecurityPolicy>,
  visit: (path: string, entry: { bytes(): Promise<Uint8Array> }) => boolean | Promise<boolean>
): Promise<void> {
  checkPackageSize(input, policy);
  const buffer = await decryptIfEncrypted(input, options?.password, policy.maxPackageSize);
  try {
    if (isFlatOpc(buffer)) {
      for (const [path, data] of flatOpcEntries(buffer, policy)) {
        if (await visit(path.replace(/^\/+/, ""), { bytes: async () => data })) {
          return;
        }
      }
      return;
    }
    const budget = createPartBudget(policy);
    for await (const entry of unzipZip(buffer, {
      checkCrc32: options?.checkCrc32,
      maxEntrySize: policy.maxPartSize
    }).entries()) {
      // Folder entries are not parts (OPC §9.1.1.1) and are not written back.
      if (isDirectoryEntry(entry)) {
        continue;
      }
      budget.count();
      const path = entry.path.replace(/^\/+/, "");
      const charged = {
        bytes: (): Promise<Uint8Array> => budget.read(path, entry)
      };
      if (await visit(path, charged)) {
        return;
      }
    }
  } catch (e) {
    if (e instanceof DocxError) {
      throw e;
    }
    const msg = e instanceof Error ? e.message : String(e);
    throw new DocxParseError(`Failed to read DOCX: ${msg}`, { cause: e });
  }
}

const EDIT_TYPES = new Set<string>([
  "replacePart",
  "replacePartText",
  "deletePart",
  "replaceBody",
  "replaceHeader",
  "replaceFooter"
] satisfies IncrementalEdit["type"][]);

/** Every edit's shape, before any is applied. */
function checkIncrementalEdits(edits: unknown, where: string): void {
  checkArray(edits, "edits", where);
  (edits as readonly unknown[]).forEach((value, i) => {
    const at = `edits[${i}]`;
    checkRecord(value, at, where);
    const edit = value as Record<string, unknown>;
    checkOneOf(edit.type, EDIT_TYPES, `${at}.type`, where);
    if (edit.type !== "replaceBody") {
      checkString(edit.path, `${at}.path`, where);
    }
    switch (edit.type) {
      case "replacePart":
        checkBytes(edit.data, `${at}.data`, where);
        break;
      case "replacePartText":
        checkString(edit.text, `${at}.text`, where);
        break;
      case "replaceBody":
        checkNodeArray(edit.body, `${at}.body`, where);
        break;
      case "replaceHeader":
      case "replaceFooter":
        checkNodeArray(edit.children, `${at}.children`, where);
        break;
    }
  });
}

// =============================================================================
// Internal helpers
// =============================================================================

/** Whether a ZIP entry is a folder rather than a part. */
function isDirectoryEntry(entry: { readonly type: string; readonly path: string }): boolean {
  return entry.type === "directory" || /[/\\]$/.test(entry.path);
}

/**
 * Give every part this edit added a content type, as OPC requires of every
 * part (ECMA-376 Part 2 §10.1.2.4). One the final `[Content_Types].xml`
 * already types — by `Override`, or by its extension's `Default` — is left
 * alone. A media part (`…/media/…`) without one is refused: a picture typed
 * generically would not render. Any other gets the generic
 * `application/octet-stream`, by a `Default` for its extension or, without
 * one, an `Override`. Parts the original already had are not judged: they
 * are the input's, not this edit's.
 */
function typeNewParts(
  parts: Map<string, Uint8Array>,
  originalPaths: ReadonlySet<string>,
  maxDepth: number
): void {
  const added = [...parts.keys()].filter(
    path => !originalPaths.has(path) && path !== CONTENT_TYPES_PART && !path.endsWith(".rels")
  );
  for (const path of added) {
    assertValidPartName(path, "Package.edit");
  }
  if (added.length === 0) {
    return;
  }
  const ctXml = parts.get(CONTENT_TYPES_PART);
  if (!ctXml) {
    return;
  }
  const types = readContentTypes(ctXml, maxDepth);
  const { defaults, overrides } = types;
  const declarations: string[] = [];
  // Written in the content-types namespace the root is in, with the prefix the root has for it.
  const prefix = types.rootPrefix === "" ? "" : `${types.rootPrefix}:`;
  for (const path of added) {
    const leaf = path.slice(path.lastIndexOf("/") + 1);
    const dot = leaf.lastIndexOf(".");
    const ext = dot > 0 ? leaf.slice(dot + 1).toLowerCase() : "";
    if (overrides.has(path.toLowerCase()) || (ext && defaults.has(ext))) {
      continue;
    }
    if (MEDIA_PART.test(path)) {
      throw new DocxWriteError(
        `Package.edit: new media part "${path}" has no content type — add a Default for its extension or an Override for it to [Content_Types].xml in the same edit`
      );
    }
    if (ext) {
      defaults.add(ext);
      declarations.push(
        `<${prefix}Default Extension="${xmlEncodeAttr(ext)}" ContentType="${GENERIC_CONTENT_TYPE}"/>`
      );
    } else {
      overrides.add(path.toLowerCase());
      declarations.push(
        `<${prefix}Override PartName="/${xmlEncodeAttr(path)}" ContentType="${GENERIC_CONTENT_TYPE}"/>`
      );
    }
  }
  if (declarations.length > 0) {
    // Into the root element, wherever its end is: the rest of the part is kept byte for byte.
    const inserted = utf8Encoder.encode(
      types.rootSelfClosing
        ? `>${declarations.join("")}</${types.rootName}>`
        : declarations.join("")
    );
    const at = types.rootSelfClosing ? types.rootEnd - 2 : types.rootCloseStart;
    const until = types.rootSelfClosing ? types.rootEnd : at;
    const out = new Uint8Array(ctXml.length - (until - at) + inserted.length);
    out.set(ctXml.subarray(0, at));
    out.set(inserted, at);
    out.set(ctXml.subarray(until), at + inserted.length);
    parts.set(CONTENT_TYPES_PART, out);
  }
}

const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";

/** What `Package.edit` reads of `[Content_Types].xml` to add to it. */
interface ContentTypes {
  /** The extensions it gives a `Default`, lowercase. */
  readonly defaults: Set<string>;
  /** The part names it gives an `Override`, without the leading `/`, lowercase. */
  readonly overrides: Set<string>;
  /** The root element's name, and the prefix it has for the content-types namespace. */
  readonly rootName: string;
  readonly rootPrefix: string;
  readonly rootSelfClosing: boolean;
  /** Where the root's end tag begins in the part's bytes (not self-closing). */
  readonly rootCloseStart: number;
  /** Where the root element ends: just past its `>`. */
  readonly rootEnd: number;
}

/**
 * Read `[Content_Types].xml` from its bytes, with namespaces, so its declarations are known and the
 * root's end is found where it is — not by searching the text, which a comment or a self-closing root
 * misleads. A part that is not such a document is a `DocxParseError`.
 */
function readContentTypes(bytes: Uint8Array, maxDepth: number): ContentTypes {
  // From the first `<`: the lead Office tolerates (see `trimXmlPartLead`) is skipped, not parsed.
  let lead = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  while (
    lead < bytes.length &&
    (bytes[lead] === 0x20 || bytes[lead] === 0x09 || bytes[lead] === 0x0a || bytes[lead] === 0x0d)
  ) {
    lead++;
  }
  const parser = new SaxParser({
    position: false,
    xmlns: true,
    byteOffset: lead,
    maxDepth
  });
  const defaults = new Set<string>();
  const overrides = new Set<string>();
  let root: { name: string; prefix: string } | undefined;
  let rootSelfClosing = false;
  let rootCloseStart = -1;
  let rootEnd = -1;
  parser.on("opentag", tag => {
    if (parser.depth === 0) {
      if (tag.uri !== CONTENT_TYPES_NAMESPACE || tag.local !== "Types") {
        parser.fail(`expected the root <Types> of the content-types namespace, got <${tag.name}>`);
      }
      root = { name: tag.name, prefix: tag.prefix ?? "" };
      rootSelfClosing = tag.isSelfClosing;
      if (tag.isSelfClosing) {
        rootEnd = parser.endByte!;
      }
      return;
    }
    if (parser.depth !== 1 || tag.uri !== CONTENT_TYPES_NAMESPACE) {
      return;
    }
    // A declaration without its content type declares nothing (ECMA-376 Part 2 §10.1.2.2).
    if (!tag.attributes.ContentType) {
      return;
    }
    if (tag.local === "Default" && tag.attributes.Extension) {
      defaults.add(tag.attributes.Extension.toLowerCase());
    } else if (tag.local === "Override" && tag.attributes.PartName) {
      overrides.add(tag.attributes.PartName.replace(/^\/+/, "").toLowerCase());
    }
  });
  parser.on("closetag", tag => {
    if (parser.depth === 0 && !tag.isSelfClosing) {
      rootCloseStart = parser.startByte!;
      rootEnd = parser.endByte!;
    }
  });
  try {
    parser.write(bytes.subarray(lead)).close();
  } catch (error) {
    const limited = asPartLimitError(error, maxDepth);
    if (limited !== error) {
      throw limited;
    }
    throw new DocxParseError(
      `Package.edit: [Content_Types].xml cannot be read: ${(error as Error).message}`,
      {
        cause: error
      }
    );
  }
  const read = root!;
  return {
    defaults,
    overrides,
    rootName: read.name,
    rootPrefix: read.prefix,
    rootSelfClosing,
    rootCloseStart,
    rootEnd
  };
}

const GENERIC_CONTENT_TYPE = "application/octet-stream";
const CONTENT_TYPES_PART = "[Content_Types].xml";
const MEDIA_PART = /(?:^|\/)media\/[^/]+$/;

/** The parts whose content can carry bookmarks (story parts). */
const STORY_PART = /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/;

/** Bookmark, range-permission, revision and drawing ids new content must not take. */
interface ReservedRangeIds {
  readonly reservedBookmarkIds: Set<number>;
  readonly reservedPermIds: Set<number>;
  readonly reservedRevisionIds: Set<number>;
  /** Id generators whose drawing ids (`wp:docPr`) avoid the ids already used. */
  readonly ids: IdGenerators;
}

/**
 * Bookmark, range-permission, revision and drawing ids new content written into `path`
 * must not take: those of every other story part, which this editor passes
 * through untouched. Each kind of id is unique across the whole document, not
 * per part. (The new content's own markers and drawings, raw or modelled, are
 * numbered as it is written.)
 */
function reservedBookmarkIds(
  parts: ReadonlyMap<string, Uint8Array>,
  path: string
): ReservedRangeIds {
  const reservedBookmarkIds = new Set<number>();
  const reservedPermIds = new Set<number>();
  const reservedRevisionIds = new Set<number>();
  const docPrs = new Set<number>();
  for (const [name, data] of parts) {
    if (name !== path && STORY_PART.test(name)) {
      const xml = utf8Decoder.decode(data);
      collectBookmarkStartIds(xml, reservedBookmarkIds);
      collectRangeStartIds(xml, reservedPermIds, "permStart");
      collectRevisionIds(xml, reservedRevisionIds);
      collectDocPrIds(xml, docPrs);
    }
  }
  return {
    reservedBookmarkIds,
    reservedPermIds,
    reservedRevisionIds,
    ids: createIdGenerators({ reservedDocPrIds: docPrs })
  };
}

/** The parts a comment can be anchored in. */
const ANCHOR_PART = /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;

/** Relationship type suffixes of the comment parts, and their usual names. */
const COMMENT_PARTS = {
  comments: "word/comments.xml",
  commentsExtended: "word/commentsExtended.xml",
  commentsIds: "word/commentsIds.xml",
  commentsExtensible: "word/commentsExtensible.xml"
} as const;

/**
 * Remove the comments this edit orphaned: those a replaced body, header or
 * footer anchored (`replacedAnchors`) — by a range or reference mark — and no
 * story part anchors after it. Word reports a comment whose anchor an edit
 * took away as a damaged document. A comment that had no anchor before the
 * edit is not this edit's to remove, and stays. A removed comment's entries in
 * `commentsExtended.xml` (`w15:commentEx`, keyed by the paraId of its
 * paragraphs), `commentsIds.xml` and, through the durable ids found there,
 * `commentsExtensible.xml` go with it; a reply whose parent is removed keeps
 * its own entry without the parent link. `people.xml` is left as it is: an
 * author without comments is still a valid entry.
 */
function dropOrphanedComments(
  parts: Map<string, Uint8Array>,
  replacedAnchors: ReadonlySet<string>
): void {
  const paths = commentPartPaths(parts);
  const commentsData = parts.get(paths.comments);
  if (!commentsData) {
    return;
  }
  const anchored = new Set<string>();
  for (const [name, data] of parts) {
    if (ANCHOR_PART.test(name)) {
      collectCommentAnchors(utf8Decoder.decode(data), anchored);
    }
  }
  let removed = false;
  const droppedParaIds = new Set<string>();
  const comments = utf8Decoder
    .decode(commentsData)
    .replace(/<w:comment\b[^>]*?(?:\/>|>[\s\S]*?<\/w:comment>)/g, element => {
      const id = /\bw:id="([^"]*)"/.exec(element)?.[1];
      if (id === undefined || !replacedAnchors.has(id) || anchored.has(id)) {
        return element;
      }
      for (const m of element.matchAll(/\bw14:paraId="([^"]*)"/g)) {
        droppedParaIds.add(m[1]);
      }
      removed = true;
      return "";
    });
  if (!removed) {
    return;
  }
  parts.set(paths.comments, utf8Encoder.encode(comments));
  if (droppedParaIds.size === 0) {
    return;
  }
  editPart(parts, paths.commentsExtended, xml =>
    withoutCommentsExtendedEntries(xml, droppedParaIds)
  );
  const droppedDurableIds = new Set<string>();
  editPart(parts, paths.commentsIds, xml =>
    withoutCommentIdEntries(xml, droppedParaIds, droppedDurableIds)
  );
  if (droppedDurableIds.size > 0) {
    editPart(parts, paths.commentsExtensible, xml =>
      withoutCommentExtensibleEntries(xml, droppedDurableIds)
    );
  }
}

/** Add the id of every comment anchor (range start or end, reference mark) in `xml` to `out`. */
function collectCommentAnchors(xml: string, out: Set<string>): void {
  for (const m of xml.matchAll(
    /<w:comment(?:Reference|RangeStart|RangeEnd)\b[^>]*?\bw:id="([^"]*)"/g
  )) {
    out.add(m[1]);
  }
}

/** `commentsExtended.xml` without the entries of `paraIds`, and without reply links to them. */
function withoutCommentsExtendedEntries(xml: string, paraIds: ReadonlySet<string>): string {
  return xml
    .replace(/<w15:commentEx\b[^>]*\/>/g, element => {
      const paraId = /\bw15:paraId="([^"]*)"/.exec(element)?.[1];
      return paraId !== undefined && paraIds.has(paraId) ? "" : element;
    })
    .replace(/\s+w15:paraIdParent="([^"]*)"/g, (attr, parent: string) =>
      paraIds.has(parent) ? "" : attr
    );
}

/**
 * `commentsIds.xml` without the entries of `paraIds`; the durable ids of the
 * entries removed are added to `durableIds`.
 */
function withoutCommentIdEntries(
  xml: string,
  paraIds: ReadonlySet<string>,
  durableIds: Set<string>
): string {
  return xml.replace(/<w16cid:commentId\b[^>]*\/>/g, element => {
    const paraId = /\bw16cid:paraId="([^"]*)"/.exec(element)?.[1];
    if (paraId === undefined || !paraIds.has(paraId)) {
      return element;
    }
    const durableId = /\bw16cid:durableId="([^"]*)"/.exec(element)?.[1];
    if (durableId !== undefined) {
      durableIds.add(durableId);
    }
    return "";
  });
}

/** `commentsExtensible.xml` without the entries of `durableIds`. */
function withoutCommentExtensibleEntries(xml: string, durableIds: ReadonlySet<string>): string {
  return xml.replace(
    /<w16cex:commentExtensible\b[^>]*?(?:\/>|>[\s\S]*?<\/w16cex:commentExtensible>)/g,
    element => {
      const durableId = /\bw16cex:durableId="([^"]*)"/.exec(element)?.[1];
      return durableId !== undefined && durableIds.has(durableId) ? "" : element;
    }
  );
}

/** Where the comment parts are, by the document's relationships, else by their usual names. */
function commentPartPaths(
  parts: ReadonlyMap<string, Uint8Array>
): Record<keyof typeof COMMENT_PARTS, string> {
  const paths: Record<keyof typeof COMMENT_PARTS, string> = {
    ...COMMENT_PARTS
  };
  const rels = parts.get("word/_rels/document.xml.rels");
  if (rels) {
    for (const m of utf8Decoder.decode(rels).matchAll(/<Relationship\b[^>]*>/g)) {
      const type = /\bType="[^"]*\/([^"/]+)"/.exec(m[0])?.[1];
      const target = /\bTarget="([^"]*)"/.exec(m[0])?.[1];
      if (type && target && type in paths && !/^[a-z]+:/i.test(target)) {
        paths[type as keyof typeof COMMENT_PARTS] = target.startsWith("/")
          ? target.slice(1)
          : `word/${target}`;
      }
    }
  }
  return paths;
}

function editPart(
  parts: Map<string, Uint8Array>,
  path: string,
  edit: (xml: string) => string
): void {
  const data = parts.get(path);
  if (data) {
    parts.set(path, utf8Encoder.encode(edit(utf8Decoder.decode(data))));
  }
}

function renderBodyOnly(body: readonly BodyContent[], reserved: ReservedRangeIds): Uint8Array {
  const minimalDoc: DocxDocument = { body } as DocxDocument;
  const ctx = createRenderContext(reserved);

  const xml = new XmlWriter();
  renderDocument(xml, minimalDoc, ctx);
  return utf8Encoder.encode(closeStoryPart(xml.toString(), ctx));
}

/**
 * Replace the children inside `<w:body>...</w:body>` of an existing
 * document.xml, preserving the original section properties (the trailing
 * `w:sectPr` directly under w:body), the document background, and any other
 * document-level structure (namespace declarations, custom attributes, etc.).
 *
 * If the document doesn't contain a parseable `<w:body>` block we fall back
 * to a minimal render so callers still get a valid DOCX (this matches the
 * legacy behavior of `replaceBody`).
 */
function replaceBodyChildrenInDocumentXml(
  original: Uint8Array,
  body: readonly BodyContent[],
  reserved: ReservedRangeIds
): Uint8Array {
  const text = utf8Decoder.decode(original);

  // Locate <w:body...> opening tag and the matching </w:body>. We must
  // ignore any literal "<w:body" or "</w:body>" sitting inside an XML
  // comment, CDATA section, or processing instruction — those are common
  // in author-edited documents and confused the previous purely-regex
  // scan.
  const range = findBodyTagRange(text);
  if (!range) {
    return renderBodyOnly(body, reserved);
  }
  const { bodyStart, bodyCloseIdx } = range;

  const innerBody = text.slice(bodyStart, bodyCloseIdx);

  // Preserve the trailing <w:sectPr ...>...</w:sectPr> if present. Section
  // properties at this level apply to the whole document and would otherwise
  // be lost.
  let trailingSectPr = "";
  const sectPrMatch =
    /<w:sectPr(?:\s[^>]*)?>[\s\S]*?<\/w:sectPr>\s*$|<w:sectPr(?:\s[^>]*)?\/>\s*$/.exec(innerBody);
  if (sectPrMatch) {
    trailingSectPr = sectPrMatch[0].trimEnd();
  }

  // Render new body children (no document/body wrappers, no sectPr — we add
  // back the preserved one ourselves).
  const ctx = createRenderContext(reserved);
  const xml = new XmlWriter();
  for (const block of body) {
    renderBodyContent(xml, block, ctx);
  }
  const newInner = closeStoryPart(xml.toString(), ctx) + trailingSectPr;

  const updated = text.slice(0, bodyStart) + newInner + text.slice(bodyCloseIdx);
  return utf8Encoder.encode(updated);
}

/**
 * Locate the inner-text range of `<w:body>...</w:body>`, ignoring any
 * occurrence of the literal tag inside XML comments (`<!-- ... -->`),
 * CDATA sections (`<![CDATA[ ... ]]>`), or processing instructions
 * (`<?...?>`). Returns the offsets of the first byte after the body open
 * tag and the first byte of the body close tag, or `undefined` if either
 * tag was not found.
 *
 * The scanner is a tiny linear pass — no lexer, no DOM. It only knows
 * about the three "skip these contents" sections that can legitimately
 * contain `<` and `>` characters.
 */
function findBodyTagRange(
  text: string
): { readonly bodyStart: number; readonly bodyCloseIdx: number } | undefined {
  let bodyStart = -1;
  let bodyCloseIdx = -1;
  let i = 0;

  while (i < text.length) {
    const lt = text.indexOf("<", i);
    if (lt < 0) {
      break;
    }
    // Comment <!-- ... -->
    if (text.startsWith("<!--", lt)) {
      const end = text.indexOf("-->", lt + 4);
      i = end < 0 ? text.length : end + 3;
      continue;
    }
    // CDATA <![CDATA[ ... ]]>
    if (text.startsWith("<![CDATA[", lt)) {
      const end = text.indexOf("]]>", lt + 9);
      i = end < 0 ? text.length : end + 3;
      continue;
    }
    // Processing instruction <? ... ?>
    if (text.startsWith("<?", lt)) {
      const end = text.indexOf("?>", lt + 2);
      i = end < 0 ? text.length : end + 2;
      continue;
    }
    // Closing tag </w:body>
    if (text.startsWith("</w:body>", lt)) {
      // Track the LAST close tag (matches the previous behaviour of
      // picking up the outermost body) — there should only be one in a
      // well-formed document anyway, but a stale match would just be
      // overwritten by the real one.
      bodyCloseIdx = lt;
      i = lt + "</w:body>".length;
      continue;
    }
    // Opening tag <w:body...>
    if (text.startsWith("<w:body", lt)) {
      const next = text.charCodeAt(lt + 7);
      // Must be either whitespace before attributes or '>' for the bare tag,
      // or '/' for the (rare) self-closing form. Anything else (e.g.
      // "<w:bodyExtra") is a different element name.
      if (
        next === 0x20 ||
        next === 0x09 ||
        next === 0x0a ||
        next === 0x0d ||
        next === 0x3e ||
        next === 0x2f
      ) {
        const gt = text.indexOf(">", lt);
        if (gt < 0) {
          break;
        }
        // Self-closing `<w:body/>` carries no children — fall back to a
        // full render rather than trying to splice into nothing.
        if (text.charCodeAt(gt - 1) === 0x2f) {
          return undefined;
        }
        if (bodyStart < 0) {
          bodyStart = gt + 1;
        }
        i = gt + 1;
        continue;
      }
    }
    // Any other tag — skip past its '>'.
    const gt = text.indexOf(">", lt + 1);
    i = gt < 0 ? text.length : gt + 1;
  }

  if (bodyStart < 0 || bodyCloseIdx < 0 || bodyStart >= bodyCloseIdx) {
    return undefined;
  }
  return { bodyStart, bodyCloseIdx };
}

function renderHeaderFooter(
  kind: "header" | "footer",
  children: readonly (Paragraph | Table)[],
  reserved: ReservedRangeIds
): Uint8Array {
  // Reject content that requires fresh part-level relationships (images,
  // external hyperlinks, charts, alt-chunks, embedded fonts, etc.). The
  // incremental editor only rewrites individual XML parts; it does not
  // rebuild the per-part `*.rels` file or copy in new media binaries, so
  // any new `r:embed` / `r:id` reference would dangle. Callers wanting to
  // add such content should use the full `patchDocument` / `packageDocx`
  // flow instead.
  assertNoExternalReferences(children, kind);

  const ctx = createRenderContext(reserved);
  const helpers = renderHelpersFor(ctx);
  const xml = new XmlWriter();
  const previous = enterRevisionIds(ctx.revisionIds);
  try {
    if (kind === "header") {
      renderHeader(xml, { children }, helpers);
    } else {
      renderFooter(xml, { children }, helpers);
    }
  } finally {
    exitRevisionIds(previous);
  }
  return utf8Encoder.encode(closeStoryPart(xml.toString(), helpers));
}

/**
 * Walk a header/footer / body fragment and throw if it contains any element
 * that needs a per-part relationship registration the incremental editor
 * cannot perform.
 *
 * The list mirrors what {@link renderBodyContent} writes as `r:embed` /
 * `r:id` / `r:link` attributes:
 *   - inline & floating images (`r:embed` / `r:link`)
 *   - external hyperlinks (`r:id` on `w:hyperlink`)
 *   - charts and chart extensions (`r:id` on `w:chart` / `cx:chart`)
 *   - alt-chunks (`r:id` on `w:altChunk`)
 *   - opaque drawings carrying preserved relationships
 */
function assertNoExternalReferences(blocks: readonly BodyContent[], context: string): void {
  let offending: string | undefined;
  walkBlocks(blocks, {
    enterRun(run) {
      for (const c of run.content) {
        if (c.type === "image") {
          offending ??= "inline image";
          return "stop";
        }
      }
      return "continue";
    },
    enterHyperlink(h) {
      if (h.url || h.rId) {
        offending ??= "hyperlink";
        return "stop";
      }
      return "continue";
    },
    visitFloatingImage() {
      offending ??= "floating image";
    },
    visitChart() {
      offending ??= "chart";
    },
    visitChartEx() {
      offending ??= "chart (extended)";
    },
    visitAltChunk() {
      offending ??= "altChunk";
    },
    visitOpaqueDrawing(drawing) {
      if (drawing.referencedRIds.length > 0) {
        offending ??= "opaque drawing with relationships";
      }
    }
  });

  if (offending) {
    throw new DocxUnsupportedFeatureError(
      `incremental ${context} replacement does not support ${offending}: ` +
        "relationship parts (.rels) are not rewritten. Use patchDocument() or " +
        "rebuild the document with packageDocx() instead."
    );
  }
}
