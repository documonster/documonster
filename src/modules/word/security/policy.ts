/**
 * DOCX Module - Security Policy
 *
 * Defines configurable security policies that control how potentially
 * dangerous content (VBA, OLE, external targets, etc.) is handled
 * during document reading and writing.
 *
 * **Implementation status of each field (read this before relying on
 * the strict preset for security guarantees):**
 *
 * - `preserveVbaProject` — IMPLEMENTED on the read path. When `false`
 *   the macro binary is dropped from the returned model.
 * - `preserveAltChunks` — IMPLEMENTED on the read path. When `false`
 *   altChunk body items are removed before the document is returned.
 * - `allowExternalTargets` — IMPLEMENTED on the read path. When `false`
 *   `Hyperlink.url` is left undefined; only anchor-based hyperlinks
 *   survive. The link wrapper and its inner runs remain.
 * - `maxPackageSize`, `maxPartSize`, `maxPartCount`, `maxXmlDepth` —
 *   IMPLEMENTED. Enforced during ZIP parsing / XML parsing; oversize
 *   inputs raise `DocxLimitExceededError` instead of being processed.
 * - `dropSignaturesOnModify` — IMPLEMENTED. Both writers skip any opaque
 *   parts under `_xmlsignatures/` when `true` (the default). The
 *   `_rels/.rels` file is rebuilt from scratch by the packager, so any
 *   sigOrigin relationship pointing at the now-absent signatures is also
 *   removed automatically. Set `false` to keep the parts; in that case
 *   downstream verifiers will treat the signatures as broken because the
 *   document body has been re-emitted.
 * - `preserveOleObjects` — IMPLEMENTED. On the read path, OLE embedding
 *   binaries (word/embeddings/*.bin and similar) are excluded from
 *   `opaqueParts` when `false`. On the write path (both buffered and
 *   streaming) the same paths are skipped when emitting opaqueParts. The
 *   relationships that referenced the binary remain in the owning part's
 *   .rels — strip those manually if you need a fully clean document.
 * - `rawXmlPolicy` — IMPLEMENTED on the write path. `"preserve"` (default)
 *   writes opaque rawXml verbatim; `"strip"` skips the offending fragments
 *   while keeping the surrounding element well-formed; `"reject"` throws
 *   `DocxRawXmlPolicyError` so the caller can sanitise the model first.
 *   Sites covered: opaqueRun, opaqueParagraphChild, opaqueDrawing,
 *   drawingShape rawXml/_advancedFillXml/_advancedEffectsXml,
 *   settings.mailMergeRawXml, theme.formatScheme.rawXml, theme.extLstXml,
 *   webSettings.rawXml, numPicBullet.rawVmlXml.
 *
 * Pull requests welcome for the not-yet-enforced fields. Until they
 * arrive, callers needing those guarantees should sanitise the input
 * model themselves before serialisation.
 */

// =============================================================================
// Security Policy Interface
// =============================================================================

export interface WordSecurityPolicy {
  /** Whether to preserve VBA project binary on read/write. Default: true (preserve but don't execute). */
  readonly preserveVbaProject?: boolean;
  /** Whether to preserve OLE embedded objects. Default: true. */
  readonly preserveOleObjects?: boolean;
  /** Whether to preserve altChunk parts (embedded HTML/RTF). Default: true. */
  readonly preserveAltChunks?: boolean;
  /** Whether to drop digital signatures when document is modified. Default: true. */
  readonly dropSignaturesOnModify?: boolean;
  /** Maximum total package size in bytes. Default: 500MB. */
  readonly maxPackageSize?: number;
  /** Maximum single part size in bytes. Default: 100MB. */
  readonly maxPartSize?: number;
  /** Maximum ZIP entry count. Default: 10000. */
  readonly maxPartCount?: number;
  /** Maximum nesting depth for XML parsing. Default: 256. */
  readonly maxXmlDepth?: number;
  /** Raw XML policy for opaque/preserved content. */
  readonly rawXmlPolicy?: "preserve" | "strip" | "reject";
  /** Whether to allow external relationship targets (URLs). Default: true. */
  readonly allowExternalTargets?: boolean;
}

// =============================================================================
// Default Policies
// =============================================================================

/** Default security policy: preserve everything, generous limits. */
export const DEFAULT_SECURITY_POLICY: Required<WordSecurityPolicy> = {
  preserveVbaProject: true,
  preserveOleObjects: true,
  preserveAltChunks: true,
  dropSignaturesOnModify: true,
  maxPackageSize: 500 * 1024 * 1024,
  maxPartSize: 100 * 1024 * 1024,
  maxPartCount: 10000,
  maxXmlDepth: 256,
  rawXmlPolicy: "preserve",
  allowExternalTargets: true
};

/** Strict security policy: strip dangerous content, tight limits. */
export const STRICT_SECURITY_POLICY: Required<WordSecurityPolicy> = {
  preserveVbaProject: false,
  preserveOleObjects: false,
  preserveAltChunks: false,
  dropSignaturesOnModify: true,
  maxPackageSize: 100 * 1024 * 1024,
  maxPartSize: 50 * 1024 * 1024,
  maxPartCount: 1000,
  maxXmlDepth: 128,
  rawXmlPolicy: "strip",
  allowExternalTargets: false
};

// =============================================================================
// Policy Resolution
// =============================================================================

/**
 * Merge a partial security policy with the defaults.
 * Any unset fields fall back to DEFAULT_SECURITY_POLICY values.
 *
 * @param policy - Partial policy overrides (or undefined for full defaults).
 * @returns A fully resolved security policy.
 */
export function resolveSecurityPolicy(
  policy?: Partial<WordSecurityPolicy>
): Required<WordSecurityPolicy> {
  if (!policy) {
    return DEFAULT_SECURITY_POLICY;
  }
  return {
    preserveVbaProject: policy.preserveVbaProject ?? DEFAULT_SECURITY_POLICY.preserveVbaProject,
    preserveOleObjects: policy.preserveOleObjects ?? DEFAULT_SECURITY_POLICY.preserveOleObjects,
    preserveAltChunks: policy.preserveAltChunks ?? DEFAULT_SECURITY_POLICY.preserveAltChunks,
    dropSignaturesOnModify:
      policy.dropSignaturesOnModify ?? DEFAULT_SECURITY_POLICY.dropSignaturesOnModify,
    maxPackageSize: policy.maxPackageSize ?? DEFAULT_SECURITY_POLICY.maxPackageSize,
    maxPartSize: policy.maxPartSize ?? DEFAULT_SECURITY_POLICY.maxPartSize,
    maxPartCount: policy.maxPartCount ?? DEFAULT_SECURITY_POLICY.maxPartCount,
    maxXmlDepth: policy.maxXmlDepth ?? DEFAULT_SECURITY_POLICY.maxXmlDepth,
    rawXmlPolicy: policy.rawXmlPolicy ?? DEFAULT_SECURITY_POLICY.rawXmlPolicy,
    allowExternalTargets:
      policy.allowExternalTargets ?? DEFAULT_SECURITY_POLICY.allowExternalTargets
  };
}
