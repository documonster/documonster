/**
 * The limits a {@link DocxSecurityPolicy} sets on a package, applied the same
 * way by every reader of one — `Document.read`, the `Package` functions — so a
 * limit cannot hold in one and not in the other.
 */

import { parseFlatOpc } from "@word/convert/flat-opc";
import { checkSchema, model, w } from "@word/core/option-schema";
import { DocxLimitExceededError } from "@word/errors";
import { checkSecurityPolicy, withPolicyDefaults } from "@word/security/policy";
import type { DocxSecurityPolicy } from "@word/security/policy";

/**
 * The keys of `ReadDocxOptions`, which every reader of a package takes. `securityPolicy` is
 * checked by {@link checkSecurityPolicy}, after them: `null` has always read as the default policy.
 */
export const readOptionsShape = {
  password: w.string().optional(),
  securityPolicy: model<DocxSecurityPolicy>(),
  checkCrc32: w.boolean().optional()
};

const readOptionsSchema = w.record(readOptionsShape).optional();

/**
 * The read options every reader of a package takes (`Document.read`, the `Package`
 * functions, the streaming reader), checked in the name of `where`. Returns the security policy
 * the call reads with: each of its fields read once, checked, and given its default when unset —
 * a reader uses this, never `options.securityPolicy` again.
 */
export function checkReadOptions(options: unknown, where: string): Required<DocxSecurityPolicy> {
  checkSchema(readOptionsSchema, options, where);
  const policy = (options as { securityPolicy?: unknown } | undefined)?.securityPolicy;
  return withPolicyDefaults(
    policy != null ? checkSecurityPolicy(policy, "options.securityPolicy", where) : undefined
  );
}

/** A package larger than the policy's `maxPackageSize` is refused before anything reads it. */
export function checkPackageSize(buffer: Uint8Array, policy: Required<DocxSecurityPolicy>): void {
  if (buffer.length > policy.maxPackageSize) {
    throw new DocxLimitExceededError(
      "packageSize",
      policy.maxPackageSize,
      buffer.length,
      "compressed input larger than maxPackageSize"
    );
  }
}

/**
 * Counts a package's parts against the policy's `maxPartCount`, `maxPartSize`
 * and `maxPackageSize` (cumulative, uncompressed) as they are read — one
 * budget for every reader of a ZIP's entries.
 */
export function createPartBudget(policy: Required<DocxSecurityPolicy>): {
  /** Count one more part, before its bytes are read. */
  count(): void;
  /**
   * Charge `size` uncompressed bytes of the part `path` without reading them: a part not read
   * (by its declared size), or the next piece of one streamed.
   */
  charge(path: string, size: number): void;
  /**
   * Read a part's bytes, held to `maxPartSize` and to what is left of `maxPackageSize` as it is
   * decompressed: a part declaring more is refused before anything is inflated, and one inflating
   * past its declaration is stopped at the chunk that crosses the limit.
   */
  read(path: string, entry: BudgetedEntry): Promise<Uint8Array>;
} {
  let parts = 0;
  let total = 0;
  const check = (path: string, size: number): void => {
    if (size > policy.maxPartSize) {
      throw new DocxLimitExceededError(
        "partSize",
        policy.maxPartSize,
        size,
        `entry "${path}" exceeds maxPartSize`
      );
    }
    if (total + size > policy.maxPackageSize) {
      throw new DocxLimitExceededError(
        "packageSize",
        policy.maxPackageSize,
        total + size,
        "cumulative uncompressed entry size exceeds maxPackageSize"
      );
    }
  };
  return {
    count() {
      parts++;
      if (parts > policy.maxPartCount) {
        throw new DocxLimitExceededError(
          "partCount",
          policy.maxPartCount,
          parts,
          "ZIP contains more entries than maxPartCount"
        );
      }
    },
    charge(path, size) {
      check(path, size);
      total += size;
    },
    async read(path, entry) {
      check(path, entry.info.uncompressedSize);
      const chunks: Uint8Array[] = [];
      let size = 0;
      for await (const chunk of entry.stream()) {
        size += chunk.length;
        check(path, size);
        chunks.push(chunk);
      }
      total += size;
      if (chunks.length === 1) {
        return chunks[0]!;
      }
      const out = new Uint8Array(size);
      let at = 0;
      for (const chunk of chunks) {
        out.set(chunk, at);
        at += chunk.length;
      }
      return out;
    }
  };
}

/** What {@link createPartBudget} reads a part from: a ZIP entry. */
export interface BudgetedEntry {
  /** The entry's record; `uncompressedSize` is the size it declares. */
  readonly info: { readonly uncompressedSize: number };
  stream(): AsyncIterable<Uint8Array>;
}

/** The parts of a Flat OPC document, held to the limits a ZIP's entries are. */
export function flatOpcEntries(
  buffer: Uint8Array,
  policy: Required<DocxSecurityPolicy>
): Map<string, Uint8Array> {
  const entries = parseFlatOpc(buffer);
  if (entries.size > policy.maxPartCount) {
    throw new DocxLimitExceededError(
      "partCount",
      policy.maxPartCount,
      entries.size,
      "Flat OPC document contains more parts than maxPartCount"
    );
  }
  let total = 0;
  for (const [path, data] of entries) {
    if (data.length > policy.maxPartSize) {
      throw new DocxLimitExceededError(
        "partSize",
        policy.maxPartSize,
        data.length,
        `part "${path}" exceeds maxPartSize`
      );
    }
    total += data.length;
  }
  if (total > policy.maxPackageSize) {
    throw new DocxLimitExceededError(
      "packageSize",
      policy.maxPackageSize,
      total,
      "cumulative part size exceeds maxPackageSize"
    );
  }
  return entries;
}
