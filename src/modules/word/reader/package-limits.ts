/**
 * The limits a {@link DocxSecurityPolicy} sets on a package, applied the same
 * way by every reader of one — `Document.read`, the `Package` functions — so a
 * limit cannot hold in one and not in the other.
 */

import { parseFlatOpc } from "@word/convert/flat-opc";
import { checkSchema, model, w } from "@word/core/option-schema";
import { DocxLimitExceededError } from "@word/errors";
import { checkSecurityPolicy } from "@word/security/policy";
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
 * functions, the streaming reader), checked in the name of `where`.
 */
export function checkReadOptions(options: unknown, where: string): void {
  checkSchema(readOptionsSchema, options, where);
  const policy = options?.securityPolicy;
  if (policy != null) {
    checkSecurityPolicy(policy, "options.securityPolicy", where);
  }
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
  /** Charge a part's uncompressed bytes. */
  charge(path: string, size: number): void;
} {
  let parts = 0;
  let total = 0;
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
      if (size > policy.maxPartSize) {
        throw new DocxLimitExceededError(
          "partSize",
          policy.maxPartSize,
          size,
          `entry "${path}" exceeds maxPartSize`
        );
      }
      total += size;
      if (total > policy.maxPackageSize) {
        throw new DocxLimitExceededError(
          "packageSize",
          policy.maxPackageSize,
          total,
          "cumulative uncompressed entry size exceeds maxPackageSize"
        );
      }
    }
  };
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
