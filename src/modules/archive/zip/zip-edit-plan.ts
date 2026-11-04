import { checkArchiveSource, checkString, checkZipEntryOptions } from "@archive/core/entry-checks";
import { ArchiveArgumentError, ArchiveError } from "@archive/core/errors";
import type { ArchiveSource } from "@archive/io/archive-source";
import { resolveArchiveSourceToBuffer } from "@archive/io/archive-source";
import type { ZipEntryOptions } from "@archive/zip";

/**
 * Serializable operation types for ZipEditPlan
 */
export type ZipEditOp =
  | {
      type: "set";
      name: string;
      source: ArchiveSource;
      options?: ZipEntryOptions;
    }
  | {
      type: "delete";
      name: string;
    }
  | {
      type: "deleteDirectory";
      prefix: string;
    }
  | {
      type: "rename";
      from: string;
      to: string;
    }
  | {
      type: "comment";
      comment?: string;
    };

/**
 * Serialized form of a ZipEditOp.
 *
 * For "set" operations, the source is stored as base64-encoded data.
 * This allows the plan to be safely serialized to JSON and restored later.
 */
export type SerializedZipEditOp =
  | {
      type: "set";
      name: string;
      /** Base64-encoded source data */
      data: string;
      options?: ZipEntryOptions;
    }
  | {
      type: "delete";
      name: string;
    }
  | {
      type: "deleteDirectory";
      prefix: string;
    }
  | {
      type: "rename";
      from: string;
      to: string;
    }
  | {
      type: "comment";
      comment?: string;
    };

/**
 * Serialized form of the entire ZipEditPlan.
 */
export interface SerializedZipEditPlan {
  version: 1;
  ops: SerializedZipEditOp[];
}

/**
 * A reusable, serializable description of ZIP edits.
 *
 * This is intentionally decoupled from any particular ZIP input.
 * Apply it to a {@link ZipEditor} to execute.
 */
export class ZipEditPlan {
  private readonly _ops: ZipEditOp[];

  constructor(ops: readonly ZipEditOp[] = []) {
    this._ops = [...ops];
  }

  /** The operations, in order: a frozen copy, so changing it changes nothing here. */
  getOperations(): readonly ZipEditOp[] {
    return Object.freeze(this._ops.map(op => Object.freeze({ ...op })));
  }

  /**
   * Add or replace an entry. The source and options are checked now; the name is checked
   * when the plan is applied, against the editor's naming rules.
   */
  set(name: string, source: ArchiveSource, options?: ZipEntryOptions): this {
    checkString("Archive.ZipEditPlan#set", "name", name);
    checkArchiveSource("Archive.ZipEditPlan#set", "source", source);
    checkZipEntryOptions("Archive.ZipEditPlan#set", options);
    this._ops.push({ type: "set", name, source, options });
    return this;
  }

  delete(name: string): this {
    this._ops.push({ type: "delete", name });
    return this;
  }

  /**
   * Delete a directory and all its contents recursively.
   *
   * @param prefix - The directory path prefix to delete (with or without trailing slash)
   */
  deleteDirectory(prefix: string): this {
    this._ops.push({ type: "deleteDirectory", prefix });
    return this;
  }

  rename(from: string, to: string): this {
    this._ops.push({ type: "rename", from, to });
    return this;
  }

  setComment(comment?: string): this {
    checkString("Archive.ZipEditPlan#setComment", "comment", comment, true);
    this._ops.push({ type: "comment", comment });
    return this;
  }

  /**
   * Apply this plan to a target editor-like object.
   */
  applyTo(target: {
    set(name: string, source: ArchiveSource, options?: ZipEntryOptions): unknown;
    delete(name: string): unknown;
    deleteDirectory(prefix: string): unknown;
    rename(from: string, to: string): unknown;
    setComment(comment?: string): unknown;
  }): void {
    const t = target as Record<string, unknown> | null;
    if (
      typeof t !== "object" ||
      t === null ||
      ["set", "delete", "deleteDirectory", "rename", "setComment"].some(
        method => typeof t[method] !== "function"
      )
    ) {
      throw new ArchiveArgumentError(
        "target",
        "Archive.ZipEditPlan#applyTo: target must have set, delete, deleteDirectory, rename and setComment methods"
      );
    }
    for (const op of this._ops) {
      switch (op.type) {
        case "set":
          target.set(op.name, op.source, op.options);
          break;
        case "delete":
          target.delete(op.name);
          break;
        case "deleteDirectory":
          target.deleteDirectory(op.prefix);
          break;
        case "rename":
          target.rename(op.from, op.to);
          break;
        case "comment":
          target.setComment(op.comment);
          break;
        default: {
          const _exhaustive: never = op;
          throw new Error(`Unknown ZipEditOp: ${String((_exhaustive as { type: string }).type)}`);
        }
      }
    }
  }

  /**
   * Return a new plan which runs this plan, then `other`.
   */
  concat(other: ZipEditPlan): ZipEditPlan {
    if (!(other instanceof ZipEditPlan)) {
      throw new ArchiveArgumentError(
        "other",
        "Archive.ZipEditPlan#concat: other must be a ZipEditPlan"
      );
    }
    return new ZipEditPlan([...this._ops, ...other._ops]);
  }

  /**
   * Serialize this plan to a JSON-compatible object.
   *
   * All source data will be resolved to buffers and encoded as base64.
   * This is useful for storing the plan in a database, file, or sending over a network.
   *
   * @param signal - Optional abort signal for cancellation
   * @returns Serialized plan that can be safely JSON.stringify'd
   */
  async serialize(signal?: AbortSignal): Promise<SerializedZipEditPlan> {
    const ops: SerializedZipEditOp[] = [];

    for (const op of this._ops) {
      if (op.type === "set") {
        // Resolve source to buffer and encode as base64
        const buffer = await resolveArchiveSourceToBuffer(op.source, { signal });
        const base64 = uint8ArrayToBase64(buffer);
        ops.push({
          type: "set",
          name: op.name,
          data: base64,
          options: op.options
        });
      } else {
        // Other operations don't contain sources, pass through
        ops.push(op as SerializedZipEditOp);
      }
    }

    return { version: 1, ops };
  }

  /**
   * Serialize this plan to a JSON string.
   *
   * @param signal - Optional abort signal for cancellation
   * @returns JSON string representation of the plan
   */
  async toJson(signal?: AbortSignal): Promise<string> {
    const serialized = await this.serialize(signal);
    return JSON.stringify(serialized);
  }

  /**
   * @deprecated Use {@link toJson}. Being named `toJSON`, this is also what
   * `JSON.stringify(plan)` calls — and since it returns a Promise, that call serialises
   * `{}`. Removed in 1.0.
   */
  async toJSON(signal?: AbortSignal): Promise<string> {
    return this.toJson(signal);
  }

  /**
   * Create a ZipEditPlan from a serialized object.
   *
   * @param data - Serialized plan object
   * @returns Restored ZipEditPlan
   */
  static deserialize(data: SerializedZipEditPlan): ZipEditPlan {
    if (typeof data !== "object" || data === null) {
      throw new ArchiveArgumentError(
        "data",
        `Archive.ZipEditPlan.deserialize: data must be a serialized plan object, got ${
          data === null ? "null" : typeof data
        }`
      );
    }
    if (data.version !== 1) {
      throw new ArchiveError(`Unsupported ZipEditPlan version: ${data.version}`);
    }

    if (!Array.isArray(data.ops)) {
      throw new ArchiveArgumentError(
        "data.ops",
        "Archive.ZipEditPlan.deserialize: data.ops must be an array"
      );
    }
    const ops: ZipEditOp[] = data.ops.map(op => {
      if (op.type === "set") {
        // Decode base64 back to Uint8Array
        const buffer = base64ToUint8Array(op.data);
        return {
          type: "set" as const,
          name: op.name,
          source: buffer,
          options: op.options
        };
      }
      return op as ZipEditOp;
    });

    return new ZipEditPlan(ops);
  }

  /**
   * Create a ZipEditPlan from a JSON string.
   *
   * @param json - JSON string representation of a serialized plan
   * @returns Restored ZipEditPlan
   */
  static fromJson(json: string): ZipEditPlan {
    checkString("Archive.ZipEditPlan.fromJson", "json", json);
    let data: SerializedZipEditPlan;
    try {
      data = JSON.parse(json) as SerializedZipEditPlan;
    } catch (cause) {
      throw new ArchiveArgumentError(
        "json",
        "Archive.ZipEditPlan.fromJson: json is not valid JSON",
        {
          cause
        }
      );
    }
    return ZipEditPlan.deserialize(data);
  }

  /**
   * @deprecated Use {@link ZipEditPlan.fromJson}. Removed in 1.0.
   */
  static fromJSON(json: string): ZipEditPlan {
    return ZipEditPlan.fromJson(json);
  }
}

/**
 * Convert Uint8Array to base64 string
 */
function uint8ArrayToBase64(bytes: Uint8Array): string {
  // Use btoa in browser, Buffer in Node.js
  if (typeof Buffer !== "undefined") {
    if (Buffer.isBuffer(bytes)) {
      return bytes.toString("base64");
    }
    return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString("base64");
  }
  // Browser fallback using btoa
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

/**
 * Convert base64 string to Uint8Array
 */
function base64ToUint8Array(base64: string): Uint8Array {
  // Use Buffer in Node.js, atob in browser
  if (typeof Buffer !== "undefined") {
    return Buffer.from(base64, "base64");
  }
  // Browser fallback using atob
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}
