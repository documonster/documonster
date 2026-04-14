import {
  encodeZipStringWithCodec,
  resolveZipStringCodec,
  type ZipStringCodec,
  type ZipStringEncoding
} from "@archive/shared/text";
import {
  buildZipTimestampExtraField,
  dateToZipDos,
  type ZipExtraTimestamps,
  type ZipTimestampMode
} from "@archive/zip-spec/timestamps";
import {
  buildUnicodeCommentExtraField,
  buildUnicodePathExtraField
} from "@archive/zip-spec/zip-extra-fields";
import {
  COMPRESSION_DEFLATE,
  COMPRESSION_STORE,
  FLAG_DATA_DESCRIPTOR,
  FLAG_UTF8,
  concatExtraFields
} from "@archive/zip-spec/zip-records";

export interface ZipEntryMetadata {
  nameBytes: Uint8Array;
  commentBytes: Uint8Array;
  dosTime: number;
  dosDate: number;
  extraField: Uint8Array;
  compressionMethod: number;
  flags: number;
}

export interface ZipEntryMetadataInput {
  name: string;
  comment?: string;
  modTime: Date;
  atime?: Date;
  ctime?: Date;
  birthTime?: Date;
  timestamps: ZipTimestampMode;
  /** If true, set FLAG_DATA_DESCRIPTOR and expect CRC/sizes written later. */
  useDataDescriptor: boolean;
  /** If true, use DEFLATE; else STORE. */
  deflate: boolean;

  /**
   * String codec for name/comment.
   * Can be a pre-resolved ZipStringCodec or a ZipStringEncoding shorthand.
   */
  codec?: ZipStringCodec | ZipStringEncoding;
}

export function resolveZipCompressionMethod(deflate: boolean): number {
  return deflate ? COMPRESSION_DEFLATE : COMPRESSION_STORE;
}

export function resolveZipFlags(useDataDescriptor: boolean, useUtf8Flag = true): number {
  let flags = 0;
  if (useUtf8Flag) {
    flags |= FLAG_UTF8;
  }
  if (useDataDescriptor) {
    flags |= FLAG_DATA_DESCRIPTOR;
  }
  return flags;
}

export function buildZipEntryMetadata(input: ZipEntryMetadataInput): ZipEntryMetadata {
  const codec = resolveZipStringCodec(input.codec);
  const nameBytes = codec.encode(input.name);
  const commentBytes = encodeZipStringWithCodec(input.comment, codec);
  const { dosTime, dosDate } = dateToZipDos(input.modTime);
  const extra: ZipExtraTimestamps | undefined =
    input.atime || input.ctime || input.birthTime
      ? { atime: input.atime, ctime: input.ctime, birthTime: input.birthTime }
      : undefined;
  let extraField = buildZipTimestampExtraField(input.modTime, input.timestamps, extra);

  if (!codec.useUtf8Flag && codec.useUnicodeExtraFields) {
    if (input.name) {
      extraField = concatExtraFields(extraField, buildUnicodePathExtraField(nameBytes, input.name));
    }
    if (input.comment) {
      extraField = concatExtraFields(
        extraField,
        buildUnicodeCommentExtraField(commentBytes, input.comment)
      );
    }
  }

  return {
    nameBytes,
    commentBytes,
    dosTime,
    dosDate,
    extraField,
    compressionMethod: resolveZipCompressionMethod(input.deflate),
    flags: resolveZipFlags(input.useDataDescriptor, codec.useUtf8Flag)
  };
}
