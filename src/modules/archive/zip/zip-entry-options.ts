/**
 * Shared helpers for building ZipDeflateFile options.
 */

import type { ZipStringEncoding } from "@archive/shared/text";
import type { ZipTimestampMode } from "@archive/zip-spec/timestamps";
import type { ZipPathOptions } from "@archive/zip-spec/zip-path";
import type { Zip64Mode } from "@archive/zip-spec/zip-records";

export type ZipDeflateFileEntryOptions = {
  level?: number;
  modTime?: Date;
  atime?: Date;
  ctime?: Date;
  birthTime?: Date;
  comment?: string;
  zip64?: Zip64Mode;
  encoding?: ZipStringEncoding;
  mode?: number;
  msDosAttributes?: number;
  externalAttributes?: number;
  versionMadeBy?: number;
};

export type ZipDeflateFileDefaults = {
  level: number;
  modTime: Date;
  timestamps: ZipTimestampMode;
  smartStore: boolean;
  zip64: Zip64Mode;
  path: false | ZipPathOptions;
  encoding?: ZipStringEncoding;
};

export function buildZipDeflateFileOptions(
  entryOptions: ZipDeflateFileEntryOptions | undefined,
  defaults: ZipDeflateFileDefaults
) {
  return {
    level: entryOptions?.level ?? defaults.level,
    modTime: entryOptions?.modTime ?? defaults.modTime,
    atime: entryOptions?.atime,
    ctime: entryOptions?.ctime,
    birthTime: entryOptions?.birthTime,
    timestamps: defaults.timestamps,
    comment: entryOptions?.comment,
    smartStore: defaults.smartStore,
    zip64: entryOptions?.zip64 ?? defaults.zip64,
    path: defaults.path,
    encoding: entryOptions?.encoding ?? defaults.encoding,
    mode: entryOptions?.mode,
    msDosAttributes: entryOptions?.msDosAttributes,
    externalAttributes: entryOptions?.externalAttributes,
    versionMadeBy: entryOptions?.versionMadeBy
  };
}
