/**
 * `WordCrypto.Cfb` — the Compound File Binary container an encrypted DOCX is
 * stored in: `WordCrypto.Cfb.read(bytes)` lists its streams,
 * `WordCrypto.Cfb.write(entries)` builds one.
 */
import { checkBytes } from "@word/core/argument-checks";
import { readCfb } from "@word/security/cfb-reader";
import type { CfbEntry } from "@word/security/cfb-reader";

export { writeCfb as write } from "@word/security/cfb-reader";

/**
 * List the streams of a CFB (OLE2 Compound File), v3 or v4.
 *
 * @throws {DocxArgumentError} If `buffer` is not bytes.
 * @throws {DocxParseError} If the bytes are not a CFB file.
 */
export function read(buffer: Uint8Array | ArrayBuffer | ArrayBufferView): CfbEntry[] {
  return readCfb(checkBytes(buffer, "buffer", "WordCrypto.Cfb.read"));
}
