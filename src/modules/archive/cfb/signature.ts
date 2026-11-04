/**
 * Recognising a compound file (MS-CFB) and an encrypted OOXML package by their first bytes — small and free of the
 * CFB reader and the encryption code, so a reader can tell an encrypted file apart before loading either.
 */

/** CFB file signature: D0CF11E0A1B11AE1. */
export const CFB_SIGNATURE: readonly number[] = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Whether bytes start with the CFB signature `D0 CF 11 E0 A1 B1 1A E1`. */
export function isCfb(bytes: Uint8Array): boolean {
  return (
    bytes.length >= CFB_SIGNATURE.length && CFB_SIGNATURE.every((byte, i) => bytes[i] === byte)
  );
}

/** `name` as a compound-file directory entry stores it: UTF-16LE. */
function directoryName(name: string): Uint8Array {
  return Uint8Array.from([...name].flatMap(ch => [ch.charCodeAt(0), 0]));
}

const ENCRYPTION_INFO = directoryName("EncryptionInfo");
const ENCRYPTED_PACKAGE = directoryName("EncryptedPackage");

function contains(haystack: Uint8Array, needle: Uint8Array): boolean {
  const first = needle[0]!;
  for (
    let i = haystack.indexOf(first);
    i >= 0 && i + needle.length <= haystack.length;
    i = haystack.indexOf(first, i + 1)
  ) {
    let j = 1;
    while (j < needle.length && haystack[i + j] === needle[j]) {
      j++;
    }
    if (j === needle.length) {
      return true;
    }
  }
  return false;
}

/**
 * Whether a compound file holds the `EncryptionInfo` and `EncryptedPackage` streams of an encrypted OOXML package
 * (MS-OFFCRYPTO §2.3.4) — what Office writes for a file with a password to open — rather than being a legacy binary
 * document (`.xls`, `.doc`). The directory names are found as the directory stores them, UTF-16LE, without parsing
 * the container; the decryption that follows parses it.
 */
export function hasEncryptedPackageStreams(bytes: Uint8Array): boolean {
  return isCfb(bytes) && contains(bytes, ENCRYPTION_INFO) && contains(bytes, ENCRYPTED_PACKAGE);
}
