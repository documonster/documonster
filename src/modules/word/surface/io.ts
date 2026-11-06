/**
 * `Io` namespace surface — deprecated: reading and writing a document are
 * `Document.read`, `Document.toBuffer`, `Document.toBase64`,
 * `Document.toFlatOpc`, and on Node `Document.readFile` / `Document.writeFile`,
 * beside `Document.create`.
 *
 * The browser build resolves this module to `io.browser.ts`, which publishes no
 * file-path member.
 */
import { readFile as readFileImpl, writeFile as writeFileImpl } from "@word/document-io";

export * from "@word/surface/io-portable";

/** @deprecated Use `Document.readFile`. Removed in 1.0. */
export const readFile = readFileImpl;
/** @deprecated Use `Document.writeFile`. Removed in 1.0. */
export const writeFile = writeFileImpl;
