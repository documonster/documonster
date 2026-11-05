/**
 * `Io` namespace surface — document serialization / deserialization.
 *
 * `import { Io } from "documonster/word"` →
 *   `Io.read(buffer)`, `Io.toBuffer(doc)`, `Io.toBase64(doc)`,
 *   `Io.toFlatOpc(doc)`, and on Node `Io.readFile(path)`, `Io.writeFile(doc, path)`
 *   — tree-shaken via `export * as Io`. The browser build resolves this module to
 *   `io.browser.ts`, which publishes no file-path member.
 */
export * from "@word/surface/io-portable";
export { readFile, writeFile } from "@word/document-io";
