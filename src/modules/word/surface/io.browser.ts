/**
 * `Io` namespace surface for the browser build — deprecated, as `io.ts` is:
 * everything it publishes except the Node-only file-path members `readFile` /
 * `writeFile`. Use `Document.read` and `Document.toBuffer`.
 */
export * from "@word/surface/io-portable";
