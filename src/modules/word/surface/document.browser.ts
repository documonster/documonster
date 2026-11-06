/**
 * `Document` namespace surface for the browser build: everything `document.ts`
 * publishes except the Node-only file-path members `readFile` / `writeFile`.
 * Read the bytes yourself and use `Document.read`; write the bytes of
 * `Document.toBuffer`.
 */
export * from "@word/surface/document-portable";
