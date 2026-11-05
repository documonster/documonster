/**
 * `Io` namespace surface for the browser build: everything `io.ts` publishes
 * except the Node-only file-path members `readFile` / `writeFile`. Read the
 * bytes yourself and use `Io.read`; write the bytes of `Io.toBuffer`.
 */
export * from "@word/surface/io-portable";
