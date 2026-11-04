/**
 * `Archive` namespace surface — Node entry: every cross-platform member plus the
 * on-disk `ArchiveFile` (built with `new ArchiveFile(options)`, opened with
 * `ArchiveFile.fromFile` / `fromFileSync` / `fromBuffer`) and the Node-stream adapter
 * `toNodeReadable`.
 */
export * from "@archive/surface/archive";
export { ArchiveFile } from "@archive/fs/archive-file";
export { toNodeReadable } from "@archive/io/archive-source.node";
