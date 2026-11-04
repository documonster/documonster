/**
 * `Fs` namespace surface (Node only) — the file-system helpers the archive layer is
 * built on: traversal, globbing, existence checks, byte/text I/O, streams, temp dirs.
 *
 * `import { Fs } from "documonster/archive"` → `Fs.glob("**\/*.json")`,
 * `Fs.ensureDir(dir)`, `Fs.readFileBytes(path)`, …
 *
 * A namespace rather than flat exports because these names (`glob`, `remove`,
 * `copyFile`, `createReadStream`) are generic enough to collide with `node:fs` and
 * with a consumer's own helpers. It lives on the archive entry because it is the layer
 * `ArchiveFile` (`addGlob`, `extractTo`, `writeFile`) is built on: the glob semantics
 * and the overwrite/permission behaviour an archive caller sees are these helpers', and
 * publishing them anywhere else would be a second entry for one implementation.
 */
export {
  traverseDirectory,
  traverseDirectorySync,
  glob,
  globSync,
  globToRegex,
  matchGlob,
  matchGlobAny,
  ensureDir,
  ensureDirSync,
  fileExists,
  fileExistsSync,
  readFileBytes,
  readFileBytesSync,
  writeFileBytes,
  writeFileBytesSync,
  setFileTime,
  setFileTimeSync,
  safeStats,
  safeStatsSync,
  readFileText,
  readFileTextSync,
  writeFileText,
  writeFileTextSync,
  remove,
  removeSync,
  copyFile,
  copyFileSync,
  createReadStream,
  createWriteStream,
  createTempDir,
  createTempDirSync
} from "@utils/fs";
