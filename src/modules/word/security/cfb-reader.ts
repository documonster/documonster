/**
 * DOCX Module - CFB (Compound File Binary) Reader
 *
 * Minimal reader for OLE2/CFB format used by encrypted Office documents.
 * Implements enough of MS-CFB to extract named streams (EncryptionInfo, EncryptedPackage).
 *
 * References:
 *   - MS-CFB: Compound File Binary File Format
 *   - https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb
 */

import { DocxParseError } from "../errors";

/** A stream entry extracted from/written to a CFB file. */
export interface CfbEntry {
  /** Stream name (UTF-16LE decoded). */
  readonly name: string;
  /** Stream data. */
  readonly data: Uint8Array;
  /**
   * Optional storage path for the stream. Each element is a storage
   * (directory) name; the stream lives inside the nested storages.
   * For example `path: ["\u0006DataSpaces", "TransformInfo"]` with
   * `name: "..."` places the stream at
   * `\u0006DataSpaces/TransformInfo/<name>`. Omit or use `[]` for a
   * top-level stream under the root storage.
   */
  readonly path?: readonly string[];
}

/** CFB file signature: D0CF11E0A1B11AE1. */
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Special sector values. */
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const _FATSECT = 0xfffffffd;

/**
 * Read a CFB (OLE2 Compound File) and extract all stream entries.
 *
 * Only reads v3 and v4 CFB files (sector sizes 512 and 4096).
 *
 * @param buffer - The CFB file data.
 * @returns Array of named stream entries.
 */
export function readCfb(buffer: Uint8Array): CfbEntry[] {
  if (buffer.length < 512) {
    throw new DocxParseError("CFB: file too small");
  }

  // Verify signature
  for (let i = 0; i < 8; i++) {
    if (buffer[i] !== CFB_SIGNATURE[i]) {
      throw new DocxParseError("CFB: invalid signature");
    }
  }

  const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);

  // Header fields
  const majorVersion = view.getUint16(26, true);
  const sectorSize = majorVersion === 4 ? 4096 : 512;
  const miniSectorSize = 64;
  const miniStreamCutoff = view.getUint32(56, true); // usually 4096

  const fatSectorCount = view.getUint32(44, true);
  const firstDirectorySector = view.getUint32(48, true);
  const firstMiniFatSector = view.getUint32(60, true);
  const miniFatSectorCount = view.getUint32(64, true);
  const firstDifatSector = view.getUint32(68, true);
  const difatSectorCount = view.getUint32(72, true);

  // Helper: sector offset in file
  const sectorOffset = (sector: number) => (sector + 1) * sectorSize;

  // Read DIFAT: first 109 entries in header, then chained
  const difat: number[] = [];
  for (let i = 0; i < 109 && difat.length < fatSectorCount; i++) {
    const val = view.getUint32(76 + i * 4, true);
    if (val !== FREESECT && val !== ENDOFCHAIN) {
      difat.push(val);
    }
  }
  // Follow DIFAT chain
  let difatSector = firstDifatSector;
  for (let i = 0; i < difatSectorCount && difatSector !== ENDOFCHAIN; i++) {
    const off = sectorOffset(difatSector);
    const entriesPerSector = sectorSize / 4 - 1; // last 4 bytes = next DIFAT sector
    for (let j = 0; j < entriesPerSector && difat.length < fatSectorCount; j++) {
      const val = view.getUint32(off + j * 4, true);
      if (val !== FREESECT && val !== ENDOFCHAIN) {
        difat.push(val);
      }
    }
    difatSector = view.getUint32(off + entriesPerSector * 4, true);
  }

  // Build FAT (sector allocation table).
  //
  // `fatSectorCount` comes from the file header and is attacker-controlled.
  // A hostile CFB advertising `fatSectorCount = 100_000` would push us to
  // build a FAT array with ~12.8M Uint32 entries (~100 MB at JS object
  // overhead). Cap it conservatively — real DOCX-encryption packages have
  // at most a few dozen FAT sectors.
  const MAX_FAT_ENTRIES = 1_000_000;
  const fat: number[] = [];
  for (const fatSec of difat) {
    const off = sectorOffset(fatSec);
    for (let i = 0; i < sectorSize / 4; i++) {
      if (fat.length >= MAX_FAT_ENTRIES) {
        throw new Error(
          `CFB FAT exceeds maximum entry count (${MAX_FAT_ENTRIES}). ` +
            `Aborting to avoid runaway memory consumption.`
        );
      }
      fat.push(view.getUint32(off + i * 4, true));
    }
  }

  // Read sector chain.
  //
  // Resource limits:
  //   - `visited` already protects against infinite cycles.
  //   - We additionally cap the total bytes assembled per chain at 256 MB
  //     to bound memory consumption when a malicious CFB advertises a
  //     legitimately-shaped but enormous chain. Real DOCX-encryption
  //     packages are well under this; the cap exists to fail fast on
  //     hostile input rather than burn memory.
  const MAX_CHAIN_BYTES = 256 * 1024 * 1024;
  const readChain = (startSector: number): Uint8Array => {
    const sectors: Uint8Array[] = [];
    let sector = startSector;
    let totalBytes = 0;
    const visited = new Set<number>();
    while (sector !== ENDOFCHAIN && sector < fat.length && !visited.has(sector)) {
      visited.add(sector);
      const off = sectorOffset(sector);
      if (off + sectorSize <= buffer.length) {
        sectors.push(buffer.slice(off, off + sectorSize));
        totalBytes += sectorSize;
        if (totalBytes > MAX_CHAIN_BYTES) {
          throw new Error(
            `CFB sector chain exceeds maximum size (${MAX_CHAIN_BYTES} bytes). ` +
              `Aborting to avoid runaway memory consumption.`
          );
        }
      }
      sector = fat[sector];
    }
    // Concatenate
    const total = sectors.reduce((sum, s) => sum + s.length, 0);
    const result = new Uint8Array(total);
    let pos = 0;
    for (const s of sectors) {
      result.set(s, pos);
      pos += s.length;
    }
    return result;
  };

  // Read directory entries
  const dirData = readChain(firstDirectorySector);
  const dirView = new DataView(dirData.buffer, dirData.byteOffset, dirData.byteLength);
  const dirEntrySize = 128;
  const numDirEntries = Math.floor(dirData.length / dirEntrySize);

  interface DirEntry {
    name: string;
    type: number; // 0=unknown, 1=storage, 2=stream, 5=root
    startSector: number;
    size: number;
  }

  const dirEntries: DirEntry[] = [];
  for (let i = 0; i < numDirEntries; i++) {
    const off = i * dirEntrySize;
    const nameSize = dirView.getUint16(off + 64, true); // bytes (including null terminator)
    const type = dirView.getUint8(off + 66);
    const startSector = dirView.getUint32(off + 116, true);
    const size =
      majorVersion === 4
        ? Number(dirView.getBigUint64(off + 120, true))
        : dirView.getUint32(off + 120, true);

    // Decode UTF-16LE name (nameSize includes null terminator)
    let name = "";
    const nameLen = Math.max(0, nameSize - 2); // exclude null terminator
    for (let j = 0; j < nameLen; j += 2) {
      name += String.fromCharCode(dirView.getUint16(off + j, true));
    }

    dirEntries.push({ name, type, startSector, size });
  }

  // Build mini-FAT if needed
  const miniFat: number[] = [];
  if (miniFatSectorCount > 0 && firstMiniFatSector !== ENDOFCHAIN) {
    const miniFatData = readChain(firstMiniFatSector);
    const mfView = new DataView(miniFatData.buffer, miniFatData.byteOffset, miniFatData.byteLength);
    for (let i = 0; i < miniFatData.length / 4; i++) {
      miniFat.push(mfView.getUint32(i * 4, true));
    }
  }

  // Read mini-stream (from root entry's chain)
  let miniStream: Uint8Array | undefined;
  const rootEntry = dirEntries[0];
  if (rootEntry && rootEntry.type === 5 && rootEntry.startSector !== ENDOFCHAIN) {
    miniStream = readChain(rootEntry.startSector);
  }

  // Read a stream from mini-sectors. The declared `size` originates from a
  // directory entry inside the file and must therefore be treated as
  // untrusted; cap it at the actual mini-stream length to prevent malformed
  // CFB containers from triggering oversized buffer allocations.
  const readMiniChain = (startSector: number, size: number): Uint8Array => {
    if (!miniStream) {
      return new Uint8Array(0);
    }
    const safeSize = Math.max(0, Math.min(size, miniStream.length));
    const result = new Uint8Array(safeSize);
    let sector = startSector;
    let pos = 0;
    const visited = new Set<number>();
    while (sector !== ENDOFCHAIN && pos < safeSize && !visited.has(sector)) {
      visited.add(sector);
      const off = sector * miniSectorSize;
      const copyLen = Math.min(miniSectorSize, safeSize - pos);
      if (off + copyLen <= miniStream.length) {
        result.set(miniStream.slice(off, off + copyLen), pos);
      }
      pos += copyLen;
      sector = sector < miniFat.length ? miniFat[sector] : ENDOFCHAIN;
    }
    return result;
  };

  // Extract stream entries
  const entries: CfbEntry[] = [];
  for (let i = 1; i < dirEntries.length; i++) {
    const entry = dirEntries[i];
    if (entry.type !== 2 || entry.name === "") {
      continue;
    } // only stream entries

    // The directory's declared stream size cannot exceed the size of the
    // backing storage. Without this clamp a malformed `size` field could
    // request a multi-GiB allocation in `slice(0, entry.size)`.
    const safeSize = Math.max(0, Math.min(entry.size, buffer.length));
    let data: Uint8Array;
    if (safeSize < miniStreamCutoff) {
      // Read from mini-stream
      data = readMiniChain(entry.startSector, safeSize);
    } else {
      // Read from regular sectors
      const raw = readChain(entry.startSector);
      data = raw.slice(0, safeSize);
    }

    entries.push({ name: entry.name, data });
  }

  return entries;
}

// =============================================================================
// CFB Writer (v3, sector size 512)
// =============================================================================

const MINI_SECTOR_SIZE = 64;
const MINI_STREAM_CUTOFF = 4096;
const NOSTREAM = 0xffffffff;

/** Internal directory-tree node used while assembling the CFB. */
interface DirNode {
  name: string;
  /** 1 = storage, 2 = stream, 5 = root. */
  type: number;
  data?: Uint8Array;
  children: DirNode[];
  // Filled in during serialization:
  left: number;
  right: number;
  child: number;
  startSector: number;
  size: number;
}

/** Create a directory node with the sibling/child pointers in their unset state. */
function makeDirNode(name: string, type: number, data?: Uint8Array): DirNode {
  return {
    name,
    type,
    data,
    children: [],
    left: NOSTREAM,
    right: NOSTREAM,
    child: NOSTREAM,
    startSector: ENDOFCHAIN,
    size: data?.length ?? 0
  };
}

/**
 * MS-CFB sibling ordering: compare by UTF-16 code-unit length first, then
 * by upper-cased code units. This is the exact ordering Office uses to lay
 * out the red-black directory tree; getting it wrong makes Word reject the
 * container even though the bytes are otherwise valid.
 */
function compareCfbNames(a: string, b: string): number {
  if (a.length !== b.length) {
    return a.length - b.length;
  }
  const ua = a.toUpperCase();
  const ub = b.toUpperCase();
  if (ua < ub) {
    return -1;
  }
  if (ua > ub) {
    return 1;
  }
  return 0;
}

/**
 * Write a set of named stream entries into a CFB (OLE2 Compound File) container.
 *
 * Produces a v3 CFB with 512-byte sectors. Streams smaller than 4096 bytes are
 * stored in the mini-stream (64-byte mini-sectors) exactly as Office does;
 * larger streams use regular sectors. Entries may declare a `path` to nest the
 * stream inside one or more storages — required for the `\u0006DataSpaces`
 * structure that Office demands in encrypted documents.
 *
 * @param entries - Named stream entries to include.
 * @returns The CFB file as a Uint8Array.
 */
export function writeCfb(entries: readonly CfbEntry[]): Uint8Array {
  const SECTOR_SIZE = 512;
  const DIR_ENTRY_SIZE = 128;

  // ---------------------------------------------------------------------------
  // 1. Build the directory tree (root storage + nested storages + streams).
  // ---------------------------------------------------------------------------
  const root = makeDirNode("Root Entry", 5);

  const getOrCreateStorage = (parent: DirNode, name: string): DirNode => {
    let node = parent.children.find(c => c.type === 1 && c.name === name);
    if (!node) {
      node = makeDirNode(name, 1);
      parent.children.push(node);
    }
    return node;
  };

  for (const entry of entries) {
    let parent = root;
    for (const seg of entry.path ?? []) {
      parent = getOrCreateStorage(parent, seg);
    }
    parent.children.push(makeDirNode(entry.name, 2, entry.data));
  }

  // ---------------------------------------------------------------------------
  // 2. Flatten the tree into a directory-entry array in DFS order and assign
  //    each node a directory index. Build the red-black sibling tree for each
  //    storage's children using the CFB name ordering.
  // ---------------------------------------------------------------------------
  const dir: DirNode[] = [];
  const assignIndices = (node: DirNode): void => {
    dir.push(node);
    for (const child of node.children) {
      assignIndices(child);
    }
  };
  assignIndices(root);
  const indexOf = new Map<DirNode, number>();
  dir.forEach((n, i) => indexOf.set(n, i));

  // Build a balanced BST over the (sorted) sibling list. We produce a valid
  // search tree; Office does not require strict red-black balancing, only a
  // consistent ordering, so a balanced BST is accepted.
  const buildSiblingTree = (siblings: DirNode[]): number => {
    const sorted = siblings.slice().sort((a, b) => compareCfbNames(a.name, b.name));
    const build = (lo: number, hi: number): number => {
      if (lo > hi) {
        return NOSTREAM;
      }
      const mid = (lo + hi) >> 1;
      const node = sorted[mid];
      node.left = build(lo, mid - 1);
      node.right = build(mid + 1, hi);
      return indexOf.get(node)!;
    };
    return build(0, sorted.length - 1);
  };

  // Root + every storage gets a child pointer to the tree root of its children.
  const linkChildren = (node: DirNode): void => {
    if (node.children.length > 0) {
      node.child = buildSiblingTree(node.children);
    }
    for (const child of node.children) {
      linkChildren(child);
    }
  };
  linkChildren(root);

  // ---------------------------------------------------------------------------
  // 3. Split streams into mini-stream (<4096) vs regular sectors. The
  //    mini-stream itself is a chain of regular sectors owned by the root entry.
  // ---------------------------------------------------------------------------
  const streamNodes = dir.filter(n => n.type === 2);
  const miniNodes = streamNodes.filter(n => n.size > 0 && n.size < MINI_STREAM_CUTOFF);
  const regularNodes = streamNodes.filter(n => n.size >= MINI_STREAM_CUTOFF);

  // Assemble the mini-stream and the mini-FAT.
  const miniFat: number[] = [];
  const miniStream = (() => {
    const parts: Uint8Array[] = [];
    let miniSectorIdx = 0;
    let totalLen = 0;
    for (const node of miniNodes) {
      const sectorCount = Math.ceil(node.size / MINI_SECTOR_SIZE);
      node.startSector = miniSectorIdx;
      for (let i = 0; i < sectorCount; i++) {
        miniFat.push(i < sectorCount - 1 ? miniSectorIdx + 1 : ENDOFCHAIN);
        miniSectorIdx++;
      }
      const padded = new Uint8Array(sectorCount * MINI_SECTOR_SIZE);
      padded.set(node.data!);
      parts.push(padded);
      totalLen += padded.length;
    }
    const stream = new Uint8Array(totalLen);
    let off = 0;
    for (const p of parts) {
      stream.set(p, off);
      off += p.length;
    }
    return stream;
  })();
  root.size = miniStream.length;

  // ---------------------------------------------------------------------------
  // 4. Lay out regular sectors.
  //    Layout order in file: [FAT sectors][Directory][MiniFAT][MiniStream][Regular streams]
  //    We compute counts first, then assign sector indices, then fill the FAT.
  // ---------------------------------------------------------------------------
  const dirSectors = Math.ceil((dir.length * DIR_ENTRY_SIZE) / SECTOR_SIZE);

  const miniFatBytes = miniFat.length * 4;
  const miniFatSectors = miniFatBytes > 0 ? Math.ceil(miniFatBytes / SECTOR_SIZE) : 0;

  const miniStreamSectors = miniStream.length > 0 ? Math.ceil(miniStream.length / SECTOR_SIZE) : 0;

  const regularSectorCounts = regularNodes.map(n => Math.ceil(n.size / SECTOR_SIZE));
  const totalRegularStreamSectors = regularSectorCounts.reduce((a, b) => a + b, 0);

  const nonFatSectors = dirSectors + miniFatSectors + miniStreamSectors + totalRegularStreamSectors;

  // Solve for the number of FAT sectors (each FAT sector indexes 128 sectors).
  let fatSectors = 1;
  while (true) {
    const totalSectors = nonFatSectors + fatSectors;
    if (fatSectors * (SECTOR_SIZE / 4) >= totalSectors) {
      break;
    }
    fatSectors++;
  }

  // The header stores up to 109 DIFAT entries inline; beyond that a DIFAT
  // sector chain is required, which this minimal writer does not emit. Bail
  // out rather than silently produce a corrupt container. 109 FAT sectors
  // cover ~6.8 MB of sectors → tens of MB of stream data, far larger than any
  // realistic EncryptedPackage.
  if (fatSectors > 109) {
    throw new DocxParseError(
      `CFB writer: ${fatSectors} FAT sectors exceeds the 109-entry header DIFAT ` +
        `limit (input too large for the minimal v3 writer).`
    );
  }

  const totalSectors = nonFatSectors + fatSectors;
  const fileSize = (1 + totalSectors) * SECTOR_SIZE; // +1 header sector
  const output = new Uint8Array(fileSize);
  const view = new DataView(output.buffer);

  // Assign sector ranges.
  let cursor = 0;
  const fatStart = cursor;
  cursor += fatSectors;
  const dirStart = cursor;
  cursor += dirSectors;
  const miniFatStart = miniFatSectors > 0 ? cursor : ENDOFCHAIN;
  cursor += miniFatSectors;
  const miniStreamStart = miniStreamSectors > 0 ? cursor : ENDOFCHAIN;
  cursor += miniStreamSectors;

  // Regular stream start sectors.
  for (let i = 0; i < regularNodes.length; i++) {
    regularNodes[i].startSector = cursor;
    cursor += regularSectorCounts[i];
  }
  root.startSector = miniStreamStart;

  // ---------------------------------------------------------------------------
  // 5. Header.
  // ---------------------------------------------------------------------------
  const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  for (let i = 0; i < 8; i++) {
    output[i] = sig[i];
  }
  view.setUint16(24, 0x003e, true); // minor version
  view.setUint16(26, 0x0003, true); // major version (v3)
  view.setUint16(28, 0xfffe, true); // byte order LE
  view.setUint16(30, 9, true); // sector shift (2^9 = 512)
  view.setUint16(32, 6, true); // mini sector shift (2^6 = 64)
  view.setUint32(40, 0, true); // number of directory sectors (v3: 0)
  view.setUint32(44, fatSectors, true);
  view.setUint32(48, dirStart, true); // first directory sector
  view.setUint32(52, 0, true); // transaction signature
  view.setUint32(56, MINI_STREAM_CUTOFF, true); // mini stream cutoff
  view.setUint32(60, miniFatStart, true); // first mini-FAT sector
  view.setUint32(64, miniFatSectors, true); // mini-FAT sector count
  view.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector
  view.setUint32(72, 0, true); // DIFAT sector count

  // DIFAT in header (109 entries): point at the FAT sectors.
  for (let i = 0; i < 109; i++) {
    view.setUint32(76 + i * 4, i < fatSectors ? fatStart + i : FREESECT, true);
  }

  // ---------------------------------------------------------------------------
  // 6. FAT.
  // ---------------------------------------------------------------------------
  const fatFileOffset = (1 + fatStart) * SECTOR_SIZE;
  const fatEntryCount = fatSectors * (SECTOR_SIZE / 4);
  const fatView = new DataView(output.buffer, fatFileOffset, fatSectors * SECTOR_SIZE);
  for (let i = 0; i < fatEntryCount; i++) {
    fatView.setUint32(i * 4, FREESECT, true);
  }

  // Helper: write a chain of `count` sectors starting at `start` into the FAT.
  const writeFatChain = (start: number, count: number): void => {
    for (let i = 0; i < count; i++) {
      const sec = start + i;
      fatView.setUint32(sec * 4, i < count - 1 ? sec + 1 : ENDOFCHAIN, true);
    }
  };

  // FAT sectors themselves are marked FATSECT (0xFFFFFFFD).
  for (let i = 0; i < fatSectors; i++) {
    fatView.setUint32((fatStart + i) * 4, 0xfffffffd, true);
  }
  writeFatChain(dirStart, dirSectors);
  if (miniFatSectors > 0) {
    writeFatChain(miniFatStart, miniFatSectors);
  }
  if (miniStreamSectors > 0) {
    writeFatChain(miniStreamStart, miniStreamSectors);
  }
  for (let i = 0; i < regularNodes.length; i++) {
    writeFatChain(regularNodes[i].startSector, regularSectorCounts[i]);
  }

  // ---------------------------------------------------------------------------
  // 7. Directory entries.
  // ---------------------------------------------------------------------------
  const dirFileOffset = (1 + dirStart) * SECTOR_SIZE;
  for (let i = 0; i < dir.length; i++) {
    const node = dir[i];
    const off = dirFileOffset + i * DIR_ENTRY_SIZE;

    // UTF-16LE name (max 31 chars + null terminator).
    const nameLen = Math.min(node.name.length, 31);
    for (let j = 0; j < nameLen; j++) {
      view.setUint16(off + j * 2, node.name.charCodeAt(j), true);
    }
    view.setUint16(off + nameLen * 2, 0, true); // null terminator
    view.setUint16(off + 64, (nameLen + 1) * 2, true); // name byte length

    output[off + 66] = node.type; // object type
    output[off + 67] = 1; // color = black
    view.setUint32(off + 68, node.left, true); // left sibling
    view.setUint32(off + 72, node.right, true); // right sibling
    view.setUint32(off + 76, node.child, true); // child

    // CLSID (16 bytes) left zero. State bits / timestamps left zero.
    view.setUint32(off + 116, node.startSector, true);
    // Size: 8 bytes in v4; in v3 the low 4 bytes hold the size and the high
    // 4 bytes must be zero (already zero-initialized).
    view.setUint32(off + 120, node.size, true);
  }
  // Pad any unused directory slots in the last directory sector with
  // free/unknown entries (type 0, siblings NOSTREAM) so readers don't trip.
  const dirSlots = dirSectors * (SECTOR_SIZE / DIR_ENTRY_SIZE);
  for (let i = dir.length; i < dirSlots; i++) {
    const off = dirFileOffset + i * DIR_ENTRY_SIZE;
    view.setUint32(off + 68, NOSTREAM, true);
    view.setUint32(off + 72, NOSTREAM, true);
    view.setUint32(off + 76, NOSTREAM, true);
  }

  // ---------------------------------------------------------------------------
  // 8. Mini-FAT.
  // ---------------------------------------------------------------------------
  if (miniFatSectors > 0) {
    const mfOffset = (1 + miniFatStart) * SECTOR_SIZE;
    const mfCapacity = miniFatSectors * (SECTOR_SIZE / 4);
    const mfView = new DataView(output.buffer, mfOffset, miniFatSectors * SECTOR_SIZE);
    for (let i = 0; i < mfCapacity; i++) {
      mfView.setUint32(i * 4, i < miniFat.length ? miniFat[i] : FREESECT, true);
    }
  }

  // ---------------------------------------------------------------------------
  // 9. Mini-stream data.
  // ---------------------------------------------------------------------------
  if (miniStreamSectors > 0) {
    output.set(miniStream, (1 + miniStreamStart) * SECTOR_SIZE);
  }

  // ---------------------------------------------------------------------------
  // 10. Regular stream data.
  // ---------------------------------------------------------------------------
  for (const node of regularNodes) {
    output.set(node.data!, (1 + node.startSector) * SECTOR_SIZE);
  }

  return output;
}
