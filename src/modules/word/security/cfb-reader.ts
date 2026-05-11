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

  // Build FAT (sector allocation table)
  const fat: number[] = [];
  for (const fatSec of difat) {
    const off = sectorOffset(fatSec);
    for (let i = 0; i < sectorSize / 4; i++) {
      fat.push(view.getUint32(off + i * 4, true));
    }
  }

  // Read sector chain
  const readChain = (startSector: number): Uint8Array => {
    const sectors: Uint8Array[] = [];
    let sector = startSector;
    const visited = new Set<number>();
    while (sector !== ENDOFCHAIN && sector < fat.length && !visited.has(sector)) {
      visited.add(sector);
      const off = sectorOffset(sector);
      if (off + sectorSize <= buffer.length) {
        sectors.push(buffer.slice(off, off + sectorSize));
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

  // Read a stream from mini-sectors
  const readMiniChain = (startSector: number, size: number): Uint8Array => {
    if (!miniStream) {
      return new Uint8Array(0);
    }
    const result = new Uint8Array(size);
    let sector = startSector;
    let pos = 0;
    const visited = new Set<number>();
    while (sector !== ENDOFCHAIN && pos < size && !visited.has(sector)) {
      visited.add(sector);
      const off = sector * miniSectorSize;
      const copyLen = Math.min(miniSectorSize, size - pos);
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

    let data: Uint8Array;
    if (entry.size < miniStreamCutoff) {
      // Read from mini-stream
      data = readMiniChain(entry.startSector, entry.size);
    } else {
      // Read from regular sectors
      const raw = readChain(entry.startSector);
      data = raw.slice(0, entry.size);
    }

    entries.push({ name: entry.name, data });
  }

  return entries;
}

// =============================================================================
// CFB Writer (v3, sector size 512)
// =============================================================================

/**
 * Write a set of named stream entries into a CFB (OLE2 Compound File) container.
 *
 * Produces a minimal v3 CFB with 512-byte sectors. Does not use mini-streams
 * (all data is stored in regular sectors). This is suitable for encrypted Office
 * documents where every stream exceeds 4096 bytes.
 *
 * @param entries - Named stream entries to include.
 * @returns The CFB file as a Uint8Array.
 */
export function writeCfb(entries: readonly CfbEntry[]): Uint8Array {
  const SECTOR_SIZE = 512;
  const DIR_ENTRY_SIZE = 128;

  // Compute sector count for each entry
  const entrySectors = entries.map(e => Math.ceil(e.data.length / SECTOR_SIZE));

  // Directory entries: Root Entry + one per stream entry
  const dirEntryCount = 1 + entries.length;
  const dirSectors = Math.ceil((dirEntryCount * DIR_ENTRY_SIZE) / SECTOR_SIZE);

  // Total data sectors
  const totalDataSectors = entrySectors.reduce((a, b) => a + b, 0);

  // FAT entries needed: directory sectors + data sectors + FAT sectors themselves
  // We solve iteratively since FAT sectors count depends on total sector count
  let fatSectors = 1;
  while (true) {
    const totalSectors = dirSectors + totalDataSectors + fatSectors;
    const fatCapacity = fatSectors * (SECTOR_SIZE / 4);
    if (fatCapacity >= totalSectors) {
      break;
    }
    fatSectors++;
  }

  const totalSectors = dirSectors + totalDataSectors + fatSectors;
  const fileSize = (1 + totalSectors) * SECTOR_SIZE; // +1 for header sector
  const output = new Uint8Array(fileSize);
  const view = new DataView(output.buffer);

  // --- Header (sector 0 area, 512 bytes) ---
  // Signature
  const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  for (let i = 0; i < 8; i++) {
    output[i] = sig[i];
  }

  // Minor version = 0x003E, Major version = 0x0003 (v3)
  view.setUint16(24, 0x003e, true);
  view.setUint16(26, 0x0003, true);

  // Byte order = 0xFFFE (little-endian)
  view.setUint16(28, 0xfffe, true);

  // Sector size power = 9 (2^9 = 512)
  view.setUint16(30, 9, true);

  // Mini sector size power = 6 (2^6 = 64)
  view.setUint16(32, 6, true);

  // Total sectors in directory (v3: must be 0)
  view.setUint32(40, 0, true);

  // Total FAT sectors
  view.setUint32(44, fatSectors, true);

  // First directory sector SECT
  // Layout: [FAT sectors] [Directory sectors] [Data sectors]
  const firstDirSector = fatSectors;
  view.setUint32(48, firstDirSector, true);

  // Transaction signature number
  view.setUint32(52, 0, true);

  // Mini stream cutoff = 0 (all streams stored in regular sectors)
  view.setUint32(56, 0, true);

  // First mini FAT sector = ENDOFCHAIN (none)
  view.setUint32(60, ENDOFCHAIN, true);

  // Mini FAT sector count = 0
  view.setUint32(64, 0, true);

  // First DIFAT sector = ENDOFCHAIN (none needed, <=109 FAT sectors)
  view.setUint32(68, ENDOFCHAIN, true);

  // DIFAT sector count = 0
  view.setUint32(72, 0, true);

  // DIFAT array in header (109 entries starting at offset 76)
  for (let i = 0; i < 109; i++) {
    view.setUint32(76 + i * 4, i < fatSectors ? i : FREESECT, true);
  }

  // --- Build FAT ---
  const fatOffset = SECTOR_SIZE; // FAT starts at sector 0 in file
  const fatView = new DataView(output.buffer, fatOffset, fatSectors * SECTOR_SIZE);

  // Initialize all FAT entries to FREESECT
  for (let i = 0; i < fatSectors * (SECTOR_SIZE / 4); i++) {
    fatView.setUint32(i * 4, FREESECT, true);
  }

  let sectorIdx = 0;

  // FAT sectors themselves are marked as 0xFFFFFFFD (FATSECT)
  for (let i = 0; i < fatSectors; i++) {
    fatView.setUint32(sectorIdx * 4, 0xfffffffd, true);
    sectorIdx++;
  }

  // Directory sectors chain
  for (let i = 0; i < dirSectors; i++) {
    const next = i < dirSectors - 1 ? sectorIdx + 1 : ENDOFCHAIN;
    fatView.setUint32(sectorIdx * 4, next, true);
    sectorIdx++;
  }

  // Data sectors for each entry
  const entryStartSectors: number[] = [];
  for (let e = 0; e < entries.length; e++) {
    entryStartSectors.push(sectorIdx);
    for (let i = 0; i < entrySectors[e]; i++) {
      const next = i < entrySectors[e] - 1 ? sectorIdx + 1 : ENDOFCHAIN;
      fatView.setUint32(sectorIdx * 4, next, true);
      sectorIdx++;
    }
  }

  // --- Write Directory ---
  const dirFileOffset = (1 + firstDirSector) * SECTOR_SIZE;

  // Helper: write UTF-16LE name into directory entry
  const writeDirName = (off: number, name: string): void => {
    for (let i = 0; i < name.length && i < 31; i++) {
      view.setUint16(off + i * 2, name.charCodeAt(i), true);
    }
    // Null terminator
    view.setUint16(off + name.length * 2, 0, true);
    // Name size in bytes (including null terminator)
    view.setUint16(off + 64, (name.length + 1) * 2, true);
  };

  // Root Entry (index 0)
  const rootOff = dirFileOffset;
  writeDirName(rootOff, "Root Entry");
  output[rootOff + 66] = 5; // type = root storage
  output[rootOff + 67] = 1; // color = black
  // Child (left/right/child SIDs) - set up a simple tree
  view.setUint32(rootOff + 68, 0xffffffff, true); // left sibling
  view.setUint32(rootOff + 72, 0xffffffff, true); // right sibling
  // Root entry child points to first entry (or 0xFFFFFFFF if none)
  if (entries.length > 0) {
    // Build a balanced-ish tree: use the middle entry as child
    const rootChild = entries.length === 1 ? 1 : Math.ceil(entries.length / 2);
    view.setUint32(rootOff + 76, rootChild, true);
  } else {
    view.setUint32(rootOff + 76, 0xffffffff, true);
  }
  // Start sector = ENDOFCHAIN (no mini-stream)
  view.setUint32(rootOff + 116, ENDOFCHAIN, true);
  // Size = 0
  view.setUint32(rootOff + 120, 0, true);

  // Stream entries (index 1..N)
  // Build as a red-black tree: simple approach — balanced binary tree
  // For small entry counts (2-3), just arrange siblings
  const buildTree = (indices: number[]): { leftSib: number; rightSib: number; root: number }[] => {
    // Map each entry to its left/right sibling
    const nodes: { leftSib: number; rightSib: number; root: number }[] = indices.map(idx => ({
      leftSib: 0xffffffff,
      rightSib: 0xffffffff,
      root: idx
    }));

    if (indices.length <= 1) {
      return nodes;
    }

    // Simple sorted insertion: entry at mid is root, left half is left subtree, right half is right subtree
    const assignTree = (arr: number[]): number => {
      if (arr.length === 0) {
        return 0xffffffff;
      }
      const mid = Math.floor(arr.length / 2);
      const midIdx = arr[mid];
      const nodeEntry = nodes.find(n => n.root === midIdx)!;
      nodeEntry.leftSib = assignTree(arr.slice(0, mid));
      nodeEntry.rightSib = assignTree(arr.slice(mid + 1));
      return midIdx;
    };

    assignTree(indices);
    return nodes;
  };

  const dirIndices = entries.map((_, i) => i + 1); // 1-based directory indices
  const treeNodes = buildTree(dirIndices);

  for (let e = 0; e < entries.length; e++) {
    const entryOff = dirFileOffset + (e + 1) * DIR_ENTRY_SIZE;
    writeDirName(entryOff, entries[e].name);
    output[entryOff + 66] = 2; // type = stream
    output[entryOff + 67] = 1; // color = black

    const node = treeNodes.find(n => n.root === e + 1)!;
    view.setUint32(entryOff + 68, node.leftSib, true); // left sibling
    view.setUint32(entryOff + 72, node.rightSib, true); // right sibling
    view.setUint32(entryOff + 76, 0xffffffff, true); // child (streams have none)

    // Start sector
    view.setUint32(entryOff + 116, entryStartSectors[e], true);
    // Size (32-bit for v3)
    view.setUint32(entryOff + 120, entries[e].data.length, true);
  }

  // Update root entry child to the tree root
  if (entries.length > 0) {
    const sortedIndices = dirIndices.slice();
    const mid = Math.floor(sortedIndices.length / 2);
    view.setUint32(rootOff + 76, sortedIndices[mid], true);
  }

  // --- Write Data Sectors ---
  for (let e = 0; e < entries.length; e++) {
    const dataFileOffset = (1 + entryStartSectors[e]) * SECTOR_SIZE;
    output.set(entries[e].data, dataFileOffset);
  }

  return output;
}
