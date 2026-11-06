/**
 * The metadata part, `xl/metadata.bin`, as far as dynamic arrays need it.
 *
 * Excel writes this part for a workbook holding a dynamic array, and its content does not depend on how many there
 * are: one metadata type named `XLDAPR`, one formula-metadata block whose `BrtBeginDynamicArrayPr` scope carries
 * `fDynamic`, and one cell-metadata block naming that type. Every dynamic-array master then points at it with a
 * `BrtCellMeta` of 1. The records below are Excel's own — `data/dynamic-array.excel.xlsb`, saved by Excel 16.114 —
 * and the test that compares the part byte for byte is `dynamic-array.node.test.ts`.
 *
 * The XLSX reader and writer treat `xl/metadata.xml` the same way: it is read for the dynamic-array cell metadata it
 * declares and regenerated on write, never preserved.
 */
import {
  encodeBiffRecords,
  encodeWideString,
  iterateInterpretableRecords,
  XlsbReader,
  readWideString
} from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { record } from "@excel/xlsb/write/emit";
import { BinaryWriter, concatUint8Arrays } from "@utils/binary";

/** The metadata type name Excel gives dynamic-array formula metadata. */
const DYNAMIC_ARRAY_TYPE = "XLDAPR";

/**
 * `BrtMdtinfo`'s flag word and `metadataID` for `XLDAPR`, as Excel writes them.
 *
 * `0xD86AC0B0` sets the merge/copy/adjust bits Excel assigns this type and `120000` is its metadata id; neither is
 * derived from anything in the workbook.
 */
const DYNAMIC_ARRAY_TYPE_FLAGS = 0xd86ac0b0;
const DYNAMIC_ARRAY_TYPE_ID = 120000;

/** `BrtFRTBegin`'s `ProductVersion`: product 2, file version 2, as Excel writes it around the `XLDAPR` block. */
const DYNAMIC_ARRAY_FRT_VERSION = 0x00020002;

/** The bytes of `xl/metadata.bin` for a workbook that holds dynamic arrays. */
export function writeDynamicArrayMetadata(): Uint8Array {
  const u32 = (...values: number[]) => {
    const writer = new BinaryWriter();
    for (const value of values) {
      writer.writeUint32(value);
    }
    return writer.toUint8Array();
  };
  return encodeBiffRecords([
    record("BrtBeginMetadata"),
    record("BrtBeginEsmdtinfo", u32(1)),
    record(
      "BrtMdtinfo",
      concatUint8Arrays([
        u32(DYNAMIC_ARRAY_TYPE_FLAGS, DYNAMIC_ARRAY_TYPE_ID),
        encodeWideString(DYNAMIC_ARRAY_TYPE)
      ])
    ),
    record("BrtEndEsmdtinfo"),
    // One formula-metadata block of type 1 (`XLDAPR`, one-based).
    record("BrtBeginEsfmd", concatUint8Arrays([u32(1), encodeWideString(DYNAMIC_ARRAY_TYPE)])),
    record("BrtBeginFmd"),
    record("BrtFRTBegin", u32(DYNAMIC_ARRAY_FRT_VERSION)),
    record("BrtBeginDynamicArrayPr"),
    // `fDynamic` set, `fCollapsed` clear.
    record("BrtEndDynamicArrayPr", Uint8Array.of(0x01, 0x00)),
    record("BrtFRTEnd"),
    record("BrtEndFmd"),
    record("BrtEndEsfmd"),
    // One cell-metadata block (`fCellMeta`), holding one record: type 1, value index 0.
    record("BrtBeginEsmdb", u32(1, 1)),
    record("BrtMdb", u32(1, 1, 0)),
    record("BrtEndEsmdb"),
    record("BrtEndMetadata")
  ]);
}

/**
 * The cell-metadata indices (one-based, as `BrtCellMeta` names them) that mark a dynamic array.
 *
 * A cell-metadata block lists `Mdir` entries naming a metadata type; an entry whose type is `XLDAPR` makes the cell
 * that points at it a dynamic array — the reading the XLSX reader applies to `cm` through `metadata.xml`. Anything
 * this does not understand yields no index rather than a guess.
 */
export function readDynamicArrayCellMetadata(bytes: Uint8Array, part: string): Set<number> {
  const types: string[] = [];
  const indices = new Set<number>();
  let inCellMetadata = false;
  let cellMetadataIndex = 0;
  try {
    for (const entry of iterateInterpretableRecords(bytes, part)) {
      const name = recordSpec(entry.id)?.name;
      if (name === "BrtMdtinfo") {
        const reader = new XlsbReader(entry.payload, 0, part);
        reader.readUint32();
        reader.readUint32();
        types.push(readWideString(reader, part));
      } else if (name === "BrtBeginEsmdb") {
        // `fCellMeta` distinguishes cell metadata from value metadata; `BrtCellMeta` indexes only the former.
        inCellMetadata =
          entry.payload.byteLength >= 8 &&
          new DataView(entry.payload.buffer, entry.payload.byteOffset, 8).getUint32(4, true) !== 0;
      } else if (name === "BrtEndEsmdb") {
        inCellMetadata = false;
      } else if (name === "BrtMdb" && inCellMetadata) {
        cellMetadataIndex++;
        const reader = new XlsbReader(entry.payload, 0, part);
        const count = reader.readUint32();
        for (let index = 0; index < count; index++) {
          const type = reader.readUint32();
          reader.readUint32();
          if (types[type - 1] === DYNAMIC_ARRAY_TYPE) {
            indices.add(cellMetadataIndex);
          }
        }
      }
    }
  } catch {
    // A part that does not decode marks nothing: its cells read as ordinary array formulas.
  }
  return indices;
}
