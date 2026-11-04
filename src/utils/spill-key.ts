/**
 * The key a spilled cell names its dynamic-array master by: the master's worksheet id and position, stable across
 * renames.
 *
 * Shared by the formula engine, which plans spills in these keys, and the Excel host, which stores them on the cells
 * a spill covers and reads them back on load and write. At Layer 0 so the host's always-loaded parts (writing a row,
 * loading a sheet) can use it without bringing the formula engine into a bundle that never calculates.
 */
export function spillCellKeyFromId(worksheetId: number, row: number, col: number): string {
  return `ws:${worksheetId}!${row}:${col}`;
}

const SPILL_KEY = /^ws:(\d+)!(\d+):(\d+)$/;

/** The parts of a key {@link spillCellKeyFromId} made, or `undefined` for any other text. */
export function parseSpillCellKey(
  key: string
): { worksheetId: number; row: number; col: number } | undefined {
  const match = SPILL_KEY.exec(key);
  return match === null
    ? undefined
    : { worksheetId: Number(match[1]), row: Number(match[2]), col: Number(match[3]) };
}
