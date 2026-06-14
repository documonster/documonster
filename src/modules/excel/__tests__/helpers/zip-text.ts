import { extractAll, type ExtractedFile } from "@archive/unzip/extract";
/**
 * Shared ZIP/text helpers for chart integration tests.
 *
 * Centralises the trivially-duplicated `entryText` / `decoder` /
 * `bytesEqual` / `loadRoundTrip` utilities that were previously copy-
 * pasted across every chart integration test file.
 */
import { Workbook } from "@excel/index";
import type { WorkbookData } from "@excel/workbook-core";

const decoder = new TextDecoder();

export type EntryMap = Map<string, ExtractedFile>;

/**
 * Decode a ZIP entry as UTF-8 text, or return undefined when the entry
 * is absent. Non-existence is a deliberate non-throw — callers commonly
 * probe for optional sidecars (e.g. `xl/charts/style1.xml`).
 */
export function entryText(entries: EntryMap, path: string): string | undefined {
  const entry = entries.get(path);
  return entry ? decoder.decode(entry.data) : undefined;
}

/**
 * Strict variant: throws if the entry is missing. Use when the entry is
 * expected to exist and its absence is a test failure rather than a
 * branch.
 */
export function requireEntryText(entries: EntryMap, path: string): string {
  const text = entryText(entries, path);
  if (text === undefined) {
    throw new Error(`required ZIP entry not found: ${path}`);
  }
  return text;
}

export function bytesEqual(a: Uint8Array | undefined, b: Uint8Array | undefined): boolean {
  if (!a || !b) {
    return a === b;
  }
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

/**
 * Load a workbook, write it back, and return the resulting workbook +
 * extracted ZIP entries. Used by the chart round-trip tests as the
 * shared "happy-path" load → write idiom.
 */
export async function loadRoundTrip(bytes: Uint8Array): Promise<{
  wb: WorkbookData;
  bytes: Uint8Array;
  entries: EntryMap;
}> {
  const wb = Workbook.create();
  await Workbook.loadXlsx(wb, bytes);
  const out = new Uint8Array(await Workbook.toXlsxBuffer(wb));
  const entries = await extractAll(out);
  return { wb, bytes: out, entries };
}

/**
 * Load → mutate → write variant. Returns the original (`before`) and
 * post-mutation (`after`) entry maps so callers can compare passthrough
 * fidelity.
 */
export async function loadRoundTripDiff(
  bytes: Uint8Array,
  mutate: (wb: WorkbookData) => void | Promise<void>
): Promise<{ before: EntryMap; after: EntryMap }> {
  const wb = Workbook.create();
  await Workbook.loadXlsx(wb, bytes);
  await mutate(wb);
  const out = new Uint8Array(await Workbook.toXlsxBuffer(wb));
  return {
    before: await extractAll(bytes),
    after: await extractAll(out)
  };
}
