/**
 * Shared test helper: validate a serialised xlsx buffer against the
 * OOXML validator and fail the test if any problems are found.
 *
 * Usage:
 *
 *   const buffer = await Workbook.toBuffer(wb);
 *   await expectValidXlsx(buffer);
 *
 * The helper accepts the output of `writeBuffer()` directly
 * (`ArrayBuffer` or `Uint8Array`) and throws an informative error on
 * any detected problem. The error message includes the full problem
 * list formatted for readability.
 *
 * This helper is the canonical unit-test integration point — all
 * integration tests in the excel module should call it after any
 * `writeBuffer()` / `write()` so the CI run doubles as a continuous
 * OOXML conformance audit.
 */

import { Workbook } from "@excel/index";
import type { OoxmlValidateOptions } from "@excel/utils/ooxml-validator";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import type { WorkbookData } from "@excel/workbook-core";
import { expect } from "vitest";

export interface ExpectValidXlsxOptions extends OoxmlValidateOptions {
  /**
   * A short context string included in the error message. Useful when
   * the same test has multiple `expectValidXlsx` calls for different
   * workbooks (e.g. "after mutation", "roundtrip pass 2").
   */
  label?: string;
}

/**
 * Validate the given xlsx bytes. Throws (via `expect().toBe()`) when
 * any error-severity problem is detected so vitest reports a clean
 * failure.
 */
export async function expectValidXlsx(
  buffer: ArrayBuffer | Uint8Array,
  options: ExpectValidXlsxOptions = {}
): Promise<void> {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const { label, ...validateOptions } = options;
  const report = await validateXlsxBuffer(bytes, { maxProblems: 50, ...validateOptions });
  if (!report.ok) {
    const header = label ? `[${label}] OOXML validation failed:` : "OOXML validation failed:";
    const formatted = report.problems
      .map(
        (p, i) =>
          `  ${i + 1}. [${p.severity}] ${p.kind} @ ${p.file ?? "<package>"}\n     ${p.message}`
      )
      .join("\n");
    // `expect(…).toBe(true)` with a Cause-carrying Error produces a tidy
    // vitest diff AND preserves the full problem list.
    expect.fail(`${header}\n${formatted}\n\nStats: ${JSON.stringify(report.stats)}`);
  }
}

/**
 * Like {@link expectValidXlsx} but writes the workbook first. Saves a
 * call to `writeBuffer()` in test code.
 *
 * ```ts
 * await expectValidWorkbook(wb);
 * ```
 */
export async function expectValidWorkbook(
  workbook: WorkbookData,
  options: ExpectValidXlsxOptions = {}
): Promise<Uint8Array> {
  const buffer = await Workbook.toBuffer(workbook);
  const bytes = new Uint8Array(buffer);
  await expectValidXlsx(bytes, options);
  return bytes;
}
