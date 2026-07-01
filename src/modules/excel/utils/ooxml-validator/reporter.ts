/**
 * Problem reporter shared by all checkers.
 *
 * Two responsibilities:
 *   1. Filter warnings when `includeWarnings` is `false`.
 *   2. Short-circuit after `maxProblems` is reached.
 *
 * The reporter does not sort or dedupe — checkers are expected to emit at
 * most one problem per logical issue. Callers get problems in encounter
 * order, which is stable across runs for the same input.
 */

import type {
  OoxmlProblemKind,
  OoxmlProblemSeverity,
  OoxmlValidationProblem
} from "@excel/utils/ooxml-validator/types";

export interface ReporterOptions {
  maxProblems?: number;
  includeWarnings?: boolean;
}

export class Reporter {
  readonly problems: OoxmlValidationProblem[] = [];
  private errorCount = 0;
  private readonly maxProblems?: number;
  private readonly includeWarnings: boolean;

  constructor(options: ReporterOptions = {}) {
    this.maxProblems = options.maxProblems;
    this.includeWarnings = options.includeWarnings ?? false;
  }

  /**
   * `true` once the configured `maxProblems` cap is reached. Checkers
   * should consult this at the top of their per-part loops to avoid
   * doing unnecessary work.
   */
  get capped(): boolean {
    return this.maxProblems !== undefined && this.problems.length >= this.maxProblems;
  }

  get hasErrors(): boolean {
    return this.errorCount > 0;
  }

  /** Report an error-severity problem. */
  error(kind: OoxmlProblemKind, message: string, file?: string): void {
    this.push({ kind, severity: "error", message, file });
    this.errorCount += this.capped ? 0 : 1;
  }

  /** Report a warning-severity problem. Only recorded if includeWarnings. */
  warning(kind: OoxmlProblemKind, message: string, file?: string): void {
    if (!this.includeWarnings) {
      return;
    }
    this.push({ kind, severity: "warning", message, file });
  }

  private push(problem: OoxmlValidationProblem): void {
    if (this.maxProblems !== undefined && this.problems.length >= this.maxProblems) {
      return;
    }
    this.problems.push(problem);
  }
}

/**
 * Convenience alias for the legacy `severity` argument used in the old
 * flat API. Kept for internal checker code that wants to produce a
 * problem with a runtime-chosen severity.
 */
export function reportWithSeverity(
  reporter: Reporter,
  severity: OoxmlProblemSeverity,
  kind: OoxmlProblemKind,
  message: string,
  file?: string
): void {
  if (severity === "error") {
    reporter.error(kind, message, file);
  } else {
    reporter.warning(kind, message, file);
  }
}
