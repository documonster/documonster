/**
 * Stream Module - Common Options
 *
 * Shared option interfaces and guards for pipeline/finished operations.
 * Used by both Node.js and Browser implementations.
 */

// =============================================================================
// Pipeline Options
// =============================================================================

export interface PipelineOptions {
  signal?: AbortSignal;
  end?: boolean;
}

export function isPipelineOptions(value: unknown): value is PipelineOptions {
  if (!value || typeof value !== "object") {
    return false;
  }

  // Avoid treating streams as options objects.
  const candidate = value as Record<string, unknown>;
  if (
    typeof candidate.pipe === "function" ||
    typeof candidate.write === "function" ||
    typeof candidate.end === "function" ||
    typeof candidate.getReader === "function" ||
    typeof candidate.getWriter === "function"
  ) {
    return false;
  }

  return (
    Object.prototype.hasOwnProperty.call(value, "signal") ||
    Object.prototype.hasOwnProperty.call(value, "end")
  );
}

// =============================================================================
// Finished Options
// =============================================================================

export interface FinishedOptions {
  readable?: boolean;
  writable?: boolean;
  error?: boolean;
  signal?: AbortSignal;
}

// =============================================================================
// Pipeline Callback
// =============================================================================

export type PipelineCallback = (err?: Error | null) => void;
