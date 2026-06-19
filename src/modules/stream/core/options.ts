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

export const isPipelineOptions = (value: unknown): value is PipelineOptions => {
  if (!value || typeof value !== "object") {
    return false;
  }

  // Avoid treating streams as options objects.
  if (
    typeof (value as any).pipe === "function" ||
    typeof (value as any).write === "function" ||
    typeof (value as any).end === "function" ||
    typeof (value as any).getReader === "function" ||
    typeof (value as any).getWriter === "function"
  ) {
    return false;
  }

  return (
    Object.prototype.hasOwnProperty.call(value, "signal") ||
    Object.prototype.hasOwnProperty.call(value, "end")
  );
};

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
