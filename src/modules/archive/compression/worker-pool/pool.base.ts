/**
 * Worker Pool Base Module
 *
 * Environment-agnostic interfaces and utilities for worker pool management.
 */

import type { WorkerPoolOptions } from "@archive/compression/worker-pool/types";

export type { WorkerPoolOptions };

/**
 * Resolved worker pool options with all values defined
 */
export interface ResolvedPoolOptions {
  maxWorkers: number;
  minWorkers: number;
  idleTimeout: number;
  useTransferables: boolean;
  workerUrl: string | undefined;
}

/**
 * Default worker pool options
 */
export const DEFAULT_POOL_OPTIONS: ResolvedPoolOptions = {
  maxWorkers: typeof navigator !== "undefined" ? navigator.hardwareConcurrency || 4 : 4,
  minWorkers: 0,
  idleTimeout: 30_000,
  useTransferables: true,
  workerUrl: undefined
};

/**
 * Resolve worker pool options with defaults
 */
export function resolvePoolOptions(options?: WorkerPoolOptions): ResolvedPoolOptions {
  const maxWorkers = options?.maxWorkers ?? DEFAULT_POOL_OPTIONS.maxWorkers;
  const minWorkers = options?.minWorkers ?? DEFAULT_POOL_OPTIONS.minWorkers;

  // Validate: minWorkers should not exceed maxWorkers
  const validatedMinWorkers = Math.min(minWorkers, maxWorkers);

  return {
    maxWorkers: Math.max(1, maxWorkers), // At least 1 worker
    minWorkers: Math.max(0, validatedMinWorkers),
    idleTimeout: Math.max(0, options?.idleTimeout ?? DEFAULT_POOL_OPTIONS.idleTimeout),
    useTransferables: options?.useTransferables ?? DEFAULT_POOL_OPTIONS.useTransferables,
    workerUrl: options?.workerUrl
  };
}

/**
 * Task priority values for sorting (internal)
 */
const PRIORITY_VALUES = {
  high: 3,
  normal: 2,
  low: 1
} as const;

type PriorityKey = keyof typeof PRIORITY_VALUES;

/**
 * Get priority value for sorting
 */
export function getPriorityValue(priority?: string): number {
  if (!priority) {
    return PRIORITY_VALUES.normal;
  }
  return PRIORITY_VALUES[priority as PriorityKey] ?? PRIORITY_VALUES.normal;
}

/**
 * Check if Web Workers are available in the current environment
 */
export function hasWorkerSupport(): boolean {
  return typeof Worker !== "undefined" && typeof Blob !== "undefined";
}

// Re-export from shared errors module
export { createAbortError } from "@archive/core/errors";
