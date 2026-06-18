/**
 * CSV Worker Script - Lazy Loading Bundle
 *
 * Uses dynamic import to avoid bundling the 80KB+ worker script
 * into applications that don't use the worker pool functionality.
 */

// =============================================================================
// Blob URL Management (Lazy Loading)
// =============================================================================

let workerBlobUrl: string | null = null;
let workerBlobRefCount = 0;
let workerScriptPromise: Promise<string> | null = null;

/**
 * Lazily load the worker script.
 * Uses dynamic import to avoid bundling when not needed.
 */
function loadWorkerScript(): Promise<string> {
  if (!workerScriptPromise) {
    workerScriptPromise = import("@csv/worker/worker-script.generated").then(
      m => m.CSV_WORKER_SCRIPT
    );
  }
  return workerScriptPromise;
}

/** Get or create the worker blob URL */
export async function getWorkerBlobUrl(): Promise<string> {
  if (!workerBlobUrl) {
    const script = await loadWorkerScript();
    const blob = new Blob([script], { type: "application/javascript" });
    workerBlobUrl = URL.createObjectURL(blob);
  }
  workerBlobRefCount++;
  return workerBlobUrl;
}

/** Release the worker blob URL reference */
export function releaseWorkerBlobUrl(): void {
  workerBlobRefCount--;
  if (workerBlobRefCount <= 0 && workerBlobUrl) {
    URL.revokeObjectURL(workerBlobUrl);
    workerBlobUrl = null;
    workerBlobRefCount = 0;
  }
}
