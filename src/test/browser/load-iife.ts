/**
 * Browser-test helper: load a built IIFE bundle from `dist/iife/` by injecting
 * a `<script>` tag, then return the namespace it exposes under the shared
 * `Documonster` global.
 *
 * Each public module ships its own IIFE (see `rolldown.config.ts`), so a smoke
 * test can load just the bundle it needs and assert the delivered artifact
 * actually runs in a real browser — catching bundling/runtime regressions that
 * the source-import browser tests cannot (e.g. an accidental Node-only API
 * reference surviving into the shipped bundle).
 */

/** The `Documonster` global the IIFE bundles extend. */
interface DocumonsterGlobal {
  [namespace: string]: Record<string, unknown> | undefined;
}

const loaded = new Map<string, Promise<void>>();

/**
 * Inject `dist/iife/documonster.<file>.iife.min.js` once and resolve when it
 * has executed. Repeated calls for the same file share a single load.
 */
export function loadIifeScript(file: string): Promise<void> {
  let pending = loaded.get(file);
  if (!pending) {
    pending = new Promise<void>((resolve, reject) => {
      const script = document.createElement("script");
      script.src = `/dist/iife/documonster.${file}.iife.min.js`;
      script.onload = () => resolve();
      script.onerror = e => {
        // eslint-disable-next-line no-console
        console.error(`Failed to load IIFE bundle "${file}":`, e);
        reject(new Error(`Failed to load IIFE bundle "${file}"`));
      };
      document.head.appendChild(script);
    });
    loaded.set(file, pending);
  }
  return pending;
}

/**
 * Load an IIFE bundle and return its `Documonster.<namespace>` object,
 * throwing if the namespace is missing after the script executes.
 *
 * @param file - Bundle basename, e.g. `"word"` for `documonster.word.iife.min.js`.
 * @param namespace - Global member, e.g. `"Word"` for `Documonster.Word`.
 */
export async function loadIife<T = Record<string, unknown>>(
  file: string,
  namespace: string
): Promise<T> {
  await loadIifeScript(file);
  const root = (globalThis as unknown as { Documonster?: DocumonsterGlobal }).Documonster;
  const ns = root?.[namespace];
  if (!ns) {
    throw new Error(`IIFE bundle "${file}" loaded but Documonster.${namespace} is not defined`);
  }
  return ns as T;
}
