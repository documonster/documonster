/**
 * A published version is not live until an anonymous reader can fetch it.
 *
 * ## Why this needs a check, and why the previous one could not work
 *
 * `pnpm publish` prints `✅ Published` as soon as the registry accepts the
 * tarball, which is earlier than the point at which anyone can install it. So
 * the release workflow verifies the result rather than trusting that line — and
 * the verification it used was structurally unable to pass:
 *
 * ```sh
 * sleep 30
 * npm view "${PACKAGE}@${VERSION}" version 2>/dev/null || echo NOT_FOUND
 * ```
 *
 * `npm view pkg@version` does not read a per-version document. It fetches the
 * **packument** — the whole package metadata — and looks the version up inside
 * it. That document is CDN-cached with `cache-control: public, max-age=300`, so
 * a read taken any time in the five minutes after a publish can be served a copy
 * that predates it. Measured on the 0.14.0 canary: published at 00:29:43, the
 * packument's `last-modified` was 00:33:54 — **4m11s later** — and the check
 * sampled it once at 00:30:19. A 30-second wait could not have succeeded; the
 * two runs before this script landed both failed here while both packages were
 * in fact published, and the run before those passed only by winning the race
 * against a cache expiry.
 *
 * Three things follow, and this script does all three:
 *
 * 1. **Read the per-version document, not the packument.** `/<name>/<version>`
 *    is served `cf-cache-status: DYNAMIC` — uncached — so it answers about the
 *    registry's state rather than about a cached snapshot.
 * 2. **Poll.** The registry is eventually consistent, so a single sample is the
 *    wrong instrument no matter how long the preceding sleep is.
 * 3. **Distinguish the failures.** The old check collapsed every outcome into
 *    `NOT_FOUND` by discarding stderr, so "not indexed yet", "published
 *    restricted" and "network broke" were one message. A 403 means the version
 *    exists but is not public — retrying cannot fix it, and it is the exact
 *    failure `--access public` exists to prevent — so it fails immediately,
 *    while a 404 is retried.
 *
 * The request carries **no credentials**, deliberately: an authenticated read
 * succeeds for a restricted package, which would hide the one problem this is
 * most needed for. It also HEADs the tarball, because a version document naming
 * an unfetchable tarball is not an installable release.
 *
 * Usage:
 *   node scripts/verify-published.ts documonster@1.2.3 @documonster/mcp@1.2.3
 *   node scripts/verify-published.ts --attempts 20 --interval-ms 15000 pkg@1.0.0
 *   node scripts/verify-published.ts --registry http://localhost:8080 pkg@1.0.0
 */

import path from "node:path";

/** Where to look. Overridable so a test can point at a local server. */
const DEFAULT_REGISTRY = "https://registry.npmjs.org";
/**
 * Polling budget. 20 × 15s = 5 minutes, chosen to exceed the worst propagation
 * this repository has measured (4m11s) with room to spare, and to outlast the
 * packument's own 300-second TTL even though this script does not read it.
 */
const DEFAULT_ATTEMPTS = 20;
const DEFAULT_INTERVAL_MS = 15_000;

export interface VerifyOptions {
  readonly registry?: string;
  readonly attempts?: number;
  readonly intervalMs?: number;
  readonly log?: (line: string) => void;
  /** Injectable so a test does not wait in real time. */
  readonly sleep?: (ms: number) => Promise<void>;
}

export interface PackageSpec {
  readonly name: string;
  readonly version: string;
}

export type VerifyResult =
  | { readonly ok: true; readonly attempts: number }
  | { readonly ok: false; readonly problem: string };

/**
 * Split `name@version`, allowing for a scope.
 *
 * The separator is the *last* `@`, because `@documonster/mcp@1.0.0` has two and
 * the first belongs to the scope.
 */
export function parseSpec(spec: string): PackageSpec {
  const at = spec.lastIndexOf("@");
  if (at <= 0) {
    throw new Error(`not a <name>@<version> spec: ${spec}`);
  }
  const name = spec.slice(0, at);
  const version = spec.slice(at + 1);
  if (name.length === 0 || version.length === 0) {
    throw new Error(`not a <name>@<version> spec: ${spec}`);
  }
  return { name, version };
}

/** A scope's `/` is encoded in a registry path. */
function versionUrl(registry: string, spec: PackageSpec): string {
  return `${registry.replace(/\/$/, "")}/${spec.name.replace("/", "%2F")}/${encodeURIComponent(spec.version)}`;
}

/**
 * Poll until an anonymous reader can fetch the version and its tarball.
 *
 * Returns rather than throws, so the CLI can report every spec instead of
 * stopping at the first one.
 */
export async function verifyPublished(
  spec: PackageSpec,
  options: VerifyOptions = {}
): Promise<VerifyResult> {
  const registry = options.registry ?? DEFAULT_REGISTRY;
  const attempts = options.attempts ?? DEFAULT_ATTEMPTS;
  const intervalMs = options.intervalMs ?? DEFAULT_INTERVAL_MS;
  const log = options.log ?? ((line: string) => console.log(line));
  const sleep =
    options.sleep ?? ((ms: number) => new Promise<void>(resolve => setTimeout(resolve, ms)));

  const url = versionUrl(registry, spec);
  const label = `${spec.name}@${spec.version}`;
  let last = "no attempt made";

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    let response: Response;
    try {
      response = await fetch(url, {
        headers: { accept: "application/json", "cache-control": "no-cache" }
      });
    } catch (cause) {
      last = `request failed: ${cause instanceof Error ? cause.message : String(cause)}`;
      if (attempt < attempts) {
        await sleep(intervalMs);
      }
      continue;
    }

    if (response.status === 401 || response.status === 403) {
      // Not a propagation delay: the version is there and the registry is
      // refusing to serve it anonymously. Retrying cannot change that.
      return {
        ok: false,
        problem:
          `${label} is published but not readable anonymously (HTTP ${response.status}). ` +
          `It was almost certainly published restricted — check with ` +
          `\`npm access get status ${spec.name}\` and republish with \`--access public\`.`
      };
    }

    if (response.status === 404) {
      last = "HTTP 404 — not indexed yet";
      log(`  ${label}: not there yet (attempt ${attempt}/${attempts})`);
      if (attempt < attempts) {
        await sleep(intervalMs);
      }
      continue;
    }

    if (!response.ok) {
      last = `HTTP ${response.status}`;
      log(`  ${label}: ${last} (attempt ${attempt}/${attempts})`);
      if (attempt < attempts) {
        await sleep(intervalMs);
      }
      continue;
    }

    const document = (await response.json()) as { version?: unknown; dist?: { tarball?: unknown } };
    if (document.version !== spec.version) {
      // The registry answered for a different version than the path asked for,
      // which would make every later assertion meaningless.
      return {
        ok: false,
        problem: `${label}: the registry returned version ${JSON.stringify(document.version)}`
      };
    }

    const tarball = document.dist?.tarball;
    if (typeof tarball !== "string" || tarball.length === 0) {
      return { ok: false, problem: `${label}: the version document names no tarball` };
    }

    // A version document naming a tarball nobody can fetch is not a release.
    const head = await fetch(tarball, { method: "HEAD" }).catch(() => undefined);
    if (head === undefined || !head.ok) {
      last = `tarball not fetchable (${head === undefined ? "request failed" : `HTTP ${head.status}`})`;
      log(`  ${label}: ${last} (attempt ${attempt}/${attempts})`);
      if (attempt < attempts) {
        await sleep(intervalMs);
      }
      continue;
    }

    log(`✓ ${label} is live and installable (attempt ${attempt}/${attempts})`);
    return { ok: true, attempts: attempt };
  }

  return {
    ok: false,
    problem:
      `${label} did not become readable within ${attempts} attempt(s) ` +
      `(last: ${last}). Anonymous reads only — if the package published ` +
      `successfully, check that it is public: \`npm access get status ${spec.name}\`.`
  };
}

/**
 * Does this exact version exist on the registry at all?
 *
 * Used by the publish steps' idempotency guard, which has the same stale-read
 * problem as the verification above but a worse consequence: a guard that fails
 * to see an already-published version makes the workflow publish over it, and
 * npm answers that with `cannot publish over the previously published version`.
 * That turns a re-run after a partial release — the case the guard exists for —
 * into a hard failure.
 *
 * Unlike {@link verifyPublished} this makes one request, does not fetch the
 * tarball, and treats 401/403 as **present**: the version is there and merely
 * not readable anonymously, so republishing would be rejected. Anything
 * indeterminate reports absent, which preserves the previous behaviour of
 * attempting the publish and letting the registry arbitrate.
 */
export async function versionExists(
  spec: PackageSpec,
  options: Pick<VerifyOptions, "registry"> = {}
): Promise<boolean> {
  const url = versionUrl(options.registry ?? DEFAULT_REGISTRY, spec);
  try {
    const response = await fetch(url, {
      headers: { accept: "application/json", "cache-control": "no-cache" }
    });
    if (response.status === 401 || response.status === 403) {
      return true;
    }
    if (!response.ok) {
      return false;
    }
    const document = (await response.json()) as { version?: unknown };
    return document.version === spec.version;
  } catch {
    return false;
  }
}

interface ParsedArgs {
  readonly specs: readonly string[];
  readonly exists: boolean;
  readonly options: VerifyOptions;
}

function parseArgs(argv: readonly string[]): ParsedArgs {
  const specs: string[] = [];
  let registry: string | undefined;
  let attempts: number | undefined;
  let intervalMs: number | undefined;
  let exists = false;

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]!;
    const next = (): string => {
      const value = argv[index + 1];
      if (value === undefined) {
        throw new Error(`${arg} needs a value`);
      }
      index += 1;
      return value;
    };
    if (arg === "--registry") {
      registry = next();
    } else if (arg === "--attempts") {
      attempts = Number(next());
    } else if (arg === "--interval-ms") {
      intervalMs = Number(next());
    } else if (arg === "--exists") {
      exists = true;
    } else if (arg.startsWith("--")) {
      throw new Error(`unknown flag: ${arg}`);
    } else {
      specs.push(arg);
    }
  }

  if (specs.length === 0) {
    throw new Error("usage: node scripts/verify-published.ts [--exists] <name>@<version> [...]");
  }
  if (exists && specs.length !== 1) {
    throw new Error("--exists takes exactly one <name>@<version>");
  }
  return {
    specs,
    exists,
    options: {
      ...(registry === undefined ? {} : { registry }),
      ...(attempts === undefined ? {} : { attempts }),
      ...(intervalMs === undefined ? {} : { intervalMs })
    }
  };
}

async function main(): Promise<void> {
  const { specs, exists, options } = parseArgs(process.argv.slice(2));

  if (exists) {
    const present = await versionExists(parseSpec(specs[0]!), options);
    process.exitCode = present ? 0 : 1;
    return;
  }

  const problems: string[] = [];
  for (const spec of specs) {
    const result = await verifyPublished(parseSpec(spec), options);
    if (!result.ok) {
      problems.push(result.problem);
    }
  }

  if (problems.length > 0) {
    for (const problem of problems) {
      console.error(`✗ ${problem}`);
    }
    process.exitCode = 1;
    return;
  }
  console.log(`✓ verify:published — ${specs.length} package version(s) live on the registry.`);
}

// Only run when invoked directly, so importing it for tests does not exit the process.
if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(import.meta.filename)) {
  void main();
}
