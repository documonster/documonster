/**
 * `scripts/verify-published.ts` tests.
 *
 * This check has been wrong twice, and both times it failed *open* in the sense
 * that mattered: the release workflow reported a problem that did not exist,
 * while the case it was written to catch — a package published restricted, which
 * `pnpm publish` reports as success — was never exercised at all. The registry
 * cannot produce that on demand, so the script takes `--registry` and the
 * scenarios are served by a local HTTP server here.
 *
 * Run as a subprocess rather than imported, matching the other script tests:
 * `scripts/` is outside the typed project, and the exit code is part of what
 * CI depends on.
 */

import { execFile } from "node:child_process";
import { createServer, type Server } from "node:http";
import type { AddressInfo } from "node:net";
import path from "node:path";
import process from "node:process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const SCRIPT = path.resolve(
  import.meta.dirname,
  "..",
  "..",
  "..",
  "scripts",
  "verify-published.ts"
);

/** A JSON body this harness can serve. */
type JsonObject = Record<string, unknown>;

/**
 * One programmed reply for the version document.
 *
 * `body` may be a factory, because a realistic version document has to name a
 * tarball on *this* server and the port is only known once it is listening.
 */
interface Reply {
  readonly status: number;
  readonly body?: JsonObject | ((registry: string) => JsonObject);
}

interface Harness {
  readonly registry: string;
  /** How many times the version document was requested. */
  versionRequests: number;
  /** How many times the tarball was HEADed. */
  tarballRequests: number;
  close(): Promise<void>;
}

/**
 * Serve `replies` in order for the version document, repeating the last one, and
 * answer the tarball HEAD with `tarballStatus`.
 */
async function harness(replies: readonly Reply[], tarballStatus = 200): Promise<Harness> {
  const state = { versionRequests: 0, tarballRequests: 0, registry: "" };
  const server: Server = createServer((request, response) => {
    if (request.url?.includes("/-/tarball")) {
      state.tarballRequests += 1;
      response.writeHead(tarballStatus).end();
      return;
    }
    const reply = replies[Math.min(state.versionRequests, replies.length - 1)]!;
    state.versionRequests += 1;
    if (reply.body === undefined) {
      response.writeHead(reply.status).end();
      return;
    }
    const body = typeof reply.body === "function" ? reply.body(state.registry) : reply.body;
    response.writeHead(reply.status, { "content-type": "application/json" });
    response.end(JSON.stringify(body));
  });

  await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
  const { port } = server.address() as AddressInfo;
  const registry = `http://127.0.0.1:${port}`;
  state.registry = registry;
  return {
    registry,
    get versionRequests() {
      return state.versionRequests;
    },
    get tarballRequests() {
      return state.tarballRequests;
    },
    close: () =>
      new Promise<void>(resolve => {
        server.close(() => resolve());
      })
  };
}

/** A well-formed version document pointing at the serving harness's tarball. */
function versionDocument(version: string): (registry: string) => JsonObject {
  return registry => ({ name: "pkg", version, dist: { tarball: `${registry}/-/tarball` } });
}

interface Run {
  readonly code: number;
  readonly stdout: string;
  readonly stderr: string;
}

async function run(args: readonly string[]): Promise<Run> {
  try {
    const { stdout, stderr } = await execFileAsync(process.execPath, [SCRIPT, ...args]);
    return { code: 0, stdout, stderr };
  } catch (cause) {
    const failure = cause as { code?: number; stdout?: string; stderr?: string };
    return { code: failure.code ?? -1, stdout: failure.stdout ?? "", stderr: failure.stderr ?? "" };
  }
}

describe("verify-published", () => {
  it("passes when the version and its tarball are both readable", async () => {
    const live = await harness([{ status: 200, body: versionDocument("1.0.0") }]);
    try {
      const result = await run(["--registry", live.registry, "--interval-ms", "1", "pkg@1.0.0"]);
      expect(result.code).toBe(0);
      expect(result.stdout).toContain("1 package version(s) live");
      // The tarball must actually be checked: a version document naming an
      // unfetchable tarball is not an installable release.
      expect(live.tarballRequests).toBe(1);
    } finally {
      await live.close();
    }
  });

  it("keeps polling through a 404 and succeeds once the version appears", async () => {
    // The defect this script replaces: one sample taken 30s after publish, when
    // the real propagation measured 4m11s.
    const server = await harness([
      { status: 404 },
      { status: 404 },
      { status: 200, body: versionDocument("1.0.0") }
    ]);
    try {
      const result = await run([
        "--registry",
        server.registry,
        "--attempts",
        "5",
        "--interval-ms",
        "1",
        "pkg@1.0.0"
      ]);
      expect(result.code).toBe(0);
      expect(result.stdout).toContain("attempt 3/5");
      expect(server.versionRequests).toBe(3);
    } finally {
      await server.close();
    }
  });

  it("fails immediately on a 403 instead of retrying", async () => {
    // A restricted publish. `pnpm publish` calls that a success, an
    // authenticated read confirms it, and only an anonymous read shows the
    // package is invisible to everyone else. Retrying cannot change it, so the
    // request count is part of the assertion.
    const server = await harness([{ status: 403 }]);
    try {
      const result = await run([
        "--registry",
        server.registry,
        "--attempts",
        "5",
        "--interval-ms",
        "1",
        "pkg@1.0.0"
      ]);
      expect(result.code).toBe(1);
      expect(result.stderr).toContain("published but not readable anonymously");
      expect(result.stderr).toContain("--access public");
      expect(server.versionRequests).toBe(1);
    } finally {
      await server.close();
    }
  });

  it("fails when the budget runs out, naming the last status", async () => {
    const server = await harness([{ status: 404 }]);
    try {
      const result = await run([
        "--registry",
        server.registry,
        "--attempts",
        "3",
        "--interval-ms",
        "1",
        "pkg@1.0.0"
      ]);
      expect(result.code).toBe(1);
      expect(result.stderr).toContain("did not become readable within 3 attempt(s)");
      expect(result.stderr).toContain("HTTP 404");
      expect(server.versionRequests).toBe(3);
    } finally {
      await server.close();
    }
  });

  it("rejects a version document that answers for a different version", async () => {
    const server = await harness([{ status: 200, body: versionDocument("9.9.9") }]);
    try {
      const result = await run(["--registry", server.registry, "--interval-ms", "1", "pkg@1.0.0"]);
      expect(result.code).toBe(1);
      expect(result.stderr).toContain('returned version "9.9.9"');
    } finally {
      await server.close();
    }
  });

  it("rejects a version whose tarball cannot be fetched", async () => {
    const server = await harness([{ status: 200, body: versionDocument("1.0.0") }], 404);
    try {
      const result = await run([
        "--registry",
        server.registry,
        "--attempts",
        "2",
        "--interval-ms",
        "1",
        "pkg@1.0.0"
      ]);
      expect(result.code).toBe(1);
      expect(result.stderr).toContain("tarball not fetchable");
    } finally {
      await server.close();
    }
  });

  it("verifies every spec and reports all the failures", async () => {
    const server = await harness([{ status: 404 }]);
    try {
      const result = await run([
        "--registry",
        server.registry,
        "--attempts",
        "1",
        "--interval-ms",
        "1",
        "documonster@1.0.0",
        "@documonster/mcp@1.0.0"
      ]);
      expect(result.code).toBe(1);
      expect(result.stderr).toContain("documonster@1.0.0");
      expect(result.stderr).toContain("@documonster/mcp@1.0.0");
    } finally {
      await server.close();
    }
  });

  it("encodes a scope in the registry path", async () => {
    // `@documonster/mcp` is `@documonster%2Fmcp` in a registry URL; sending the
    // raw slash asks for a different path entirely.
    const requested: string[] = [];
    const server = createServer((request, response) => {
      requested.push(request.url ?? "");
      response.writeHead(404).end();
    });
    await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
    const { port } = server.address() as AddressInfo;
    try {
      await run([
        "--registry",
        `http://127.0.0.1:${port}`,
        "--attempts",
        "1",
        "--interval-ms",
        "1",
        "@documonster/mcp@1.0.0"
      ]);
      expect(requested[0]).toBe("/@documonster%2Fmcp/1.0.0");
    } finally {
      await new Promise<void>(resolve => server.close(() => resolve()));
    }
  });

  describe("--exists", () => {
    // The publish steps' idempotency guard. Its failure mode is the opposite of
    // the verification's: a guard that misses an already-published version makes
    // the workflow publish over it, and npm rejects that outright — turning a
    // re-run after a partial release into a hard failure.
    it("exits 0 when the version is there", async () => {
      const server = await harness([{ status: 200, body: versionDocument("1.0.0") }]);
      try {
        const result = await run(["--registry", server.registry, "--exists", "pkg@1.0.0"]);
        expect(result.code).toBe(0);
        // No polling, and no tarball fetch: existence is all that is asked.
        expect(server.versionRequests).toBe(1);
        expect(server.tarballRequests).toBe(0);
      } finally {
        await server.close();
      }
    });

    it("exits 1 when it is absent", async () => {
      const server = await harness([{ status: 404 }]);
      try {
        const result = await run(["--registry", server.registry, "--exists", "pkg@1.0.0"]);
        expect(result.code).toBe(1);
        expect(server.versionRequests).toBe(1);
      } finally {
        await server.close();
      }
    });

    it("treats a 403 as present, so a restricted version is not republished", async () => {
      // The version exists and is merely unreadable anonymously. Publishing
      // again would fail with "cannot publish over the previously published
      // version", so this must report present even though the verification
      // above treats the same status as a failure.
      const server = await harness([{ status: 403 }]);
      try {
        const result = await run(["--registry", server.registry, "--exists", "pkg@1.0.0"]);
        expect(result.code).toBe(0);
      } finally {
        await server.close();
      }
    });

    it("reports absent when the answer is for another version", async () => {
      const server = await harness([{ status: 200, body: versionDocument("9.9.9") }]);
      try {
        const result = await run(["--registry", server.registry, "--exists", "pkg@1.0.0"]);
        expect(result.code).toBe(1);
      } finally {
        await server.close();
      }
    });

    it("takes exactly one spec", async () => {
      const result = await run(["--exists", "a@1.0.0", "b@1.0.0"]);
      expect(result.code).toBe(1);
      expect(result.stderr).toContain("exactly one");
    });
  });

  it("rejects an unparseable spec and an unknown flag", async () => {
    const bare = await run(["nope"]);
    expect(bare.code).toBe(1);
    expect(bare.stderr).toContain("not a <name>@<version> spec");

    const flag = await run(["--nope", "pkg@1.0.0"]);
    expect(flag.code).toBe(1);
    expect(flag.stderr).toContain("unknown flag");

    const none = await run([]);
    expect(none.code).toBe(1);
    expect(none.stderr).toContain("usage:");
  });
});
