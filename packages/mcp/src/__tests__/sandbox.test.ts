import { mkdir, mkdtemp, realpath, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import type { ServerConfig } from "../config.js";
import { McpToolError } from "../errors.js";
import {
  assertWritable,
  isInside,
  reservedDeviceSegment,
  resolveInRoot,
  resolveOutputPath
} from "../sandbox.js";

async function makeConfig(overrides: Partial<ServerConfig> = {}): Promise<ServerConfig> {
  const root = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-sandbox-")));
  const outputRoot = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-output-")));
  return {
    root,
    outputRoot,
    allowInPlace: false,
    readonly: false,
    groups: new Set(["core"]),
    maxFileSize: 1024,
    maxOutputChars: 1000,
    ...overrides
  };
}

/** Assert a rejection carries a specific machine-readable code. */
async function expectCode(promise: Promise<unknown>, code: string): Promise<void> {
  await expect(promise).rejects.toSatisfy(
    (error: unknown) => error instanceof McpToolError && error.code === code,
    `expected an McpToolError with code "${code}"`
  );
}

describe("resolveInRoot", () => {
  it("resolves a relative path inside the root", async () => {
    const config = await makeConfig();
    await writeFile(path.join(config.root, "a.csv"), "x");
    await expect(resolveInRoot(config, "a.csv", { mustExist: true })).resolves.toBe(
      path.join(config.root, "a.csv")
    );
  });

  it("accepts an absolute path that lands inside the root", async () => {
    const config = await makeConfig();
    const target = path.join(config.root, "a.csv");
    await writeFile(target, "x");
    await expect(resolveInRoot(config, target, { mustExist: true })).resolves.toBe(target);
  });

  it("allows a not-yet-existing write target", async () => {
    const config = await makeConfig();
    await expect(resolveInRoot(config, "out/new.xlsx")).resolves.toBe(
      path.join(config.root, "out", "new.xlsx")
    );
  });

  it("rejects traversal above the root", async () => {
    const config = await makeConfig();
    await expectCode(resolveInRoot(config, "../escaped.txt"), "outside_root");
    await expectCode(resolveInRoot(config, "a/../../escaped.txt"), "outside_root");
  });

  it("rejects an absolute path outside the root", async () => {
    const config = await makeConfig();
    await expectCode(resolveInRoot(config, "/etc/passwd"), "outside_root");
  });

  it("rejects a symlink that points outside the root", async () => {
    // The critical case: the link itself is inside the root, so a naive
    // string prefix check would accept it. Only realpath catches this.
    const config = await makeConfig();
    const outside = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-outside-")));
    await writeFile(path.join(outside, "secret.txt"), "secret");
    await symlink(path.join(outside, "secret.txt"), path.join(config.root, "link.txt"));

    await expectCode(resolveInRoot(config, "link.txt", { mustExist: true }), "outside_root");
  });

  it("rejects a write target whose parent escapes via a symlinked directory", async () => {
    const config = await makeConfig();
    const outside = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-outside-")));
    await mkdir(path.join(outside, "sink"));
    // "junction" for directories: on Windows a plain symlink needs elevated
    // privileges, a junction does not, and both are what realpath must see
    // through. On POSIX the type argument is ignored.
    await symlink(path.join(outside, "sink"), path.join(config.root, "sink"), "junction");

    await expectCode(resolveInRoot(config, "sink/written.xlsx"), "outside_root");
  });

  it("rejects a URL but not a Windows drive letter shape", async () => {
    const config = await makeConfig();
    await expectCode(resolveInRoot(config, "https://example.com/a.xlsx"), "invalid_input");
    await expectCode(resolveInRoot(config, "file:///etc/passwd"), "invalid_input");
    // `C:` must not be mistaken for a URL scheme. What happens after that is
    // platform-specific and both outcomes are correct: on Windows the path is
    // drive-absolute and escapes the root, so it must be refused as
    // outside_root; on POSIX `C:` is an ordinary directory name and the path
    // stays inside. What this pins on either platform is that a drive letter is
    // never mistaken for a URL — and that Windows never lets it through.
    if (process.platform === "win32") {
      await expectCode(resolveInRoot(config, "C:/Windows/system.ini"), "outside_root");
    } else {
      await expect(resolveInRoot(config, "C:/Windows/system.ini")).resolves.toBe(
        path.join(config.root, "C:", "Windows", "system.ini")
      );
    }
  });

  it("rejects an empty path", async () => {
    const config = await makeConfig();
    await expectCode(resolveInRoot(config, "   "), "invalid_input");
  });

  it("rejects a NUL byte as invalid input rather than leaking a library error", async () => {
    // Regression: this used to reach `fs` and surface as an unclassified
    // `internal` error carrying a raw Node message, which tells a model nothing
    // it can act on.
    const config = await makeConfig();
    await expectCode(resolveInRoot(config, "a.csv\u0000.png"), "invalid_input");
    await expectCode(resolveOutputPath(config, "out\u0000.xlsx"), "invalid_input");
  });

  it("refuses a reserved Windows device name on Windows only", async () => {
    // `<root>/CON` is *inside* the root, so containment cannot catch it. On
    // Windows it opens the console — this server's own JSON-RPC transport.
    const config = await makeConfig();
    if (process.platform === "win32") {
      await expectCode(resolveInRoot(config, "CON"), "invalid_input");
      await expectCode(resolveInRoot(config, "sub/nul.txt"), "invalid_input");
      await expectCode(resolveOutputPath(config, "COM1"), "invalid_input");
    } else {
      // On POSIX these are ordinary file names and must stay usable.
      await expect(resolveInRoot(config, "CON")).resolves.toBe(path.join(config.root, "CON"));
      await expect(resolveOutputPath(config, "COM1")).resolves.toBe(
        path.join(config.outputRoot, "COM1")
      );
    }
  });

  it("reports not_found only for genuinely missing inputs", async () => {
    const config = await makeConfig();
    await expectCode(resolveInRoot(config, "missing.xlsx", { mustExist: true }), "not_found");
  });

  it("resolves the root itself", async () => {
    const config = await makeConfig();
    await expect(resolveInRoot(config, ".", { mustExist: true })).resolves.toBe(config.root);
  });
});

describe("assertWritable", () => {
  it("passes when writes are allowed", async () => {
    const config = await makeConfig();
    expect(() => assertWritable(config)).not.toThrow();
  });

  it("throws readonly when they are not", async () => {
    const config = await makeConfig({ readonly: true });
    expect(() => assertWritable(config)).toThrow(McpToolError);
  });
});

describe("isInside", () => {
  it("does not treat a sibling with a shared prefix as contained", () => {
    // The bug a `startsWith` implementation would have.
    expect(isInside("/srv/root", "/srv/root-2/file")).toBe(false);
    expect(isInside("/srv/root", "/srv/root/file")).toBe(true);
    expect(isInside("/srv/root", "/srv/root")).toBe(true);
  });
});

describe("reservedDeviceSegment", () => {
  // Tested directly because the guard using it fires on Windows only, so these
  // rules are otherwise unassertable on a POSIX host or CI runner.
  it("matches a device name however Windows would spell it", () => {
    expect(reservedDeviceSegment("NUL")).toBe("NUL");
    expect(reservedDeviceSegment("nul")).toBe("nul");
    // Everything from the first dot is an extension, so this is still the device.
    expect(reservedDeviceSegment("nul.txt")).toBe("nul.txt");
    // Trailing blanks are ignored by the Win32 path parser.
    expect(reservedDeviceSegment("NUL ")).toBe("NUL ");
    expect(reservedDeviceSegment("COM1")).toBe("COM1");
    expect(reservedDeviceSegment("LPT9.docx")).toBe("LPT9.docx");
  });

  it("matches the spellings that are easy to miss", () => {
    // A colon opens a device or an alternate data stream, and the name in front
    // of it is what Windows resolves — cutting only at "." missed both.
    expect(reservedDeviceSegment("NUL:stream")).toBe("NUL:stream");
    expect(reservedDeviceSegment("con:")).toBe("con:");
    // Windows reads the ISO 8859-1 superscripts as digits, so these are devices.
    expect(reservedDeviceSegment("COM\u00B9")).toBe("COM\u00B9");
    expect(reservedDeviceSegment("LPT\u00B2")).toBe("LPT\u00B2");
    expect(reservedDeviceSegment("com\u00B3.txt")).toBe("com\u00B3.txt");
    // The console handles: absent from Microsoft's "do not use" list, but
    // CreateFile opens them.
    expect(reservedDeviceSegment("CONIN$")).toBe("CONIN$");
    expect(reservedDeviceSegment("conout$")).toBe("conout$");
  });

  it("finds one in any segment, on either separator", () => {
    expect(reservedDeviceSegment("reports/2026/CON")).toBe("CON");
    expect(reservedDeviceSegment("reports\\AUX\\a.txt")).toBe("AUX");
    expect(reservedDeviceSegment("@output/prn.pdf")).toBe("prn.pdf");
  });

  it("leaves ordinary names alone", () => {
    // The near-misses are the point: rejecting these would be a false positive
    // on every platform, and `NULL`/`CONSOLE` are perfectly good file names.
    // `COM0`/`LPT0` are here because no such device exists.
    for (const ordinary of [
      "NULL",
      "CONSOLE",
      "console.log",
      "COM",
      "COM0",
      "COM10",
      "LPT",
      "LPT0",
      "CONIN",
      "report.xlsx",
      "a/b/c.docx",
      ".",
      "..",
      ""
    ]) {
      expect(reservedDeviceSegment(ordinary), ordinary).toBeUndefined();
    }
  });
});
