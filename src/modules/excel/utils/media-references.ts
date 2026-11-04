/**
 * Which workbook media a written package actually references.
 *
 * Removing an image's last placement (`Image.remove`) leaves the medium registered — its id is a
 * position in `workbook.media` and must stay valid until the next write, so a later `Image.place`
 * with the same id still works. The writers therefore decide at *write* time which of the media an
 * API call released to emit.
 *
 * The decision is read from the relationship parts the writer produces, not recomputed from the
 * model, and that is deliberate. In an OPC package one part reaches another **only** through a
 * relationship, so the written `.rels` parts are by definition the complete list of references —
 * drawings, chart and chartEx picture fills, chart user shapes, VML header/footer pictures, sheet
 * backgrounds, cell images, and every relationship part preserved verbatim from a loaded file. A
 * model-side enumeration would have to repeat each of those writers' relationship logic, and the
 * day one was missed it would drop a medium that is still in use; reading what was written cannot
 * miss one.
 *
 * Each `Target` is resolved against the part that owns the relationship part and compared as a
 * full part name, case-insensitively as OPC part names are. External targets (`TargetMode="External"`)
 * are links outside the package and reference no medium.
 */

const RELATIONSHIP = /<(?:\w+:)?Relationship\b[^>]*>/g;
const TARGET = /\bTarget\s*=\s*(?:"([^"]*)"|'([^']*)')/;
const EXTERNAL = /\bTargetMode\s*=\s*["']External["']/;
const RELS_PART = /^(.*?)_rels\/[^/]*\.rels$/i;
const textDecoder = new TextDecoder("utf-8");

/** The directory a relationship part's targets are relative to: that of its source part. */
function sourceDirectory(relsPath: string): string | undefined {
  const match = RELS_PART.exec(relsPath.replace(/^\//, ""));
  return match ? match[1] : undefined;
}

function decode(target: string): string {
  try {
    return decodeURIComponent(target);
  } catch {
    return target; // a malformed escape is compared as written
  }
}

/** `target` resolved against `base` (a directory ending `/`, or `""` for the package root). */
function resolvePartName(base: string, target: string): string {
  const path = decode(target).replace(/\\/g, "/");
  const segments = path.startsWith("/") ? [] : base.split("/").filter(Boolean);
  for (const segment of path.split("/")) {
    if (segment === "..") {
      segments.pop();
    } else if (segment !== "." && segment !== "") {
      segments.push(segment);
    }
  }
  return segments.join("/").toLowerCase();
}

/** Collects the package-internal relationship targets of a package as its parts are written. */
export class MediaReferenceTracker {
  private readonly targets = new Set<string>();

  /** Whether a part of this name carries references — i.e. is a relationship part. */
  wants(path: string): boolean {
    return sourceDirectory(path) !== undefined;
  }

  /** Record a part. Only relationship parts are read; nothing else can reference a part. */
  observe(path: string, data: string | Uint8Array): void {
    const base = sourceDirectory(path);
    if (base === undefined) {
      return;
    }
    const text = typeof data === "string" ? data : textDecoder.decode(data);
    for (const [element] of text.matchAll(RELATIONSHIP)) {
      const target = TARGET.exec(element);
      if (target && !EXTERNAL.test(element)) {
        this.targets.add(resolvePartName(base, target[1] ?? target[2] ?? ""));
      }
    }
  }

  /** {@link observe} for a part written in chunks. */
  observeChunks(path: string, chunks: ReadonlyArray<string | Uint8Array>): void {
    const decoder = new TextDecoder("utf-8");
    let text = "";
    for (const chunk of chunks) {
      text += typeof chunk === "string" ? chunk : decoder.decode(chunk, { stream: true });
    }
    this.observe(path, text + decoder.decode());
  }

  /** Whether the part named `partName` (e.g. `xl/media/image1.png`) is the target of a relationship. */
  isReferenced(partName: string): boolean {
    return this.targets.has(partName.replace(/^\//, "").toLowerCase());
  }
}
