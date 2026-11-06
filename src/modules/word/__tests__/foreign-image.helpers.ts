import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";

import { Document, Io } from "../index";

/** A 1×1 PNG, so a picture typed as JPEG/PNG by its content type still decodes. */
export const PIXEL_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

/** A 1×1 baseline JPEG. */
export const PIXEL_JPEG = new Uint8Array([
  0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01,
  0x00, 0x01, 0x00, 0x00, 0xff, 0xc0, 0x00, 0x0b, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11,
  0x00, 0xff, 0xc4, 0x00, 0x14, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00,
  0x3f, 0x00, 0x7b, 0x40, 0xff, 0xd9
]);

/** Bytes standing in for a compressed EMF; nothing here decodes them. */
export const EMZ_BYTES = new Uint8Array([0x1f, 0x8b, 0x08, 0x00, 1, 2, 3, 4, 5, 6]);

export async function readParts(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const parts = new Map<string, Uint8Array>();
  for await (const entry of unzip(bytes).entries()) {
    parts.set(entry.path, await entry.bytes());
  }
  return parts;
}

/** The content type `[Content_Types].xml` gives a part, by Override, then by Default. */
export function contentTypeOf(parts: Map<string, Uint8Array>, path: string): string | undefined {
  const xml = new TextDecoder().decode(parts.get("[Content_Types].xml"));
  const override = new RegExp(
    `PartName="/${path.replace(/[.]/g, "\\.")}" ContentType="([^"]+)"`,
    "i"
  );
  const o = override.exec(xml) ?? new RegExp(`ContentType="([^"]+)" PartName="/${path}"`).exec(xml);
  if (o) {
    return o[1];
  }
  const ext = path.slice(path.lastIndexOf(".") + 1);
  const d = new RegExp(`Extension="${ext}" ContentType="([^"]+)"`, "i").exec(xml);
  return d?.[1];
}

/**
 * A package as another producer writes it: one picture stored as `.emz` and typed by a `Default`,
 * one as `.jfif` typed by an `Override`. Neither extension is an `ImageMediaType`.
 */
export async function foreignImageDocx(): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addImage(doc, PIXEL_PNG, "png", 9525 as never, 9525 as never);
  Document.addImage(doc, PIXEL_JPEG, "jpeg", 9525 as never, 9525 as never);
  const parts = await readParts(await Io.toBuffer(doc));
  const archive = zip({ level: 0 });
  for (const [path, data] of parts) {
    const text = (): string => new TextDecoder().decode(data);
    if (path === "word/media/image1.png") {
      archive.add("word/media/image1.emz", EMZ_BYTES);
    } else if (path === "word/media/image2.jpeg") {
      archive.add("word/media/photo.jfif", data);
    } else if (path === "word/_rels/document.xml.rels") {
      archive.add(
        path,
        new TextEncoder().encode(
          text()
            .replace("media/image1.png", "media/image1.emz")
            .replace("media/image2.jpeg", "media/photo.jfif")
        )
      );
    } else if (path === "[Content_Types].xml") {
      archive.add(
        path,
        new TextEncoder().encode(
          text().replace(
            "</Types>",
            '<Default Extension="emz" ContentType="image/x-emz"/><Override PartName="/word/media/photo.jfif" ContentType="image/jpeg"/></Types>'
          )
        )
      );
    } else {
      archive.add(path, data);
    }
  }
  return archive.toBuffer();
}
