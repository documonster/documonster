import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { markdownToDocx } from "../convert/markdown/markdown-import";
import {
  Build,
  Convert,
  Document,
  DocxArgumentError,
  DocxInvalidImageError,
  DocxWriteError,
  Io,
  Layout,
  Package,
  Stream,
  Template,
  type DocxDocument,
  type ImageDef,
  type ImageMediaType
} from "../index";
import {
  contentTypeOf,
  EMZ_BYTES,
  foreignImageDocx,
  PIXEL_JPEG,
  readParts
} from "./foreign-image.helpers";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

async function partNames(bytes: Uint8Array): Promise<string[]> {
  const names: string[] = [];
  for await (const entry of unzip(bytes).entries()) {
    names.push(entry.path);
  }
  return names;
}

describe("image media type at the API boundary", () => {
  it("Document.addImage rejects a MIME type and names the accepted values", () => {
    const doc = Document.create();
    const call = (): unknown =>
      Document.addImage(doc, PNG, "image/png" as ImageMediaType, 100 as never, 100 as never);
    expect(call).toThrow(DocxInvalidImageError);
    expect(call).toThrow(/png, jpeg, gif/);
    expect(doc.images ?? []).toHaveLength(0);
  });

  it("Document.addFloatingImage rejects an unknown media type", () => {
    const doc = Document.create();
    expect(() =>
      Document.addFloatingImage(doc, PNG, "jpg" as ImageMediaType, 100 as never, 100 as never)
    ).toThrow(DocxInvalidImageError);
  });

  it("every ImageMediaType is accepted and written with a typed part", async () => {
    const doc = Document.create();
    Document.addImage(doc, PNG, "png", 100 as never, 100 as never);
    Document.addImage(doc, PNG, "jpeg", 100 as never, 100 as never);
    const names = await partNames(await Io.toBuffer(doc));
    expect(names).toContain("word/media/image1.png");
    expect(names).toContain("word/media/image2.jpeg");
  });
});

describe("image parts at write time", () => {
  const withImage = (fileName: string, mediaType: string): DocxDocument => ({
    body: [],
    images: [{ data: PNG, fileName, mediaType: mediaType as ImageMediaType, rId: "rId9" }]
  });

  it("Io.toBuffer refuses an image whose name has no typed extension", async () => {
    await expect(Io.toBuffer(withImage("image1.image/png", "image/png"))).rejects.toThrow(
      DocxInvalidImageError
    );
    await expect(Io.toBuffer(withImage("picture", "png"))).rejects.toThrow(DocxInvalidImageError);
    await expect(Io.toBuffer(withImage("picture.xyz", "png"))).rejects.toThrow(
      DocxInvalidImageError
    );
  });

  it("is an argument error naming the field at fault", async () => {
    const doc = Document.create();
    let error: unknown;
    try {
      Document.addImage(doc, PNG, "image/png" as ImageMediaType, 100 as never, 100 as never);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(DocxArgumentError);
    expect((error as DocxArgumentError).field).toBe("mediaType");
    const at = async (d: DocxDocument): Promise<string> =>
      ((await Io.toBuffer(d).catch((e: unknown) => e)) as DocxArgumentError).field;
    expect(await at(withImage("a.png", "image/png"))).toBe("images[0].mediaType");
    expect(await at(withImage("picture", "png"))).toBe("images[0].fileName");
    const thumb: DocxDocument = {
      body: [],
      thumbnail: { data: PNG, contentType: "image/bmp" as never }
    };
    const thumbError = await Io.toBuffer(thumb).catch((e: unknown) => e);
    expect(thumbError).toBeInstanceOf(DocxInvalidImageError);
    expect((thumbError as DocxArgumentError).field).toBe("thumbnail.contentType");
  });

  it("Io.toBuffer refuses an image with an invalid mediaType", async () => {
    await expect(Io.toBuffer(withImage("a.png", "image/png"))).rejects.toThrow(
      DocxInvalidImageError
    );
  });

  it("a read document's jpg / tif media types still write", async () => {
    await expect(Io.toBuffer(withImage("a.jpg", "jpg"))).resolves.toBeInstanceOf(Uint8Array);
    await expect(Io.toBuffer(withImage("a.tif", "tif"))).resolves.toBeInstanceOf(Uint8Array);
  });
});

describe("image media type at the other entry points", () => {
  const badImage = (): ImageDef => ({
    data: PNG,
    fileName: "logo.png",
    mediaType: "image/png" as ImageMediaType
  });

  it("Markdown resolveImage", async () => {
    await expect(
      markdownToDocx("![a](a.png)", {
        resolveImage: () => ({ data: PNG, mediaType: "image/png" as ImageMediaType })
      })
    ).rejects.toThrow(DocxInvalidImageError);
  });

  it("Template image placeholders", () => {
    const doc: DocxDocument = { body: [Build.textParagraph("{{%logo}}")] };
    expect(() =>
      Template.fill(
        doc,
        { logo: { image: badImage(), width: 100, height: 100 } },
        { richPlaceholders: true }
      )
    ).toThrow(DocxInvalidImageError);
  });

  it("Package.patch image patches", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "{{x}}");
    const bytes = await Io.toBuffer(doc);
    await expect(
      Package.patch(bytes, [
        { placeholder: "{{x}}", content: { type: "image", image: badImage(), width: 1, height: 1 } }
      ])
    ).rejects.toThrow(DocxInvalidImageError);
  });

  it("the streaming writer", () => {
    expect(
      () =>
        new Stream.DocumentWriter({
          images: [{ data: PNG, fileName: "image1.image/png", mediaType: "png", rId: "rId5" }]
        })
    ).toThrow(DocxInvalidImageError);
  });

  it("a thumbnail with an unknown content type", async () => {
    const doc: DocxDocument = {
      body: [],
      thumbnail: { data: PNG, contentType: "png" as "image/png" }
    };
    await expect(Io.toBuffer(doc)).rejects.toThrow(DocxInvalidImageError);
  });
});

describe("readers give an ImageMediaType", () => {
  it("a read .jpg picture is jpeg", async () => {
    const doc = Document.create();
    Document.addImage(doc, PNG, "png", 100 as never, 100 as never);
    const bytes = await Io.toBuffer(doc);
    const entries = new Map<string, Uint8Array>();
    for await (const e of unzip(bytes).entries()) {
      entries.set(e.path, await e.bytes());
    }
    const archive = zip({ level: 0 });
    for (const [path, data] of entries) {
      const text = (): string => new TextDecoder().decode(data);
      if (path === "word/media/image1.png") {
        archive.add("word/media/image1.jpg", data);
      } else if (path.endsWith(".rels") || path === "[Content_Types].xml") {
        archive.add(
          path,
          new TextEncoder().encode(
            text()
              .replace("media/image1.png", "media/image1.jpg")
              .replace(
                'Extension="png" ContentType="image/png"',
                'Extension="jpg" ContentType="image/jpeg"'
              )
          )
        );
      } else {
        archive.add(path, data);
      }
    }
    const read = await Io.read(await archive.toBuffer());
    expect(read.images?.[0]?.mediaType).toBe("jpeg");
  });

  it("ODT pictures carry ImageMediaType, not MIME types", async () => {
    const doc = Document.create();
    Document.addImage(doc, PNG, "png", 100 as never, 100 as never);
    const read = await Convert.fromOdt(await Convert.toOdt(doc));
    expect(read.images?.map(i => i.mediaType)).toEqual(["png"]);
    const names = await partNames(await Io.toBuffer(read));
    expect(names.some(n => /^word\/media\/[^/]+\.png$/.test(n))).toBe(true);
  });
});

describe("Package.edit", () => {
  it("refuses a new part [Content_Types].xml cannot type", async () => {
    const bytes = await Io.toBuffer(Document.create());
    await expect(
      Package.edit(bytes, [{ type: "replacePart", path: "word/media/png", data: PNG }])
    ).rejects.toThrow(DocxWriteError);
    await expect(
      Package.edit(bytes, [{ type: "replacePart", path: "word/media/a.xyz", data: PNG }])
    ).rejects.toThrow(DocxWriteError);
  });

  it("accepts a new part typed by a Default or Override added in the same edit", async () => {
    const bytes = await Io.toBuffer(Document.create());
    const ct = new TextDecoder().decode((await Package.readPart(bytes, "[Content_Types].xml"))!);
    const out = await Package.edit(bytes, [
      { type: "replacePart", path: "word/media/a.png", data: PNG },
      { type: "replacePart", path: "word/media/b", data: PNG },
      {
        type: "replacePartText",
        path: "[Content_Types].xml",
        text: ct.replace(
          "</Types>",
          '<Default Extension="png" ContentType="image/png"/><Override PartName="/word/media/b" ContentType="image/png"/></Types>'
        )
      }
    ]);
    expect(await partNames(out)).toContain("word/media/b");
  });

  it("does not judge parts the input already had", async () => {
    const bytes = await Io.toBuffer(Document.create());
    await expect(
      Package.edit(bytes, [{ type: "replacePartText", path: "word/document.xml", text: "<x/>" }])
    ).resolves.toBeInstanceOf(Uint8Array);
  });
});

describe("a read picture whose extension is not an ImageMediaType", () => {
  it("round-trips with its bytes, name and source content type", async () => {
    const first = await Io.read(await foreignImageDocx());
    const written = await readParts(await Io.toBuffer(first));
    expect(written.get("word/media/image1.emz")).toEqual(EMZ_BYTES);
    expect(written.get("word/media/photo.jfif")).toEqual(PIXEL_JPEG);
    expect(contentTypeOf(written, "word/media/image1.emz")).toBe("image/x-emz");
    expect(contentTypeOf(written, "word/media/photo.jfif")).toBe("image/jpeg");

    const again = await Io.read(await Io.toBuffer(first));
    expect(again.images?.map(i => i.fileName).sort()).toEqual(["image1.emz", "photo.jfif"]);
    const twice = await readParts(await Io.toBuffer(again));
    expect(contentTypeOf(twice, "word/media/image1.emz")).toBe("image/x-emz");
    expect(contentTypeOf(twice, "word/media/photo.jfif")).toBe("image/jpeg");
  });

  it("lays out: the JPEG by its content type, the EMZ as an opaque picture", async () => {
    const doc = await Io.read(await foreignImageDocx());
    const layout = Layout.compute(doc);
    const mimes = JSON.stringify(layout).match(/"mimeType":"[^"]+"/g) ?? [];
    expect(mimes).toContain('"mimeType":"image/jpeg"');
  });
});

describe("Package.edit types every part it adds", () => {
  it("gives a new non-media part without a type the generic application/octet-stream", async () => {
    const bytes = await Io.toBuffer(Document.create());
    const out = await readParts(
      await Package.edit(bytes, [
        { type: "replacePartText", path: "extra.txt", text: "a" },
        { type: "replacePartText", path: "notes/second.txt", text: "b" },
        { type: "replacePart", path: "custom/blob", data: PNG }
      ])
    );
    expect(contentTypeOf(out, "extra.txt")).toBe("application/octet-stream");
    expect(contentTypeOf(out, "notes/second.txt")).toBe("application/octet-stream");
    expect(contentTypeOf(out, "custom/blob")).toBe("application/octet-stream");
    const ct = new TextDecoder().decode(out.get("[Content_Types].xml"));
    expect(ct.match(/Extension="txt"/g)).toHaveLength(1);
  });

  it("leaves a part the edit typed itself alone", async () => {
    const bytes = await Io.toBuffer(Document.create());
    const ct = new TextDecoder().decode((await Package.readPart(bytes, "[Content_Types].xml"))!);
    const out = await readParts(
      await Package.edit(bytes, [
        { type: "replacePartText", path: "a.txt", text: "a" },
        {
          type: "replacePartText",
          path: "[Content_Types].xml",
          text: ct.replace(
            "</Types>",
            '<Default Extension="txt" ContentType="text/plain"/></Types>'
          )
        }
      ])
    );
    expect(contentTypeOf(out, "a.txt")).toBe("text/plain");
  });
});

describe("the writers type every caller-named part", () => {
  const untyped = (): Partial<DocxDocument> => ({
    opaqueParts: [
      { path: "word/custom/data.foo", data: PNG },
      { path: "word/custom/noext", data: PNG }
    ],
    embeddedFonts: [{ rId: "rIdF1", fileName: "face.fntdata", data: PNG }]
  });

  it("Io.toBuffer gives an untyped part application/octet-stream", async () => {
    const parts = await readParts(await Io.toBuffer({ body: [], ...untyped() }));
    expect(contentTypeOf(parts, "word/custom/data.foo")).toBe("application/octet-stream");
    expect(contentTypeOf(parts, "word/custom/noext")).toBe("application/octet-stream");
    expect(contentTypeOf(parts, "word/fonts/face.fntdata")).toBe("application/octet-stream");
  });

  it("the streaming writer does the same", async () => {
    const stream = new Stream.DocumentWriter(untyped());
    const parts = await readParts(await stream.finalize());
    expect(contentTypeOf(parts, "word/custom/data.foo")).toBe("application/octet-stream");
    expect(contentTypeOf(parts, "word/custom/noext")).toBe("application/octet-stream");
    expect(contentTypeOf(parts, "word/fonts/face.fntdata")).toBe("application/octet-stream");
  });
});
