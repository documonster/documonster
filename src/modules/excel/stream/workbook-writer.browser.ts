/**
 * WorkbookWriter - Browser Streaming Excel Writer
 *
 * This module contains the full cross-platform implementation for the streaming
 * workbook writer and a browser-compatible `WorkbookWriter` class.
 *
 * Node.js uses `workbook-writer.ts`, which extends the same base implementation
 * with filesystem-specific features (filename output + image loading).
 */

import { Zip, ZipDeflate } from "@archive/zip/stream";
import { StreamBuf } from "@excel/utils/stream-buf";
import { base64ToUint8Array } from "@utils/utils";
import { ExcelNotSupportedError, ImageError } from "@excel/errors";
import { RelType } from "@excel/xlsx/rel-type";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { SharedStrings } from "@excel/utils/shared-strings";
import { DefinedNames } from "@excel/defined-names";
import { CoreXform } from "@excel/xlsx/xform/core/core-xform";
import { RelationshipsXform } from "@excel/xlsx/xform/core/relationships-xform";
import { ContentTypesXform } from "@excel/xlsx/xform/core/content-types-xform";
import { AppXform } from "@excel/xlsx/xform/core/app-xform";
import { WorkbookXform } from "@excel/xlsx/xform/book/workbook-xform";
import { SharedStringsXform } from "@excel/xlsx/xform/strings/shared-strings-xform";
import { FeaturePropertyBagXform } from "@excel/xlsx/xform/core/feature-property-bag-xform";
import { DrawingXform } from "@excel/xlsx/xform/drawing/drawing-xform";
import { theme1Xml } from "@excel/xlsx/xml/theme1";
import type { Writable } from "@stream";
import { toWritable } from "@stream";
import { stringToUint8Array } from "@utils/binary";
import {
  drawingPath,
  drawingRelsPath,
  mediaPath,
  OOXML_PATHS,
  OOXML_REL_TARGETS,
  worksheetRelTarget
} from "@excel/utils/ooxml-paths";
import { filterDrawingAnchors } from "@excel/utils/drawing-utils";
import type { ImageData, WorkbookView, AddWorksheetOptions } from "@excel/types";
import { WorksheetWriter } from "@excel/stream/worksheet-writer";

const EMPTY_U8 = new Uint8Array(0);
const TEXT_DECODER = new TextDecoder();

// ============================================================================
// Types
// ============================================================================

interface Medium extends ImageData {
  type: "image";
  name: string;
}

interface CommentRef {
  commentName: string;
  vmlDrawing: string;
}

export interface ZlibOptions {
  flush?: number;
  finishFlush?: number;
  chunkSize?: number;
  windowBits?: number;
  level?: number;
  memLevel?: number;
  strategy?: number;
  dictionary?: Uint8Array | ArrayBuffer;
}

export interface WorkbookZipOptions {
  comment?: string;
  forceLocalTime?: boolean;
  forceZip64?: boolean;
  store?: boolean;
  zlib?: Partial<ZlibOptions>;
  compressionOptions?: { level?: number };
}

export interface WorkbookWriterOptions {
  created?: Date;
  modified?: Date;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  useSharedStrings?: boolean;
  useStyles?: boolean;
  zip?: Partial<WorkbookZipOptions>;
  stream?: Writable | WritableStream<Uint8Array>;
  filename?: string; // Node.js only
  trueStreaming?: boolean;
}

interface OutputStreamLike {
  emit(eventName: string | symbol, ...args: any[]): boolean;
  write(chunk: any): boolean | Promise<boolean>;
  end(): void;
  once(eventName: string | symbol, listener: (...args: any[]) => void): this;
  removeListener(eventName: string | symbol, listener: (...args: any[]) => void): this;
}

// ============================================================================
// WorksheetWriter interface (to avoid circular dependency)
// ============================================================================

export interface WorksheetWriterLike {
  id: number;
  name: string;
  rId?: string;
  committed?: boolean;
  stream: any;
  commit(): void;
  /** Drawing model — populated after commit if images were added */
  drawing?: { rId: string; name: string; anchors: any[]; rels: any[] };
}

export interface WorksheetWriterConstructor<T extends WorksheetWriterLike> {
  new (options: {
    id: number;
    name: string;
    workbook: any;
    useSharedStrings: boolean;
    properties?: any;
    state?: any;
    pageSetup?: any;
    views?: any;
    autoFilter?: any;
    headerFooter?: any;
  }): T;
}

// ============================================================================
// Base Class
// ============================================================================

export abstract class WorkbookWriterBase<TWorksheetWriter extends WorksheetWriterLike> {
  created: Date;
  modified: Date;
  creator: string;
  lastModifiedBy: string;
  lastPrinted?: Date;
  useSharedStrings: boolean;
  sharedStrings: SharedStrings;
  styles: StylesXform;
  private _definedNames: DefinedNames;
  private _worksheets: TWorksheetWriter[];
  views: WorkbookView[];
  zipOptions?: Partial<WorkbookZipOptions>;
  compressionLevel: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
  media: Medium[];
  commentRefs: CommentRef[];
  zip: Zip;
  stream: OutputStreamLike;
  promise: Promise<void[] | void>;
  protected _trueStreaming: boolean;
  protected WorksheetWriterClass: WorksheetWriterConstructor<TWorksheetWriter>;

  constructor(
    options: WorkbookWriterOptions,
    WorksheetWriterClass: WorksheetWriterConstructor<TWorksheetWriter>
  ) {
    this.WorksheetWriterClass = WorksheetWriterClass;
    this.created = options.created || new Date();
    this.modified = options.modified || this.created;
    this.creator = options.creator ?? "ExcelTS";
    this.lastModifiedBy = options.lastModifiedBy ?? "ExcelTS";
    this.lastPrinted = options.lastPrinted;

    this.useSharedStrings = options.useSharedStrings ?? false;
    this.sharedStrings = new SharedStrings();
    this.styles = options.useStyles ? new StylesXform(true) : new (StylesXform as any).Mock(true);
    this._definedNames = new DefinedNames();
    this._worksheets = [];
    this.views = [];

    this.zipOptions = options.zip;
    const level = options.zip?.zlib?.level ?? options.zip?.compressionOptions?.level ?? 6;
    this.compressionLevel = Math.max(0, Math.min(9, level)) as
      | 0
      | 1
      | 2
      | 3
      | 4
      | 5
      | 6
      | 7
      | 8
      | 9;

    this.media = [];
    this.commentRefs = [];
    this._trueStreaming = options.trueStreaming ?? false;

    // Create Zip instance
    this.zip = new Zip((err, data, final) => {
      if (err) {
        this.stream.emit("error", err);
      } else {
        // `streaming-zip` already emits `Uint8Array`; avoid copying per chunk.
        this.stream.write(data);
        if (final) {
          this.stream.end();
        }
      }
    });

    // Setup output stream
    this.stream = this._createOutputStream(options);

    // Theme and office rels are deferred to commit() so that worksheet files
    // are added to the ZIP first. This ensures StreamingZip sets ondata on
    // the worksheet immediately, allowing pushSync to flow data through
    // without accumulating in _dataQueue.
    this.promise = Promise.resolve();
  }

  /**
   * Create output stream - can be overridden by Node.js to support filename
   */
  protected _createOutputStream(options: WorkbookWriterOptions): OutputStreamLike {
    if (options.stream) {
      return toWritable(options.stream);
    }
    return new StreamBuf();
  }

  get definedNames(): DefinedNames {
    return this._definedNames;
  }

  /** @internal */
  _openStream(path: string): InstanceType<typeof StreamBuf> {
    const stream = new StreamBuf({
      bufSize: this._trueStreaming ? 4096 : 65536,
      batch: !this._trueStreaming
    });

    const zipFile = new ZipDeflate(path, { level: this.compressionLevel });
    this.zip.add(zipFile);

    const onData = (chunk: Uint8Array) => zipFile.push(chunk);
    stream.on("data", onData);

    stream.once("finish", () => {
      stream.removeListener("data", onData);
      zipFile.push(EMPTY_U8, true);
      stream.emit("zipped");
    });

    return stream;
  }

  protected _addFile(data: string | Uint8Array, name: string, base64?: boolean): void {
    const zipFile = new ZipDeflate(name, { level: this.compressionLevel });
    this.zip.add(zipFile);

    let buffer: Uint8Array;
    if (base64) {
      const base64Data = typeof data === "string" ? data : TEXT_DECODER.decode(data);
      buffer = base64ToUint8Array(base64Data);
    } else if (typeof data === "string") {
      buffer = stringToUint8Array(data);
    } else {
      buffer = data;
    }

    zipFile.push(buffer, true);
  }

  private _commitWorksheets(): Promise<void> {
    const commitWorksheet = (worksheet: TWorksheetWriter): Promise<void> => {
      if (!worksheet.committed) {
        return new Promise(resolve => {
          worksheet.stream.once("zipped", () => resolve());
          worksheet.commit();
        });
      }
      return Promise.resolve();
    };
    const promises = this._worksheets.map(commitWorksheet);
    return promises.length ? Promise.all(promises).then(() => {}) : Promise.resolve();
  }

  async commit(): Promise<void> {
    await this.promise;
    await this._commitWorksheets();
    await this.addMedia();
    this.addDrawings();
    await Promise.all([
      this.addThemes(),
      this.addOfficeRels(),
      this.addContentTypes(),
      this.addApp(),
      this.addCore(),
      this.addSharedStrings(),
      this.addStyles(),
      this.addFeaturePropertyBag(),
      this.addWorkbookRels()
    ]);
    await this.addWorkbook();
    await this._finalize();
  }

  get nextId(): number {
    for (let i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i]) {
        return i;
      }
    }
    return this._worksheets.length || 1;
  }

  addImage(image: ImageData): number {
    const id = this.media.length;
    const medium: Medium = {
      ...image,
      type: "image" as const,
      name: `image${id}.${image.extension}`
    };
    this.media.push(medium);
    return id;
  }

  getImage(id: number): ImageData | undefined {
    return this.media[id];
  }

  addWorksheet(name?: string, options?: Partial<AddWorksheetOptions>): TWorksheetWriter {
    const opts = options || {};
    const useSharedStrings =
      opts.useSharedStrings !== undefined ? opts.useSharedStrings : this.useSharedStrings;

    if ((opts as any).tabColor) {
      console.trace("tabColor option has moved to { properties: tabColor: {...} }");
      opts.properties = { tabColor: (opts as any).tabColor, ...opts.properties };
    }

    const id = this.nextId;
    name = name ?? `sheet${id}`;

    const worksheet = new this.WorksheetWriterClass({
      id,
      name,
      workbook: this,
      useSharedStrings,
      properties: opts.properties,
      state: opts.state,
      pageSetup: opts.pageSetup,
      views: opts.views,
      autoFilter: opts.autoFilter,
      headerFooter: opts.headerFooter
    });

    this._worksheets[id] = worksheet;
    return worksheet;
  }

  getWorksheet(id?: string | number): TWorksheetWriter | undefined {
    if (id === undefined) {
      return this._worksheets.find(() => true);
    }
    if (typeof id === "number") {
      return this._worksheets[id];
    }
    if (typeof id === "string") {
      const idLower = id.toLowerCase();
      return this._worksheets.find(ws => ws?.name?.toLowerCase() === idLower);
    }
    return undefined;
  }

  addStyles(): Promise<void> {
    return new Promise(resolve => {
      this._addFile(this.styles.xml, OOXML_PATHS.xlStyles);
      resolve();
    });
  }

  addThemes(): Promise<void> {
    return new Promise(resolve => {
      this._addFile(theme1Xml, OOXML_PATHS.xlTheme1);
      resolve();
    });
  }

  addOfficeRels(): Promise<void> {
    return new Promise(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml([
        { Id: "rId1", Type: RelType.OfficeDocument, Target: OOXML_PATHS.xlWorkbook },
        { Id: "rId2", Type: RelType.CoreProperties, Target: OOXML_PATHS.docPropsCore },
        { Id: "rId3", Type: RelType.ExtenderProperties, Target: OOXML_PATHS.docPropsApp }
      ]);
      this._addFile(xml, OOXML_PATHS.rootRels);
      resolve();
    });
  }

  addContentTypes(): Promise<void> {
    return new Promise(resolve => {
      const worksheets = this._worksheets.filter(Boolean);
      // In the streaming path, ZIP entries use ws.id which is always sequential.
      // Set fileIndex = id to satisfy the ContentTypesXform contract.
      worksheets.forEach((ws: any) => {
        ws.fileIndex = ws.id;
      });

      // Collect drawing models from worksheets that have images
      const drawings = worksheets.filter(ws => ws.drawing).map(ws => ws.drawing);

      const model = {
        worksheets,
        sharedStrings: this.sharedStrings,
        commentRefs: this.commentRefs,
        media: this.media,
        drawings,
        hasCheckboxes: this.styles.hasCheckboxes
      };
      const xform = new ContentTypesXform();
      this._addFile(xform.toXml(model), OOXML_PATHS.contentTypes);
      resolve();
    });
  }

  /**
   * Add media files - can be overridden by Node.js for file system support
   */
  addMedia(): Promise<void[]> {
    return Promise.all(
      this.media.map(async medium => {
        if (medium.type === "image") {
          const filename = mediaPath(medium.name);
          if (medium.buffer) {
            this._addFile(medium.buffer, filename);
            return;
          }
          if (medium.base64) {
            const content = medium.base64.substring(medium.base64.indexOf(",") + 1);
            this._addFile(content, filename, true);
            return;
          }
          if (medium.filename) {
            throw new ExcelNotSupportedError(
              "Loading images from filename",
              "not supported in browser. Use buffer or base64."
            );
          }
        }
        throw new ImageError("Unsupported media");
      })
    );
  }

  /**
   * Generate drawing XML and drawing relationship files for worksheets that have images.
   * Must be called after _commitWorksheets() so that each WorksheetWriter has built its
   * drawing model, and after addMedia() so that media files are already in the ZIP.
   */
  protected addDrawings(): void {
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();

    for (const ws of this._worksheets) {
      if (!ws?.drawing) {
        continue;
      }

      const { drawing } = ws;

      // Filter out invalid anchors using shared utility
      const filteredAnchors = filterDrawingAnchors(drawing.anchors);
      const drawingForWrite = { ...drawing, anchors: filteredAnchors };

      // Prepare and generate drawing XML
      drawingXform.prepare(drawingForWrite);
      const xml = drawingXform.toXml(drawingForWrite);
      this._addFile(xml, drawingPath(drawing.name));

      // Generate drawing relationships
      const relsXml = relsXform.toXml(drawing.rels);
      this._addFile(relsXml, drawingRelsPath(drawing.name));
    }
  }

  addApp(): Promise<void> {
    return new Promise(resolve => {
      const xform = new AppXform();
      this._addFile(
        xform.toXml({ worksheets: this._worksheets.filter(Boolean) }),
        OOXML_PATHS.docPropsApp
      );
      resolve();
    });
  }

  addCore(): Promise<void> {
    return new Promise(resolve => {
      const xform = new CoreXform();
      this._addFile(xform.toXml(this), OOXML_PATHS.docPropsCore);
      resolve();
    });
  }

  addSharedStrings(): Promise<void> {
    if (this.sharedStrings.count) {
      return new Promise(resolve => {
        const xform = new SharedStringsXform();
        this._addFile(xform.toXml(this.sharedStrings), OOXML_PATHS.xlSharedStrings);
        resolve();
      });
    }
    return Promise.resolve();
  }

  addFeaturePropertyBag(): Promise<void> {
    if (this.styles.hasCheckboxes) {
      const xform = new FeaturePropertyBagXform();
      this._addFile(xform.toXml({}), OOXML_PATHS.xlFeaturePropertyBag);
    }
    return Promise.resolve();
  }

  addWorkbookRels(): Promise<void> {
    let count = 1;
    const relationships: Array<{ Id: string; Type: string; Target: string }> = [
      { Id: `rId${count++}`, Type: RelType.Styles, Target: OOXML_REL_TARGETS.workbookStyles },
      { Id: `rId${count++}`, Type: RelType.Theme, Target: OOXML_REL_TARGETS.workbookTheme1 }
    ];
    if (this.sharedStrings.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: RelType.SharedStrings,
        Target: OOXML_REL_TARGETS.workbookSharedStrings
      });
    }
    // Add FeaturePropertyBag relationship if checkboxes are used
    if (this.styles.hasCheckboxes) {
      relationships.push({
        Id: `rId${count++}`,
        Type: RelType.FeaturePropertyBag,
        Target: OOXML_REL_TARGETS.workbookFeaturePropertyBag
      });
    }
    this._worksheets.forEach(ws => {
      if (ws) {
        ws.rId = `rId${count++}`;
        relationships.push({
          Id: ws.rId,
          Type: RelType.Worksheet,
          Target: worksheetRelTarget(ws.id)
        });
      }
    });

    return new Promise(resolve => {
      const xform = new RelationshipsXform();
      this._addFile(xform.toXml(relationships), OOXML_PATHS.xlWorkbookRels);
      resolve();
    });
  }

  addWorkbook(): Promise<void> {
    const model = {
      worksheets: this._worksheets.filter(Boolean),
      definedNames: this._definedNames.model,
      views: this.views,
      properties: {},
      calcProperties: {}
    };
    return new Promise(resolve => {
      const xform = new WorkbookXform();
      xform.prepare(model);
      this._addFile(xform.toXml(model), OOXML_PATHS.xlWorkbook);
      resolve();
    });
  }

  private _finalize(): Promise<this> {
    return new Promise((resolve, reject) => {
      const onError = (err: Error) => {
        this.stream.removeListener("finish", onFinish);
        reject(err);
      };
      const onFinish = () => {
        this.stream.removeListener("error", onError);
        resolve(this);
      };
      this.stream.once("error", onError);
      this.stream.once("finish", onFinish);
      this.zip.end();
    });
  }
}

export const WorkbookWriterOptionsSchema = {
  useSharedStrings: ["boolean"],
  useStyles: ["boolean"],
  trueStreaming: ["boolean"]
} as const;

// ============================================================================
// Browser-compatible WorkbookWriter
// ============================================================================

class WorkbookWriter extends WorkbookWriterBase<WorksheetWriter> {
  constructor(options: WorkbookWriterOptions = {}) {
    super(options, WorksheetWriter);
  }
}

export { WorkbookWriter };
