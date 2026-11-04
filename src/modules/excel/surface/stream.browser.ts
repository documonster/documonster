/**
 * `Stream` namespace surface — browser entry.
 *
 * Same streaming surface as the Node `surface/stream.ts`, but the
 * writer/reader classes resolve to their browser variants (Web Streams,
 * no Node file-path sinks).
 */
export { WorkbookWriter } from "@excel/stream/workbook-writer.browser";
export { WorkbookReader } from "@excel/stream/workbook-reader.browser";

export * from "@excel/surface/stream-handles";

/** The worksheet writer handed out by `WorkbookWriter.addWorksheet`. */
export type { WorksheetWriter } from "@excel/stream/worksheet-writer";
/** The worksheet reader yielded by `WorkbookReader.read`. */
export type { WorksheetReader } from "@excel/stream/worksheet-reader";
/** The hyperlink reader yielded by `WorkbookReader.read`. */
export type { HyperlinkReader } from "@excel/stream/hyperlink-reader";

export type {
  WorkbookWriterOptions,
  WorkbookZipOptions,
  ZlibOptions,
  Medium
} from "@excel/stream/workbook-writer.browser";
export type { WorkbookReaderInput } from "@excel/stream/workbook-reader.browser";
export type {
  WorkbookReaderOptions,
  ParseEvent,
  ParseEventType,
  SharedStringEvent,
  SharedStringValue,
  WorksheetReadyEvent,
  HyperlinksEvent
} from "@excel/stream/workbook-reader.browser";
export type {
  WorksheetReaderOptions,
  WorksheetEvent,
  WorksheetEventType,
  WorksheetHyperlink,
  RowEvent,
  HyperlinkEvent
} from "@excel/stream/worksheet-reader";
export type { Hyperlink, HyperlinkReaderOptions } from "@excel/stream/hyperlink-reader";
/** An image placed on a streamed sheet, as `WorksheetWriter#getImages` lists it. */
export type { WriterImageModel } from "@excel/stream/worksheet-writer";
/** What the deprecated `WorkbookWriter#packageSink` returns. */
export type { PackageSink } from "@excel/utils/package-sink";
/** A listener of the streaming readers' events (`on`, `once`, `addListener`, …). */
export type { EventListener } from "@utils/event-emitter";
