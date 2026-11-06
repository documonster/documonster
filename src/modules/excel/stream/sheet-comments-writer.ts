import type { SheetRelsWriter } from "@excel/stream/sheet-rels-writer";
import type { WorkbookWriterLike } from "@excel/stream/worksheet-writer";
import { colCache } from "@excel/utils/col-cache";
import {
  commentsPath,
  commentsRelTargetFromWorksheet,
  vmlDrawingPath,
  vmlDrawingRelTargetFromWorksheet
} from "@excel/utils/ooxml-paths";
import type { StreamBuf } from "@excel/utils/stream-buf";
import { claimVmlBlock, createVmlShapeIdSpace, VML_BLOCK_SIZE } from "@excel/utils/vml-shape-ids";
import { RelType } from "@excel/xlsx/rel-type";
import { CommentXform } from "@excel/xlsx/xform/comment/comment-xform";
import type { CommentModel } from "@excel/xlsx/xform/comment/comment-xform";
import { VmlShapeXform } from "@excel/xlsx/xform/comment/vml-shape-xform";
import { XmlWriter } from "@xml/writer";

interface SheetCommentsWriterOptions {
  id: number;
  workbook: WorkbookWriterLike;
}

class SheetCommentsWriter {
  id: number;
  count: number;
  private _worksheet: { comments?: CommentModel[] };
  private _workbook: WorkbookWriterLike;
  private _sheetRelsWriter: SheetRelsWriter;
  private _commentsStream?: StreamBuf;
  private _vmlStream?: StreamBuf;
  vmlRelId?: string;
  startedData?: boolean;
  /** The next VML shape id; `0` before the part has claimed a block. */
  private _nextShapeId = 0;

  constructor(
    worksheet: { comments?: CommentModel[] },
    sheetRelsWriter: SheetRelsWriter,
    options: SheetCommentsWriterOptions
  ) {
    // in a workbook, each sheet will have a number
    this.id = options.id;
    this.count = 0;
    this._worksheet = worksheet;
    this._workbook = options.workbook;
    this._sheetRelsWriter = sheetRelsWriter;
  }

  get commentsStream(): StreamBuf {
    if (!this._commentsStream) {
      this._commentsStream = this._workbook._openStream(commentsPath(this.id));
    }
    return this._commentsStream;
  }

  get vmlStream(): StreamBuf {
    if (!this._vmlStream) {
      this._vmlStream = this._workbook._openStream(vmlDrawingPath(this.id));
    }
    return this._vmlStream;
  }

  private _addRelationships(): void {
    const commentRel = {
      Type: RelType.Comments,
      Target: commentsRelTargetFromWorksheet(this.id)
    };
    this._sheetRelsWriter.addRelationship(commentRel);

    const vmlDrawingRel = {
      Type: RelType.VmlDrawing,
      Target: vmlDrawingRelTargetFromWorksheet(this.id)
    };
    this.vmlRelId = this._sheetRelsWriter.addRelationship(vmlDrawingRel);
  }

  private _addCommentRefs(): void {
    this._workbook.commentRefs.push({
      commentName: `comments${this.id}`,
      vmlDrawing: `vmlDrawing${this.id}`
    });
  }

  /**
   * The next note's VML shape id, from a block claimed in the workbook-wide space so no two sheets
   * share one. The part declares its first block in `o:idmap`, which is written before the count is
   * known; a sheet with more than 1023 notes continues in further claimed blocks, which stay unique
   * across the workbook but are not listed in this part's `o:idmap`.
   *
   * Listing them would mean writing the part only once the count is known, at commit. That holds
   * every note's VML in memory uncompressed until then, where now it reaches the ZIP entry, and is
   * compressed, as each row commits — a streaming writer's memory would grow with its notes. The
   * `o:idmap` element cannot be moved after the shapes either: Excel writes it first and nothing
   * establishes that it accepts it anywhere else.
   */
  private _takeShapeId(): number {
    if (this._nextShapeId === 0 || this._nextShapeId % VML_BLOCK_SIZE === 0) {
      const space = (this._workbook.vmlShapeIds ??= createVmlShapeIdSpace());
      this._nextShapeId = claimVmlBlock(space) * VML_BLOCK_SIZE + 1;
    }
    return this._nextShapeId++;
  }

  private _writeOpen(): void {
    const firstShapeId = this._takeShapeId();
    this._nextShapeId = firstShapeId;
    this.commentsStream.write(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        "<authors><author>Author</author></authors>" +
        "<commentList>"
    );
    this.vmlStream.write(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<xml xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
        '<o:shapelayout v:ext="edit">' +
        `<o:idmap v:ext="edit" data="${Math.floor(firstShapeId / VML_BLOCK_SIZE)}" />` +
        "</o:shapelayout>" +
        '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
        '<v:stroke joinstyle="miter" />' +
        '<v:path gradientshapeok="t" o:connecttype="rect" />' +
        "</v:shapetype>"
    );
  }

  private _writeComment(comment: CommentModel): void {
    const commentXform = new CommentXform();
    const commentsXmlWriter = new XmlWriter();
    commentXform.render(commentsXmlWriter, comment);
    this.commentsStream.write(commentsXmlWriter.xml);

    const vmlShapeXform = new VmlShapeXform();
    const vmlXmlWriter = new XmlWriter();
    // The comment doubles as the VML shape input (it carries note geometry +
    // refAddress); VmlShapeXform.render reads that subset.
    vmlShapeXform.render(
      vmlXmlWriter,
      comment as unknown as Parameters<VmlShapeXform["render"]>[1],
      this._takeShapeId()
    );
    this.vmlStream.write(vmlXmlWriter.xml);
  }

  private _writeClose(): void {
    this.commentsStream.write("</commentList></comments>");
    this.vmlStream.write("</xml>");
  }

  addComments(comments: CommentModel[]): void {
    if (comments && comments.length) {
      if (!this.startedData) {
        this._worksheet.comments = [];
        this._writeOpen();
        this._addRelationships();
        this._addCommentRefs();
        this.startedData = true;
      }

      comments.forEach(item => {
        item.refAddress = colCache.decodeAddress(item.ref);
      });

      comments.forEach(comment => {
        this._writeComment(comment);
        this.count += 1;
      });
    }
  }

  commit(): void {
    if (this.count) {
      this._writeClose();
      this.commentsStream.end();
      this.vmlStream.end();
    }
  }
}

export { SheetCommentsWriter };
