import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { CommentXform } from "@excel/xlsx/xform/comment/comment-xform";
import type { CommentModel } from "@excel/xlsx/xform/comment/comment-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

interface CommentsModel {
  comments: CommentModel[];
}

const DEFAULT_AUTHOR = "Author";

class CommentsXform extends BaseXform<CommentsModel> {
  declare public map: { [key: string]: CommentXform };
  declare public parser?: BaseXform;

  /** Authors collected while parsing the <authors> element. */
  private _authors: string[] = [];
  /** Whether we are currently inside the <authors> element. */
  private _inAuthors = false;
  /** Whether we are currently inside an <author> element (collecting text). */
  private _inAuthor = false;
  /** Accumulator for the current <author> text content. */
  private _currentAuthor = "";

  constructor() {
    super();
    this.map = {
      comment: new CommentXform()
    };
    this.model = { comments: [] };
  }

  render(xmlStream: XmlSink, model?: CommentsModel): void {
    const renderModel = model || this.model;
    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode("comments", CommentsXform.COMMENTS_ATTRIBUTES);

    // Collect unique authors from comments
    const authorSet = new Set<string>();
    for (const comment of renderModel!.comments) {
      authorSet.add(comment.author ?? DEFAULT_AUTHOR);
    }
    const authors = [...authorSet];

    xmlStream.openNode("authors");
    for (const author of authors) {
      xmlStream.leafNode("author", undefined, author);
    }
    xmlStream.closeNode();

    // comments
    xmlStream.openNode("commentList");
    renderModel!.comments.forEach(comment => {
      // Set the authorId based on the authors list for rendering
      const authorId = authors.indexOf(comment.author ?? DEFAULT_AUTHOR);
      this.map.comment.render(xmlStream, { ...comment, authorId: authorId >= 0 ? authorId : 0 });
    });
    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case "authors":
        this._inAuthors = true;
        this._authors = [];
        return true;
      case "author":
        if (this._inAuthors) {
          this._inAuthor = true;
          this._currentAuthor = "";
        }
        return true;
      case "commentList":
        this.model = {
          comments: []
        };
        return true;
      case "comment":
        this.parser = this.map.comment;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    if (this._inAuthor) {
      this._currentAuthor += text;
    } else if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    switch (name) {
      case "authors":
        this._inAuthors = false;
        return true;
      case "author":
        if (this._inAuthors) {
          this._authors.push(this._currentAuthor);
          this._inAuthor = false;
          this._currentAuthor = "";
        }
        return true;
      case "commentList":
        // Resolve authorId → author name on each comment
        for (const comment of this.model!.comments) {
          const { authorId } = comment;
          if (authorId != null && authorId >= 0 && authorId < this._authors.length) {
            comment.author = this._authors[authorId];
          }
        }
        return false;
      case "comment":
        this.model!.comments.push(this.parser!.model as CommentModel);
        this.parser = undefined;
        return true;
      default:
        if (this.parser) {
          this.parser.parseClose(name);
        }
        return true;
    }
  }

  static COMMENTS_ATTRIBUTES = {
    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  };
}

export { CommentsXform };
