import { BaseXform } from "@excel/xlsx/xform/base-xform";

abstract class BaseCellAnchorXform extends BaseXform {
  declare public map: { [key: string]: any };
  declare public parser: any;
  declare public model: any;

  abstract get tag(): string;

  parseOpen(node: any): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          range: {
            editAs: node.attributes.editAs
          }
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  reconcilePicture(model: any, options: any): any {
    if (model && model.rId) {
      const rel = options.rels[model.rId];
      if (!rel) {
        return undefined;
      }
      const match = rel.Target.match(/.*\/media\/(.+[.][a-zA-Z]{3,4})/);
      if (match) {
        const name = match[1];
        const mediaId = options.mediaIndex[name];
        const medium = options.media[mediaId];
        // Preserve alphaModFix (transparency) from the picture model if present
        if (medium && model.alphaModFix !== undefined) {
          return { ...medium, alphaModFix: model.alphaModFix };
        }
        return medium;
      }
    }
    return undefined;
  }
}

export { BaseCellAnchorXform };
