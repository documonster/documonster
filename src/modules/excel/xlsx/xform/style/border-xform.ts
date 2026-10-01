import type { Color } from "@excel/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { ColorXform } from "@excel/xlsx/xform/style/color-xform";
import { parseBoolean } from "@utils/utils";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface EdgeModel {
  style?: string;
  color?: Partial<Color>;
}

interface BorderModel {
  top?: EdgeModel;
  left?: EdgeModel;
  bottom?: EdgeModel;
  right?: EdgeModel;
  diagonal?: EdgeModel & { up?: boolean; down?: boolean };
  vertical?: EdgeModel;
  horizontal?: EdgeModel;
  color?: Partial<Color>;
}

class EdgeXform extends BaseXform {
  declare private name: string;
  declare public map: { color: ColorXform };
  declare private defaultColor: Partial<Color> | undefined;
  declare public parser?: BaseXform;

  constructor(name: string) {
    super();

    this.name = name;
    this.map = {
      color: new ColorXform()
    };
  }

  get tag(): string {
    return this.name;
  }

  render(xmlStream: XmlSink, model?: EdgeModel, defaultColor?: Partial<Color>): void {
    const color = (model && model.color) || defaultColor || this.defaultColor;
    xmlStream.openNode(this.name);
    if (model && model.style) {
      xmlStream.addAttribute("style", model.style);
      if (color) {
        this.map.color.render(xmlStream, color);
      }
    }
    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.name: {
        const { style } = node.attributes;
        if (style) {
          this.model = {
            style
          };
        } else {
          this.model = undefined;
        }
        return true;
      }
      case "color":
        this.parser = this.map.color;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }

    if (name === this.name) {
      if (this.map.color.model) {
        if (!this.model) {
          this.model = {};
        }
        this.model.color = this.map.color.model;
      }
    }

    return false;
  }

  validStyle(value: string): boolean {
    return EdgeXform.validStyleValues[value];
  }

  static validStyleValues: { [key: string]: boolean } = [
    "thin",
    "dashed",
    "dotted",
    "dashDot",
    "hair",
    "dashDotDot",
    "slantDashDot",
    "mediumDashed",
    "mediumDashDotDot",
    "mediumDashDot",
    "medium",
    "double",
    "thick"
  ].reduce((p: { [key: string]: boolean }, v: string) => {
    p[v] = true;
    return p;
  }, {});
}

// Border encapsulates translation from border model to/from xlsx
class BorderXform extends BaseXform {
  declare public map: { [key: string]: EdgeXform };
  declare public parser?: BaseXform;
  declare private diagonalUp: boolean | undefined;
  declare private diagonalDown: boolean | undefined;
  /**
   * How deep inside an element this xform does not model the parser currently is.
   *
   * Without it, the *close* of an unknown child was taken as the end of the border: the parent stopped
   * delegating, the real `</border>` never reached this xform, and the whole border — every edge already
   * read — was dropped. `<vertical>` and `<horizontal>` did exactly that to every differential format that
   * carried them, before they were modelled; anything the schema adds later would do it again.
   */
  declare private unknownDepth: number;

  constructor() {
    super();

    this.unknownDepth = 0;
    this.map = {
      top: new EdgeXform("top"),
      left: new EdgeXform("left"),
      bottom: new EdgeXform("bottom"),
      right: new EdgeXform("right"),
      diagonal: new EdgeXform("diagonal"),
      vertical: new EdgeXform("vertical"),
      horizontal: new EdgeXform("horizontal")
    };
  }

  render(xmlStream: XmlSink, model: BorderModel): void {
    const { color } = model;
    xmlStream.openNode("border");
    if (model.diagonal && model.diagonal.style) {
      if (model.diagonal.up) {
        xmlStream.addAttribute("diagonalUp", "1");
      }
      if (model.diagonal.down) {
        xmlStream.addAttribute("diagonalDown", "1");
      }
    }
    const add = (edgeModel: EdgeModel | undefined, edgeXform: EdgeXform): void => {
      let edge = edgeModel;
      if (edge && !edge.color && model.color) {
        // don't mess with incoming models
        edge = {
          ...edge,
          color: model.color
        };
      }
      edgeXform.render(xmlStream, edge, color);
    };
    add(model.left, this.map.left);
    add(model.right, this.map.right);
    add(model.top, this.map.top);
    add(model.bottom, this.map.bottom);
    add(model.diagonal, this.map.diagonal);
    // Inner edges only when the model has them, so a cell's border is written exactly as before.
    if (model.vertical) {
      add(model.vertical, this.map.vertical);
    }
    if (model.horizontal) {
      add(model.horizontal, this.map.horizontal);
    }

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (this.unknownDepth > 0) {
      this.unknownDepth++;
      return true;
    }
    switch (node.name) {
      case "border":
        this.reset();
        this.unknownDepth = 0;
        this.diagonalUp = parseBoolean(node.attributes.diagonalUp);
        this.diagonalDown = parseBoolean(node.attributes.diagonalDown);
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }
        this.unknownDepth = 1;
        return true;
    }
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    if (this.unknownDepth > 0) {
      this.unknownDepth--;
      return true;
    }
    if (name === "border") {
      const model: BorderModel = {};
      let hasContent = false;
      const add = (
        key: "left" | "right" | "top" | "bottom" | "diagonal" | "vertical" | "horizontal",
        edgeModel: EdgeModel | undefined,
        extensions?: { up?: boolean; down?: boolean }
      ): void => {
        if (edgeModel) {
          if (extensions) {
            Object.assign(edgeModel, extensions);
          }
          model[key] = edgeModel;
          hasContent = true;
        }
      };
      add("left", this.map.left.model);
      add("right", this.map.right.model);
      add("top", this.map.top.model);
      add("bottom", this.map.bottom.model);
      add("diagonal", this.map.diagonal.model, { up: this.diagonalUp, down: this.diagonalDown });
      add("vertical", this.map.vertical.model);
      add("horizontal", this.map.horizontal.model);
      this.model = hasContent ? model : undefined;
    }
    return false;
  }
}

export { BorderXform };
