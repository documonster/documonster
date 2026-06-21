import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";
import { XmlWriter } from "@xml/writer";

interface StaticModel {
  tag: string;
  $?: any;
  c?: StaticModel[];
  t?: string;
}

// const model = {
//   tag: 'name',
//   $: {attr: 'value'},
//   c: [
//     { tag: 'child' }
//   ],
//   t: 'some text'
// };

function build(xmlStream: XmlSink, model: StaticModel): void {
  xmlStream.openNode(model.tag, model.$);
  if (model.c) {
    model.c.forEach(child => {
      build(xmlStream, child);
    });
  }
  if (model.t) {
    xmlStream.writeText(model.t);
  }
  xmlStream.closeNode();
}

class StaticXform extends BaseXform {
  declare private _model: StaticModel;
  declare private _xml?: string;

  constructor(model: StaticModel) {
    super();

    // This class is an optimisation for static (unimportant and unchanging) xml
    // It is stateless - apart from its static model and so can be used as a singleton
    // Being stateless - it will only track entry to and exit from it's root xml tag during parsing and nothing else
    // Known limitations:
    //    since stateless - parseOpen always returns true. Parent xform must know when to start using this xform
    //    if the root tag is recursive, the parsing will behave unpredictably
    this._model = model;
  }

  render(xmlStream: XmlSink): void {
    if (!this._xml) {
      const stream = new XmlWriter();
      build(stream, this._model);
      this._xml = stream.xml;
    }
    xmlStream.writeRaw(this._xml);
  }

  parseOpen(): boolean {
    return true;
  }

  parseText(): void {}

  parseClose(name: string): boolean {
    switch (name) {
      case this._model.tag:
        return false;
      default:
        return true;
    }
  }
}

export { StaticXform };
