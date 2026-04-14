/**
 * Shared helpers for parse-sax tests.
 */

import { SaxParser } from "@xml/sax";
import { it, expect } from "vitest";

export type EventTuple = ["opentag" | "closetag" | "text" | "error", any];

export function test(options: {
  name: string;
  xml: string | string[];
  expect: EventTuple[];
  opt?: { xmlns?: boolean; position?: boolean; fileName?: string; fragment?: boolean };
}): void {
  it(options.name, () => {
    const parser = new SaxParser(options.opt);
    const events: EventTuple[] = [];

    parser.on("opentag", tag => {
      events.push([
        "opentag",
        {
          name: tag.name,
          attributes: tag.attributes,
          isSelfClosing: tag.isSelfClosing
        }
      ]);
    });

    parser.on("closetag", tag => {
      events.push([
        "closetag",
        {
          name: tag.name,
          attributes: tag.attributes,
          isSelfClosing: tag.isSelfClosing
        }
      ]);
    });

    parser.on("text", text => {
      events.push(["text", text]);
    });

    parser.on("error", err => {
      events.push(["error", err.message]);
    });

    if (Array.isArray(options.xml)) {
      for (const chunk of options.xml) {
        parser.write(chunk);
      }
      parser.close();
    } else {
      parser.write(options.xml).close();
    }

    expect(events).toEqual(options.expect);
  });
}
