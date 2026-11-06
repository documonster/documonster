/** Refusals that changed when the stream argument checks moved to `@utils/schema`. */
import { Streams } from "@stream/index";
import { describe, expect, it } from "vitest";

function refusal(call: () => unknown): {
  name: string;
  field: string;
  message: string;
  actualType: string;
} {
  try {
    call();
  } catch (err) {
    const e = err as { name: string; field: string; message: string; actualType: string };
    return { name: e.name, field: e.field, message: e.message, actualType: e.actualType };
  }
  throw new Error("expected a refusal");
}

describe("stream argument checks declared as schemas", () => {
  it("describes a refused value rather than naming its type", () => {
    expect(refusal(() => Streams.fromString(5 as never))).toEqual({
      name: "StreamArgumentError",
      field: "str",
      message: "Streams.fromString: expected a string, got 5",
      actualType: "5"
    });
    expect(refusal(() => Streams.createReadableFromArray("ab" as never)).message).toBe(
      'Streams.createReadableFromArray: expected an array, got "ab"'
    );
  });

  it("refuses an array as an options object", () => {
    expect(refusal(() => new Streams.PullStream([] as never))).toMatchObject({
      field: "options",
      message: "Streams.PullStream: expected an options object, got an array"
    });
    expect(refusal(() => new Streams.BufferedStream({ batchSize: 0 })).message).toBe(
      "Streams.BufferedStream: expected a positive integer batchSize, got 0"
    );
  });
});
