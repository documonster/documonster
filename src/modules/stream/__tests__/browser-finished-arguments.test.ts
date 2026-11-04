/**
 * The browser `finished()` names the real problem when it is given something that is not a
 * stream, rather than reporting it as an unobservable Web stream.
 */
import { finished } from "@stream/browser/pipeline";
import { UnsupportedStreamTypeError } from "@stream/errors";
import { describe, expect, it } from "vitest";

describe("browser finished()", () => {
  it.each([null, 42, {}])("refuses %s as not a stream", async value => {
    await expect(finished(value as never)).rejects.toBeInstanceOf(UnsupportedStreamTypeError);
    expect(() => finished(value as never, () => {})).toThrow(UnsupportedStreamTypeError);
  });

  it("still refuses a Web stream with its own explanation", async () => {
    await expect(finished(new ReadableStream() as never)).rejects.toThrow(/Web stream/);
  });
});
