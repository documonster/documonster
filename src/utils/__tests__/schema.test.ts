import { check, checkAt, expectedOf, fail, s } from "@utils/schema";
import type { Schema, SchemaCheckOptions } from "@utils/schema";
import { describe, expect, it } from "vitest";

class TestArgumentError extends Error {
  constructor(
    readonly field: string,
    message: string
  ) {
    super(message);
  }
}

const at = (field = "options", label?: string): SchemaCheckOptions => ({
  where: "Test.call",
  field,
  label,
  error: (f, m) => new TestArgumentError(f, m)
});

/** The `[field, message]` of the error `check` throws, or `undefined` when it accepts. */
function refusal(
  schema: Schema<unknown, boolean>,
  value: unknown,
  options = at()
): [string, string] | undefined {
  try {
    check(schema, value, options);
    return undefined;
  } catch (error) {
    expect(error).toBeInstanceOf(TestArgumentError);
    return [(error as TestArgumentError).field, (error as Error).message];
  }
}

describe("schema primitives", () => {
  it("string", () => {
    expect(refusal(s.string(), "")).toBeUndefined();
    expect(refusal(s.string(), 1)).toEqual([
      "options",
      "Test.call: options must be a string, got 1"
    ]);
    expect(refusal(s.string({ nonEmpty: true }), "")?.[1]).toBe(
      'Test.call: options must be a non-empty string, got ""'
    );
    expect(refusal(s.string({ expected: "a name" }), null)?.[1]).toBe(
      "Test.call: options must be a name, got null"
    );
  });

  it("number", () => {
    expect(refusal(s.number(), NaN)).toBeUndefined();
    expect(refusal(s.number(), "1")?.[1]).toBe('Test.call: options must be a number, got "1"');
    expect(refusal(s.number({ integer: true, min: 0 }), 1.5)?.[1]).toBe(
      "Test.call: options must be an integer >= 0, got 1.5"
    );
    expect(refusal(s.number({ integer: true, min: 0 }), NaN)).toBeDefined();
    expect(refusal(s.number({ min: 0, max: 1 }), NaN)).toBeDefined();
    expect(refusal(s.number({ min: 0, max: 1 }), 2)?.[1]).toBe(
      "Test.call: options must be a number from 0 to 1, got 2"
    );
    expect(refusal(s.number({ max: 1 }), 2)?.[1]).toContain("a number <= 1");
    expect(refusal(s.number({ finite: true }), Infinity)?.[1]).toBe(
      "Test.call: options must be a finite number, got Infinity"
    );
  });

  it("boolean, function, literal, oneOf", () => {
    expect(refusal(s.boolean(), 0)?.[1]).toBe("Test.call: options must be a boolean, got 0");
    expect(refusal(s.function(), {})?.[1]).toBe(
      "Test.call: options must be a function, got an object"
    );
    expect(refusal(s.function(), () => 1)).toBeUndefined();
    expect(refusal(s.literal("a", 1), 1)).toBeUndefined();
    expect(refusal(s.literal("a", 1), "b")?.[1]).toBe(
      'Test.call: options must be one of "a", 1, got "b"'
    );
    expect(refusal(s.oneOf([".", ","], { expected: '"." or ","' }), ";")?.[1]).toBe(
      'Test.call: options must be "." or ",", got ";"'
    );
  });

  it("bytes: any view or buffer, recognised by tag", () => {
    expect(refusal(s.bytes(), new Uint8Array(1))).toBeUndefined();
    expect(refusal(s.bytes(), new DataView(new ArrayBuffer(1)))).toBeUndefined();
    expect(refusal(s.bytes(), new ArrayBuffer(1))).toBeUndefined();
    expect(refusal(s.bytes(), [1])?.[1]).toBe(
      "Test.call: options must be an ArrayBuffer or ArrayBufferView, got an array"
    );
    // A buffer from another realm has the same tag but fails `instanceof`.
    const foreign = Object.create(null, { [Symbol.toStringTag]: { value: "ArrayBuffer" } });
    expect(refusal(s.bytes(), foreign)).toBeUndefined();
  });

  it("custom, with the shown value chosen", () => {
    const even = s.custom<number>(v => typeof v === "number" && v % 2 === 0, "even");
    expect(refusal(even, 2)).toBeUndefined();
    expect(refusal(even, 3)?.[1]).toBe("Test.call: options must be even, got 3");
    const typed = s.custom(() => false, "a known type", {
      got: v => (v as { type: unknown }).type
    });
    expect(refusal(typed, { type: "x" })?.[1]).toBe(
      'Test.call: options must be a known type, got "x"'
    );
  });
});

describe("schema modifiers", () => {
  it("optional, orUndefined and nullable", () => {
    expect(refusal(s.string(), undefined)).toBeDefined();
    expect(refusal(s.string().optional(), undefined)).toBeUndefined();
    expect(refusal(s.string().orUndefined(), undefined)).toBeUndefined();
    expect(refusal(s.string(), null)).toBeDefined();
    expect(refusal(s.string().nullable(), null)).toBeUndefined();
  });

  it("refine runs after the base check, and not for undefined", () => {
    const short = s
      .string()
      .refine(v => v.length < 3, "shorter than 3")
      .optional();
    expect(refusal(short, 5)?.[1]).toBe("Test.call: options must be a string, got 5");
    expect(refusal(short, "long")?.[1]).toBe(
      'Test.call: options must be shorter than 3, got "long"'
    );
    expect(refusal(short, undefined)).toBeUndefined();
  });
});

describe("schema containers and paths", () => {
  const shape = s.object(
    {
      name: s.string(),
      "odd key": s.number().optional(),
      columns: s
        .array(s.object({ width: s.number({ integer: true, min: 0 }) }, { expected: "a column" }))
        .optional(),
      headers: s.record(s.boolean()).optional()
    },
    { expected: "an options object" }
  );

  it("reports the object itself", () => {
    expect(refusal(shape, null)).toEqual([
      "options",
      "Test.call: options must be an options object, got null"
    ]);
    expect(refusal(shape, [])?.[1]).toContain("got an array");
  });

  it("writes keys, indices and map entries the way the library does", () => {
    expect(refusal(shape, { name: 1 })).toEqual([
      "options.name",
      "Test.call: options.name must be a string, got 1"
    ]);
    expect(refusal(shape, { name: "", "odd key": "x" })?.[0]).toBe('options["odd key"]');
    expect(
      refusal(shape, { name: "", columns: [{ width: 1 }, { width: 1 }, { width: -1 }] })
    ).toEqual([
      "options.columns[2].width",
      "Test.call: options.columns[2].width must be an integer >= 0, got -1"
    ]);
    expect(refusal(shape, { name: "", columns: [1] })?.[0]).toBe("options.columns[0]");
    expect(refusal(shape, { name: "", headers: { a: true, "b c": 1 } })).toEqual([
      'options.headers["b c"]',
      'Test.call: options.headers["b c"] must be a boolean, got 1'
    ]);
  });

  it("names an option without its parameter when label is empty", () => {
    expect(refusal(shape, { name: 1 }, at("options", ""))).toEqual([
      "options.name",
      "Test.call: name must be a string, got 1"
    ]);
    expect(refusal(shape, 1, at("options", ""))?.[1]).toBe(
      "Test.call: options must be an options object, got 1"
    );
    expect(refusal(s.string(), 1, at("input"))?.[1]).toBe(
      "Test.call: input must be a string, got 1"
    );
    expect(refusal(s.string(), 1, at("options.x", "x"))).toEqual([
      "options.x",
      "Test.call: x must be a string, got 1"
    ]);
  });

  it("allows unknown keys unless told to reject them", () => {
    expect(refusal(s.object({ a: s.string().optional() }), { b: 1 })).toBeUndefined();
    expect(
      refusal(
        s.object({ a: s.string().optional(), b: s.string().optional() }, { unknownKeys: "reject" }),
        {
          c: 1
        }
      )
    ).toEqual(['options["c"]', 'Test.call: options["c"] must be a key in a, b, got "c"']);
  });

  it("does not read inherited keys of a record", () => {
    const proto = { inherited: 1 };
    expect(refusal(s.record(s.boolean()), Object.create(proto))).toBeUndefined();
  });

  it("array minLength", () => {
    expect(refusal(s.array(s.string(), { minLength: 1 }), [])?.[1]).toBe(
      "Test.call: options must be an array, got an array"
    );
    expect(
      refusal(s.array(s.string(), { minLength: 1, expected: "a non-empty array" }), [])?.[1]
    ).toBe("Test.call: options must be a non-empty array, got an array");
  });
});

describe("schema unions", () => {
  const dynamic = s.union(
    [
      s.boolean(),
      s.record(s.union([s.boolean(), s.function()], { expected: "true, false or a function" }))
    ],
    { expected: "true, false or a map" }
  );

  it("reports a value no member claims against the union", () => {
    expect(refusal(dynamic, 1)?.[1]).toBe("Test.call: options must be true, false or a map, got 1");
    expect(refusal(dynamic, [])?.[1]).toContain("got an array");
  });

  it("reports a value against the member whose kind it is", () => {
    expect(refusal(dynamic, { a: 1 })).toEqual([
      'options["a"]',
      'Test.call: options["a"] must be true, false or a function, got 1'
    ]);
    expect(refusal(s.union([s.number({ min: 0 }), s.string()]), -1)?.[1]).toBe(
      "Test.call: options must be a number >= 0, got -1"
    );
  });

  it("joins the members' texts by default and accepts any passing member", () => {
    expect(refusal(s.union([s.string(), s.boolean()]), 1)?.[1]).toBe(
      "Test.call: options must be a string or a boolean, got 1"
    );
    expect(refusal(s.union([s.literal("a"), s.custom(v => v === 2, "two")]), 2)).toBeUndefined();
  });

  it("reports a member that failed inside the value when none claims it", () => {
    const column = s.union([
      s.string(),
      s.all(
        s.custom<unknown>(v => typeof v === "object" && v !== null && "key" in v, "a column"),
        s.object({ key: s.string(), header: s.string().optional() })
      )
    ]);
    expect(refusal(column, { key: "a", header: 1 })?.[0]).toBe("options.header");
    expect(refusal(column, {})?.[1]).toContain("must be a string or a column");
  });
});

describe("fail", () => {
  it("builds the same error for a rule a schema cannot express", () => {
    const error = fail(at("options", ""), ".lineEnding", "different from the delimiter", ",");
    expect(error).toBeInstanceOf(TestArgumentError);
    expect((error as TestArgumentError).field).toBe("options.lineEnding");
    expect(error.message).toBe(
      'Test.call: lineEnding must be different from the delimiter, got ","'
    );
  });
});

describe("schema builders for platform values", () => {
  it("abortSignal, by shape", () => {
    expect(refusal(s.abortSignal(), new AbortController().signal)).toBeUndefined();
    expect(
      refusal(s.abortSignal(), { aborted: false, addEventListener: () => {} })
    ).toBeUndefined();
    expect(refusal(s.abortSignal(), {})?.[1]).toBe(
      "Test.call: options must be an AbortSignal, got an object"
    );
  });

  it("thenable", () => {
    expect(refusal(s.thenable(), Promise.resolve(1))).toBeUndefined();
    expect(refusal(s.thenable(), { then: () => {} })).toBeUndefined();
    expect(refusal(s.thenable(), 1)?.[1]).toBe(
      "Test.call: options must be a promise or thenable, got 1"
    );
  });

  it("iterable: sync or async, and strings only on request", () => {
    async function* gen(): AsyncGenerator<number> {}
    expect(refusal(s.iterable(), [1])).toBeUndefined();
    expect(refusal(s.iterable(), gen())).toBeUndefined();
    expect(refusal(s.iterable(), new Set())).toBeUndefined();
    expect(refusal(s.iterable(), 1)?.[1]).toContain("must be an iterable or async iterable");
    expect(refusal(s.iterable(), "ab")).toBeDefined();
    expect(refusal(s.iterable({ strings: true }), "ab")).toBeUndefined();
  });

  it("builtin, by tag", () => {
    expect(refusal(s.builtin("URL"), new URL("https://example.com"))).toBeUndefined();
    expect(refusal(s.builtin("URL"), "https://example.com")?.[1]).toContain("must be a URL");
  });

  it("variant: dispatches on a tag and reports inside the chosen shape", () => {
    const op = s.variant(
      "type",
      { a: s.object({ x: s.string() }), b: s.object({ y: s.number() }) },
      { expected: "an op" }
    );
    expect(refusal(op, { type: "a", x: "1" })).toBeUndefined();
    expect(refusal(op, { type: "b", y: "1" })?.[0]).toBe("options.y");
    expect(refusal(op, { type: "c" })?.[1]).toBe("Test.call: options must be an op, got an object");
    expect(refusal(s.variant("type", { a: s.object({}) }), 1)?.[1]).toBe(
      'Test.call: options must be an object whose type is one of "a", got 1'
    );
  });
});

describe("error factory refusal details", () => {
  it("hands the parts of the message to the factory", () => {
    let seen: unknown;
    try {
      check(s.number(), "1", {
        where: "Test.call",
        field: "options.n",
        label: "n",
        error: (_field, _message, refusal) => {
          seen = refusal;
          return new Error("x");
        }
      });
    } catch {
      // expected
    }
    expect(seen).toEqual({ where: "Test.call", label: "n", expected: "a number", got: '"1"' });
  });
});

describe("schema: records, vocabularies, trees and argument lists", () => {
  it('object { record: "argument" } refuses what isRecordArgument refuses', () => {
    const shape = s.object({ a: s.number().optional() }, { record: "argument" });
    expect(refusal(shape, {})).toBeUndefined();
    for (const value of [new Date(), new Uint8Array(1), new Map(), [], null]) {
      expect(refusal(shape, value)?.[1]).toMatch(/^Test\.call: options must be an object, got /);
    }
    expect(refusal(s.object({}), new Date())).toBeUndefined();
    expect(refusal(s.record(s.string(), { record: "argument" }), new Map())).toBeDefined();
    expect(refusal(s.record(s.string(), { record: "argument" }), { a: 1 })?.[0]).toBe(
      'options["a"]'
    );
  });

  it("member takes a Set or an array and claims every string", () => {
    const fromSet = s.member<"a" | "b">(new Set(["a", "b"]));
    expect(refusal(fromSet, "a")).toBeUndefined();
    expect(refusal(fromSet, "c")?.[1]).toBe('Test.call: options must be one of "a", "b", got "c"');
    expect(refusal(s.member(["x"], { expected: "an x" }), 1)?.[1]).toBe(
      "Test.call: options must be an x, got 1"
    );
    const union = s.union([s.member(["x"]), s.number()], { expected: "x or a number" });
    expect(refusal(union, "y")?.[1]).toBe('Test.call: options must be one of "x", got "y"');
    expect(refusal(union, true)?.[1]).toBe("Test.call: options must be x or a number, got true");
  });

  it("variant { unknownTag } reports a missing or unknown tag at the tag", () => {
    const node = s.variant("kind", { a: s.object({}) }, { unknownTag: 'one of "a"' });
    expect(refusal(node, { kind: "b" })).toEqual([
      "options.kind",
      'Test.call: options.kind must be one of "a", got "b"'
    ]);
    expect(refusal(node, 1)?.[1]).toBe("Test.call: options must be an object, got 1");
  });

  it("checkAt reports as check does", () => {
    let caught: unknown;
    try {
      checkAt(s.object({ a: s.string() }), { a: 1 }, "Test.call", "options", (f, m) =>
        Object.assign(new Error(m), { field: f })
      );
    } catch (error) {
      caught = error;
    }
    expect(caught).toMatchObject({
      field: "options.a",
      message: "Test.call: options.a must be a string, got 1"
    });
  });

  it("an optional key and a required one, tested inline by the object", () => {
    const shape = s.object({ a: s.number().optional(), b: s.number() });
    expect(refusal(shape, { b: 1 })).toBeUndefined();
    expect(refusal(shape, { a: undefined, b: 1 })).toBeUndefined();
    expect(refusal(shape, {})?.[1]).toBe("Test.call: options.b must be a number, got undefined");
    const refined = s.object({
      a: s
        .number()
        .optional()
        .refine(v => v > 0, "positive")
    });
    expect(refusal(refined, {})).toBeUndefined();
    expect(refusal(refined, { a: -1 })?.[1]).toBe("Test.call: options.a must be positive, got -1");
  });

  it("is and test agree: a value passes exactly when no problem is found", () => {
    const schemas: Schema<unknown, boolean>[] = [
      s.string(),
      s.string({ nonEmpty: true }),
      s.number(),
      s.number({ integer: true, min: 0, max: 9 }),
      s.number({ finite: true, min: 0 }),
      s.number({ min: 0, max: 1 }),
      s.boolean(),
      s.function(),
      s.oneOf(["a", 1]),
      s.member(["a"]),
      s.array(s.number()),
      s.object({ a: s.number().optional(), b: s.oneOf(["x"]).optional() }),
      s.object({ a: s.string() }, { unknownKeys: "reject" }),
      s.object({}, { record: "argument" }),
      s.record(s.boolean()),
      s.union([s.string(), s.array(s.number())]),
      s.variant("t", { a: s.object({ n: s.number() }) }),
      s.custom(v => v === 1, "one"),
      s.abortSignal(),
      s.thenable(),
      s.iterable(),
      s.builtin("URL"),
      s.bytes(),
      s.string().nullable(),
      s
        .number()
        .refine(v => v > 1, "> 1")
        .optional()
    ];
    const values = [
      undefined,
      null,
      0,
      -1,
      1,
      2,
      1.5,
      NaN,
      Infinity,
      "",
      "a",
      "x",
      true,
      [],
      [1],
      ["a"],
      {},
      { a: 1 },
      { a: "a" },
      { b: "x" },
      { t: "a", n: 1 },
      { t: "a" },
      { c: 1 },
      () => 1,
      new Date(),
      new Date(NaN),
      new Uint8Array(1),
      new ArrayBuffer(1),
      new Map(),
      Promise.resolve(),
      new AbortController().signal
    ];
    for (const schema of schemas) {
      for (const value of values) {
        expect(schema.is(value)).toBe(schema.test(value) === undefined);
      }
    }
  });
});

describe("schema: unions of overlapping members, wording and extension", () => {
  it("accepts a value a narrow member takes even when a wider one claims it", () => {
    const limit = s.union([s.literal(Infinity), s.number({ integer: true, min: 1 })], {
      expected: "an integer >= 1, or Infinity"
    });
    expect(refusal(limit, Infinity)).toBeUndefined();
    expect(refusal(limit, 3)).toBeUndefined();
    expect(refusal(limit, 1.5)?.[1]).toBe("Test.call: options must be an integer >= 1, got 1.5");
    expect(refusal(limit, "x")?.[1]).toBe(
      'Test.call: options must be an integer >= 1, or Infinity, got "x"'
    );
  });

  it("expectedOf gives the text a schema refuses with at its root", () => {
    expect(expectedOf(s.number({ integer: true, min: 0, max: 9 }))).toBe("an integer from 0 to 9");
    expect(expectedOf(s.boolean().optional())).toBe("a boolean");
    expect(expectedOf(s.object({}, { expected: "an options object" }))).toBe("an options object");
    expect(expectedOf(s.union([s.string(), s.boolean()]))).toBe("a string or a boolean");
    expect(expectedOf(s.custom(() => true, "a widget"))).toBe("a widget");
    expect(expectedOf(s.string().refine(() => true, "short"))).toBe("a string");
    expect(expectedOf(s.unknown())).toBe("anything");
  });

  it("extend adds keys to an object schema and keeps its options", () => {
    const base = s.object({ a: s.number().optional() }, { expected: "an options object" });
    const extended = s.extend(base, { b: s.string().optional() });
    expect(refusal(extended, { a: 1, b: "x" })).toBeUndefined();
    expect(refusal(extended, { b: 1 })?.[0]).toBe("options.b");
    expect(refusal(extended, { a: "1" })?.[0]).toBe("options.a");
    expect(refusal(extended, 1)?.[1]).toBe("Test.call: options must be an options object, got 1");
    expect(() => s.extend(s.string() as never, {})).toThrow(TypeError);
  });
});
