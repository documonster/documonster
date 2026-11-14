/**
 * Making objects disposable with `using` and `await using`.
 *
 * `await using` disposes an object by calling its `[Symbol.asyncDispose]()`. A runtime without the symbol — a browser
 * from before it shipped it — still runs code that a compiler lowered: esbuild and Babel then look the method up under
 * `Symbol.for("Symbol.asyncDispose")`, the key `core-js-pure` gives the symbol it polyfills. Keyed on that symbol where
 * the runtime has none, the method serves both; a computed method name `[Symbol.asyncDispose]` would instead name the
 * method `"undefined"` there.
 */

/**
 * The symbol `SymbolConstructor` declares as `name`, in the types the code is compiled against — `never` where they
 * declare none. A consumer compiling for a browser with `lib: ["ES2022", "DOM"]` and no `@types/node` has no
 * `Symbol.asyncDispose`, and a declaration naming it outright is an error in every such project (TS2550).
 */
type DeclaredSymbol<K extends PropertyKey> = SymbolConstructor extends {
  readonly [P in K]: infer S extends symbol;
}
  ? S
  : never;

/**
 * `[Symbol.asyncDispose]()`, for a class to merge into its type (`interface ZipReader extends AsyncDisposeMember {}`)
 * where the consumer's types declare the symbol, and nothing where they do not. A class member declared
 * `[Symbol.asyncDispose]` could not be written without naming the symbol.
 */
export type AsyncDisposeMember = { [P in DeclaredSymbol<"asyncDispose">]: () => Promise<void> };

/** `[Symbol.dispose]()`, as {@link AsyncDisposeMember}. */
export type DisposeMember = { [P in DeclaredSymbol<"dispose">]: () => void };

/** The key `await using` disposes an object by, given the runtime's `Symbol` — see the module comment. */
export function asyncDisposeKey(symbol: {
  readonly asyncDispose?: symbol;
}): DeclaredSymbol<"asyncDispose"> {
  return (symbol.asyncDispose ??
    Symbol.for("Symbol.asyncDispose")) as DeclaredSymbol<"asyncDispose">;
}

/** {@link asyncDisposeKey} on this runtime. */
export const ASYNC_DISPOSE: DeclaredSymbol<"asyncDispose"> = asyncDisposeKey(Symbol);

/**
 * Give `prototype` its `[Symbol.asyncDispose]`, under {@link ASYNC_DISPOSE}. The class takes the member's type from
 * {@link AsyncDisposeMember}: a method written in the class body would be keyed on the runtime's symbol, or on
 * `"undefined"` where there is none.
 */
export function defineAsyncDispose<T extends object>(
  prototype: T,
  dispose: (this: T) => Promise<void>
): void {
  Object.defineProperty(prototype, ASYNC_DISPOSE, {
    configurable: true,
    writable: true,
    value: dispose
  });
}

/**
 * The key `using` disposes an object by, given the runtime's `Symbol` — `Symbol.dispose`, as
 * {@link asyncDisposeKey}.
 */
export function disposeKey(symbol: { readonly dispose?: symbol }): DeclaredSymbol<"dispose"> {
  return (symbol.dispose ?? Symbol.for("Symbol.dispose")) as DeclaredSymbol<"dispose">;
}

/** {@link disposeKey} on this runtime. */
export const DISPOSE: DeclaredSymbol<"dispose"> = disposeKey(Symbol);

/** Give `prototype` its `[Symbol.dispose]`, under {@link DISPOSE} — see {@link defineAsyncDispose}. */
export function defineDispose<T extends object>(prototype: T, dispose: (this: T) => void): void {
  Object.defineProperty(prototype, DISPOSE, {
    configurable: true,
    writable: true,
    value: dispose
  });
}
