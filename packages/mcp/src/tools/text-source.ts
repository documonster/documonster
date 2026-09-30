/**
 * Reading a tool's large content from a file instead of from its arguments.
 *
 * A tool argument cannot be longer than one model reply (see `text_write`), so every
 * parameter that can carry a lot of content — Markdown, template data, rows — has a
 * file-path twin. The rules for those twins are identical everywhere and are kept
 * here so they cannot drift: exactly one of the pair, the path confined to the
 * sandbox like any input, the file size capped like any document.
 */

import { readFile } from "node:fs/promises";

import type { ServerConfig } from "../config.js";
import { toolError } from "../errors.js";
import { resolveInRoot } from "../sandbox.js";
import { assertReadableSize } from "./fs-helpers.js";

/**
 * Require exactly one of an inline value and its file-path twin.
 *
 * @returns Which one was given.
 */
export function chooseSource(
  inline: unknown,
  from: string | undefined,
  names: readonly [inline: string, from: string]
): "inline" | "from" {
  if (inline !== undefined && from !== undefined) {
    throw toolError.invalidInput(
      `pass \`${names[0]}\` or \`${names[1]}\`, not both`,
      `Use \`${names[1]}\` for content built with text_write, \`${names[0]}\` for short content.`
    );
  }
  if (inline === undefined && from === undefined) {
    throw toolError.invalidInput(
      `\`${names[0]}\` or \`${names[1]}\` is required`,
      `Pass short content in \`${names[0]}\`; build long content with text_write and pass its path as \`${names[1]}\`.`
    );
  }
  return from === undefined ? "inline" : "from";
}

/** Read a UTF-8 text file named by a tool argument, within the sandbox and size limit. */
export async function readTextSource(config: ServerConfig, userPath: string): Promise<string> {
  const resolved = await resolveInRoot(config, userPath, { mustExist: true });
  await assertReadableSize(config, resolved, userPath);
  const text = await readFile(resolved, "utf8");
  // A BOM is an encoding marker, not content: left in, it becomes a stray
  // character at the start of a heading or makes JSON.parse fail.
  return text.startsWith("\uFEFF") ? text.slice(1) : text;
}

/** Read and parse a JSON file named by a tool argument. */
export async function readJsonSource(config: ServerConfig, userPath: string): Promise<unknown> {
  const text = await readTextSource(config, userPath);
  try {
    return JSON.parse(text) as unknown;
  } catch (cause) {
    throw toolError.invalidInput(
      `${userPath} is not valid JSON: ${cause instanceof Error ? cause.message : String(cause)}`,
      "If it was built with text_write, check that the parts join into one JSON value — read it back with doc_read.",
      { cause }
    );
  }
}
