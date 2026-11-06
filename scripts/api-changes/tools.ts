/**
 * The MCP server's tool protocol: each tool's name, group, title, annotations, description and input schema (as JSON
 * Schema). Used by `scripts/api-changes.ts` to compare two releases of `@documonster/mcp`.
 */
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

export interface ToolInfo {
  group: unknown;
  title: string;
  description: string;
  annotations: unknown;
  schema: unknown;
}

export interface ToolCatalog {
  groups: readonly string[];
  tools: Record<string, ToolInfo>;
}

/** Load the built package at `dir` (its `dist/index.js`) and describe its tools, with the zod it resolves. */
export async function describeTools(dir: string): Promise<ToolCatalog> {
  const req = createRequire(path.join(dir, "package.json"));
  const mod = await import(pathToFileURL(path.join(dir, "dist/index.js")).href);
  const { z } = await import(pathToFileURL(req.resolve("zod")).href);
  const tools: Record<string, ToolInfo> = {};
  for (const t of mod.ALL_TOOLS) {
    tools[t.name] = {
      group: t.group,
      title: t.title,
      description: t.description,
      annotations: t.annotations,
      schema: z.toJSONSchema(z.object(t.inputSchema), {
        unrepresentable: "any",
        io: "input"
      })
    };
  }
  return { groups: mod.TOOL_GROUPS, tools };
}
