import { readdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(process.cwd());
const docsRoot = path.join(repoRoot, "docs");

function normalizeNewlines(text) {
  return text.replace(/\r\n/g, "\n");
}

function rewriteRepoRelativeLinks(markdown) {
  // When the root README is copied under docs/reference/, relative links break.
  // Rewrite common repo-local targets to either site routes or GitHub URLs.
  const githubBase = "https://github.com/cjnoname/excelts/blob/main/";

  return (
    markdown
      // README.md -> docs route
      .replace(/\]\((\.?\/)?README\.md(#[^)]+)?\)/g, "](/reference/readme$2)")
      // README_zh.md -> docs route
      .replace(/\]\((\.?\/)?README_zh\.md(#[^)]+)?\)/g, "](/reference/readme-zh$2)")
      // CHANGELOG.md -> docs route
      .replace(/\]\((\.?\/)?CHANGELOG\.md(#[^)]+)?\)/g, "](/reference/changelog$2)")
      // src/... -> GitHub blob link
      .replace(/\]\((\.?\/)?src\/([^#)]+)(#[^)]+)?\)/g, (_m, _prefix, filePath, hash) => {
        const clean = `src/${filePath}`.replace(/\s/g, "%20");
        return `](${githubBase}${clean}${hash ?? ""})`;
      })
  );
}

function wrapAsVitePressPage(title, body) {
  const cleanBody = body.trimEnd();
  return `# ${title}\n\n${cleanBody}\n`;
}

async function copyMarkdownToDocs({ src, dest, title }) {
  const srcPath = path.join(repoRoot, src);
  const destPath = path.join(docsRoot, dest);

  let content = normalizeNewlines(await readFile(srcPath, "utf8"));
  if (src === "README.md" || src === "README_zh.md") {
    content = rewriteRepoRelativeLinks(content);
  }
  const page = wrapAsVitePressPage(title, content);

  await writeFile(destPath, page, "utf8");
}

async function listExampleFiles() {
  const srcRoot = path.join(repoRoot, "src");

  // Walk only a couple of levels to keep it simple and fast.
  // We specifically look for any directory named "examples" under src/**.
  const results = [];

  async function walk(dir) {
    const entries = await readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        if (entry.name === "examples") {
          const exampleEntries = await readdir(full, { withFileTypes: true });
          for (const e of exampleEntries) {
            if (!e.isFile()) continue;
            if (!/\.(ts|js|mjs|cjs|html)$/.test(e.name)) continue;
            const rel = path.relative(repoRoot, path.join(full, e.name)).replaceAll(path.sep, "/");
            results.push(rel);
          }
        } else {
          await walk(full);
        }
      }
    }
  }

  await walk(srcRoot);
  results.sort((a, b) => a.localeCompare(b));
  return results;
}

async function generateExamplesPage() {
  const files = await listExampleFiles();
  const githubTreeBase = "https://github.com/cjnoname/excelts/blob/main/";

  const lines = [
    "# Examples",
    "",
    "This page is generated at build time from `src/**/examples`.",
    ""
  ];

  if (files.length === 0) {
    lines.push("No examples found yet.");
  } else {
    for (const rel of files) {
      lines.push(`- [${rel}](${githubTreeBase}${rel})`);
    }
  }

  lines.push("");
  lines.push("Tip: you can open these files directly in the repo.");
  lines.push("");

  await writeFile(path.join(docsRoot, "examples.md"), lines.join("\n"), "utf8");
}

async function main() {
  await copyMarkdownToDocs({
    src: "README.md",
    dest: "reference/readme.md",
    title: "README"
  });

  await copyMarkdownToDocs({
    src: "README_zh.md",
    dest: "reference/readme-zh.md",
    title: "README (中文)"
  });

  await copyMarkdownToDocs({
    src: "CHANGELOG.md",
    dest: "reference/changelog.md",
    title: "Changelog"
  });

  await generateExamplesPage();
}

await main();
