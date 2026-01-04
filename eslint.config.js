import eslint from "@eslint/js";
import { defineConfig } from "eslint/config";
import prettier from "eslint-config-prettier";
import oxlint from "eslint-plugin-oxlint";
import tseslint from "typescript-eslint";
import unusedImports from "eslint-plugin-unused-imports";
import { importX } from "eslint-plugin-import-x";
import tsParser from "@typescript-eslint/parser";
import { createTypeScriptImportResolver } from "eslint-import-resolver-typescript";

export const commonLintConfig = defineConfig(
  {
    plugins: {
      ["@typescript-eslint"]: tseslint.plugin,
      "unused-imports": unusedImports,
      "import-x": importX
    }
  },
  eslint.configs.recommended,
  ...tseslint.configs.recommended,
  prettier,
  ...oxlint.configs["flat/recommended"],
  {
    ignores: [
      "**/dist/**",
      "**/node_modules/**",
      "tmp/**",
      "docs/.vitepress/cache/**",
      "docs/.vitepress/dist/**",
      "**/*.graphql",
      "**/*.mustache",
      "**/*.md",
      "entities/**",
      "**/*.js",
      "**/*.cjs",
      "**/*.mjs",
      "**/gql/**",
      "**/graphql/**"
    ]
  }
);

export const getLintModuleConfiguration = ({ files, tsConfigPath, extraRules }) =>
  defineConfig({
    files,
    languageOptions: {
      parser: tsParser,
      ecmaVersion: "latest",
      sourceType: "module",
      parserOptions: {
        ecmaVersion: "latest",
        projectService: {
          allowDefaultProject: ["docs/.vitepress/config.ts", "docs/.vitepress/theme/index.ts"]
        },
        // Use a stable repo root for the TS project service.
        // Some TS files (e.g. VitePress config) are intentionally not part of the main tsconfig.
        tsconfigRootDir: tsConfigPath ?? process.cwd(),
        warnOnUnsupportedTypeScriptVersion: false
      }
    },
    settings: {
      "import-x/resolver-next": [
        createTypeScriptImportResolver({
          project: tsConfigPath
        })
      ]
    },
    rules: {
      // conflict with recommendation
      "no-useless-escape": 0,
      "no-empty": 0,
      "comma-dangle": 0,
      "consistent-return": 0,
      "no-param-reassign": 0,
      "no-useless-return": 0,
      "no-case-declarations": 0,
      "no-async-promise-executor": 0,
      "arrow-parens": ["error", "as-needed"],

      // conflict ts
      // can open but too much work, later
      "@typescript-eslint/no-unused-vars": 0,
      "@typescript-eslint/no-explicit-any": 0,
      "@typescript-eslint/no-duplicate-enum-values": 0,
      "@typescript-eslint/ban-ts-comment": 0,
      "@typescript-eslint/ban-types": 0,
      "@typescript-eslint/no-unsafe-function-type": 0,
      "@typescript-eslint/no-unused-expressions": 0,
      // Wait to enable
      "@typescript-eslint/consistent-type-imports": 2,

      // extra rules help
      "object-curly-newline": 2,
      "eol-last": 2,
      "no-return-assign": 2,
      "no-unneeded-ternary": 2,
      "spaced-comment": 2,

      // typescript rules
      "@typescript-eslint/consistent-type-exports": 2,
      // maybe 2 by default
      "@typescript-eslint/no-non-null-asserted-optional-chain": 2,

      // 强制所有相对路径的 import 必须有扩展名（.js）
      "@typescript-eslint/no-require-imports": "error",

      "import-x/no-duplicates": "error",
      "import-x/no-unresolved": [
        "error",
        {
          ignore: ["^(?!public/).+"]
        }
      ],
      "unused-imports/no-unused-imports": "error",
      "unused-imports/no-unused-vars": [
        "off",
        { vars: "all", varsIgnorePattern: "^_", args: "after-used", argsIgnorePattern: "^_" }
      ],
      ...extraRules
    }
  });

export default defineConfig(
  ...commonLintConfig,
  ...getLintModuleConfiguration({ files: ["**/*.ts", "**/*.tsx"], extraRules: {} })
);
