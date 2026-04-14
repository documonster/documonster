import eslint from "@eslint/js";
import { defineConfig } from "eslint/config";
import prettier from "eslint-config-prettier";
import oxlint from "eslint-plugin-oxlint";
import unusedImports from "eslint-plugin-unused-imports";
import { importX } from "eslint-plugin-import-x";
import tsParser from "@typescript-eslint/parser";
import { createTypeScriptImportResolver } from "eslint-import-resolver-typescript";

export const commonLintConfig = defineConfig(
  {
    plugins: {
      "unused-imports": unusedImports,
      "import-x": importX
    }
  },
  eslint.configs.recommended,
  prettier,
  ...oxlint.configs["flat/recommended"],
  {
    ignores: [
      "**/benchmark/**",
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
      "**/graphql/**",
      "scripts/benchmark*.ts"
    ]
  }
);

export const getLintModuleConfiguration = ({ files, tsConfigPath, extraRules }) =>
  defineConfig({
    files,
    languageOptions: {
      parser: tsParser,
      ecmaVersion: "latest",
      sourceType: "module"
    },
    settings: {
      "import-x/resolver-next": [
        createTypeScriptImportResolver({
          project: tsConfigPath
        })
      ]
    },
    rules: {
      // Disable rules that TypeScript handles natively
      "no-undef": 0,
      "no-redeclare": 0,
      "no-unused-vars": 0,

      // conflict with recommendation
      "no-useless-escape": 0,
      "no-empty": 0,
      "comma-dangle": 0,
      "consistent-return": 0,
      "no-param-reassign": 0,
      "no-useless-return": 0,
      "no-case-declarations": 0,
      "no-async-promise-executor": 0,

      // extra rules help
      "no-return-assign": 2,
      "no-unneeded-ternary": 2,
      "spaced-comment": 2,

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
        {
          vars: "all",
          varsIgnorePattern: "^_",
          args: "after-used",
          argsIgnorePattern: "^_"
        }
      ],
      ...extraRules
    }
  });

export default defineConfig(
  ...commonLintConfig,
  ...getLintModuleConfiguration({
    files: ["**/*.ts", "**/*.tsx"],
    extraRules: {}
  })
);
