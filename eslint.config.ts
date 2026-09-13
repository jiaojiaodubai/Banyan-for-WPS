import js from "@eslint/js"
import globals from "globals"
import tseslint from "typescript-eslint"
import json from "@eslint/json"
import stylistic from "@stylistic/eslint-plugin"
import { defineConfig } from "eslint/config"

export default defineConfig([
  {
    ignores: [".vscode/**", "dev/**", "dist/**", "release/**", "node_modules/**"],
  },
  {
    files: ["src/**/*.{js,mjs,cjs}", "./*.js"],
    plugins: { js, "@stylistic": stylistic },
    extends: ["js/recommended"],
    languageOptions: { globals: globals.browser },
    rules: {
      "@stylistic/brace-style": [
        "error",
        "stroustrup",
        { allowSingleLine: true },
      ],
      "linebreak-style": ["error", "unix"],
    },
  },
  tseslint.configs.recommended,
  {
    // src/typings 下的 .d.ts 逐字镜像后端契约，其中的 `any` 属于契约本身
    // （例如 style script 的安全视图允许未知键），不按本项目代码规范检查。
    files: ["src/typings/**/*.d.ts"],
    rules: {
      "@typescript-eslint/no-explicit-any": "off",
    },
  },
  {
    files: ["src/**/*.{ts,mts}", "test/**/*.{ts,mts}", "./*.ts"],
    plugins: { "@stylistic": stylistic },
    rules: {
      "@stylistic/brace-style": [
        "error",
        "stroustrup",
        { allowSingleLine: true },
      ],
      "linebreak-style": ["error", "unix"],
      "@typescript-eslint/no-unused-vars": [
        "error",
        {
          argsIgnorePattern: "^_",
        },
      ],
    },
  },
  {
    files: ["**/*.jsonc", "**/*.json"],
    ignores: ["package-lock.json"],
    plugins: { json },
    language: "json/jsonc",
    extends: ["json/recommended"],
  },
])
