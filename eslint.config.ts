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
    files: ["src/**/*.{ts,mts}", "./*.ts"],
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
