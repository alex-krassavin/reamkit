// Flat ESLint config, aligned to @tanstack/config.
//
// The base rule set is TanStack's shared `tanstackConfig` (correctness +
// import hygiene + type-aware TypeScript rules + `node:` protocol). Formatting
// is still owned by Prettier — TanStack's eslint-stylistic only enforces
// `spaced-comment`, so there is no overlap/fight with Prettier.
//
// tanstackConfig runs the type-checked rules with `parserOptions.project`, so
// we point it at tsconfig.eslint.json (covers src + tests + scripts + configs).

import { tanstackConfig } from '@tanstack/config/eslint';

export default [
  {
    ignores: [
      'dist/**',
      'node_modules/**',
      'corpus/**',
      'docs/**',
      'tests/fixtures/fonts/**',
      'tests/output/**',
    ],
  },
  ...tanstackConfig,
  {
    // Resolve type-aware rules against the lint project graph (all files),
    // not the build's src-only tsconfig.
    languageOptions: {
      parserOptions: {
        project: './tsconfig.eslint.json',
        tsconfigRootDir: import.meta.dirname,
      },
    },
  },
  {
    rules: {
      // `!` is routine here: noUncheckedIndexedAccess makes `arr[i]` possibly
      // undefined, and we assert after explicit bounds/exists checks.
      '@typescript-eslint/no-non-null-assertion': 'off',
      // Underscore-prefixed bindings are intentionally unused.
      '@typescript-eslint/no-unused-vars': [
        'warn',
        { argsIgnorePattern: '^_', varsIgnorePattern: '^_', caughtErrors: 'none' },
      ],
    },
  },
];
