import eslintPlugin from '@eslint/js';
import googleappsscript from 'eslint-plugin-googleappsscript';
import jsoncPlugin from 'eslint-plugin-jsonc';
import prettierPlugin from 'eslint-plugin-prettier';
import eslintPluginPrettierRecommended from 'eslint-plugin-prettier/recommended';
import globals from 'globals';

export default [
  eslintPlugin.configs.recommended,
  eslintPluginPrettierRecommended,
  { ignores: ['dist/**', 'build/**', 'node_modules/**', 'package-lock.json'] },
  {
    files: ['src/**/*.js'],
    plugins: {
      prettier: prettierPlugin,
      googleappsscript,
    },
    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'script',
      globals: {
        ...googleappsscript.environments.googleappsscript.globals,
      },
    },
    rules: {
      // GAS loads all .gs files into a single global scope — cross-file calls are valid
      'no-undef': 'off',
      'no-console': 'warn',
      'no-underscore-dangle': 'off',
      'no-unused-vars': [
        'warn',
        {
          ignoreRestSiblings: true,
          argsIgnorePattern: 'res|next|^err|^ignore|^_',
          caughtErrors: 'none',
        },
      ],
      'no-unused-expressions': 'warn',
      'prefer-const': 'error',
      'prefer-template': 'error',
      'prefer-arrow-callback': 'error',
      'arrow-spacing': 'error',
      'object-shorthand': 'error',
      'prettier/prettier': [
        'error',
        {
          trailingComma: 'es5',
          singleQuote: true,
          printWidth: 120,
          endOfLine: 'auto',
          semi: true,
          tabWidth: 2,
        },
      ],
    },
  },

  // Config and build scripts use Node globals and ES modules
  {
    files: ['scripts/**/*.js', '*.config.js', 'eslint.config.js'],
    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'module',
      globals: {
        ...globals.node,
      },
    },
  },

  ...jsoncPlugin.configs['flat/recommended-with-jsonc'],
  {
    files: ['**/*.json', '**/*.jsonc'],
    rules: {
      'jsonc/sort-keys': 'error',
      'jsonc/no-dupe-keys': 'error',
      'jsonc/no-comments': 'off',
    },
  },

  {
    files: ['**/*.test.js', '**/*.spec.js', '**/__tests__/**/*.js'],
    languageOptions: {
      globals: {
        ...globals.jest,
      },
    },
    rules: {
      'no-console': 'off',
    },
  },
];
