import { existsSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { defineConfig, mergeConfig } from 'vite';
import { tanstackViteConfig } from '@tanstack/config/vite';

// Library build via @tanstack/config. Produces ESM JS + .d.ts into dist/esm.
// Two entries: the public API and the typed document-model subpath. The
// library is browser-pure (no node:* imports). fflate / fast-xml-parser stay
// external (runtime deps).
//
// Source uses extensionless `@/*` (and relative) imports. The emitted .js
// carries proper extensions via rollup; the .d.ts files keep relative
// specifiers, so the hook below restores explicit extensions for NodeNext
// consumers — appending `/index.js` for directory imports and `.js` for files
// (resolved against the source tree so we pick the right one).
const SRC = resolve('./src');

function fixDeclarationExtensions(declFilePath: string, content: string): string {
  // Map the emitted .d.ts path back to its source dir to resolve specifiers.
  const sourceDir = dirname(declFilePath.replace(`${resolve('./dist')}/esm`, SRC));
  return content.replace(
    /(\bfrom\s*["'])(\.\.?\/[^"']+?)(["'])/g,
    (full, pre: string, spec: string, post: string) => {
      if (/\.[a-z]+$/i.test(spec)) return full; // already has an extension
      const target = resolve(sourceDir, spec);
      const isDir = existsSync(target) && statSync(target).isDirectory();
      return `${pre}${spec}${isDir ? '/index.js' : '.js'}${post}`;
    },
  );
}

const tanstack = tanstackViteConfig({
  entry: ['./src/index.ts', './src/document-model/index.ts'],
  srcDir: './src',
  cjs: false,
  outDir: './dist',
  beforeWriteDeclarationFile: (filePath, content) => fixDeclarationExtensions(filePath, content),
});

// Keep the published package lean: no source maps (we do not ship src, so maps
// would dangle). Our override is the second arg so it wins the merge.
export default mergeConfig(tanstack, defineConfig({ build: { sourcemap: false } }));
