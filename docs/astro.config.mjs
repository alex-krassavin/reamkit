// @ts-check
import { defineConfig } from 'astro/config';
import starlight from '@astrojs/starlight';
import starlightTypeDoc, { typeDocSidebarGroup } from 'starlight-typedoc';

// Served from GitHub Pages at https://alex-krassavin.github.io/reamkit/
export default defineConfig({
  site: 'https://alex-krassavin.github.io',
  base: '/reamkit',
  integrations: [
    starlight({
      title: 'Ream',
      description:
        'Convert Word (.docx) and Excel (.xlsx) to PDF in the browser — implemented from the ECMA-376 and ISO 32000 specifications.',
      social: [
        {
          icon: 'github',
          label: 'GitHub',
          href: 'https://github.com/alex-krassavin/reamkit',
        },
      ],
      plugins: [
        // Generates the API Reference under src/content/docs/api/ from the typed public
        // surface — the root API (../src/index.ts) and the document-model subpath.
        starlightTypeDoc({
          entryPoints: ['../src/index.ts', '../src/document-model/index.ts'],
          tsconfig: '../tsconfig.json',
          typeDoc: {
            skipErrorChecking: true,
            excludeInternal: true,
          },
        }),
      ],
      sidebar: [
        {
          label: 'Guides',
          items: [
            { label: 'Getting started', slug: 'guides/getting-started' },
            { label: 'Examples', slug: 'guides/examples' },
            { label: 'Concepts', slug: 'guides/concepts' },
            { label: 'Scope & limitations', slug: 'guides/scope' },
          ],
        },
        typeDocSidebarGroup,
      ],
    }),
  ],
});
