// @ts-check
import { defineConfig } from 'astro/config';
import starlight from '@astrojs/starlight';
import starlightTypeDoc, { typeDocSidebarGroup } from 'starlight-typedoc';

// Served at https://reamkit.dev (GitHub Pages + custom domain).
export default defineConfig({
  site: 'https://reamkit.dev',
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
        {
          icon: 'email',
          label: 'info@reamkit.dev',
          href: 'mailto:info@reamkit.dev',
        },
      ],
      plugins: [
        // Generates the API Reference under src/content/docs/api/ from the typed public
        // surface — the root API (../src/index.ts) and the document-model subpath.
        starlightTypeDoc({
          entryPoints: ['../src/index.ts', '../src/core/document-model/index.ts'],
          tsconfig: '../tsconfig.json',
          typeDoc: {
            skipErrorChecking: true,
            entryFileName: 'index',
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
