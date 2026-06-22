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
        'Read Word, Excel, PowerPoint and PDF — including the legacy .doc / .xls / .ppt — and convert any of them to PDF, SVG, HTML, DOCX or XLSX. In the browser, from the ECMA-376 and ISO 32000 specifications.',
      // The warm-paper design (Claude Design handoff): theme tokens + restyled
      // sidebar / TOC / cards / search, and the branded site title.
      customCss: ['./src/styles/theme.css'],
      components: {
        SiteTitle: './src/components/SiteTitle.astro',
        Sidebar: './src/components/Sidebar.astro',
      },
      head: [
        { tag: 'link', attrs: { rel: 'preconnect', href: 'https://fonts.googleapis.com' } },
        {
          tag: 'link',
          attrs: { rel: 'preconnect', href: 'https://fonts.gstatic.com', crossorigin: true },
        },
      ],
      // Dark code blocks (#211C15) in IBM Plex Mono, matching the design.
      expressiveCode: {
        themes: ['github-dark'],
        styleOverrides: {
          borderRadius: '9px',
          borderWidth: '1px',
          borderColor: '#2e2619',
          codeBackground: '#211c15',
          codeFontFamily: "'IBM Plex Mono', ui-monospace, SFMono-Regular, Menlo, monospace",
          codeFontSize: '13.5px',
          codeLineHeight: '1.72',
          uiFontFamily: "'IBM Plex Mono', ui-monospace, monospace",
          frames: {
            editorTabBarBackground: '#2a2318',
            editorActiveTabBackground: '#2a2318',
            editorActiveTabForeground: '#c99a63',
            editorActiveTabIndicatorBottomColor: '#c2632b',
            editorTabBarBorderBottomColor: '#342a1c',
            terminalTitlebarBackground: '#2a2318',
            terminalTitlebarForeground: '#c99a63',
            terminalBackground: '#211c15',
            frameBoxShadowCssValue: '0 16px 38px -26px rgba(40,28,10,.7)',
          },
        },
      },
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
