// @ts-check
import { defineConfig } from 'astro/config';
import starlight from '@astrojs/starlight';
import starlightTypeDoc, { typeDocSidebarGroup } from 'starlight-typedoc';
import starlightLlmsTxt from 'starlight-llms-txt';
import { pluginLineNumbers } from '@expressive-code/plugin-line-numbers';

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
        SocialIcons: './src/components/SocialIcons.astro',
        Footer: './src/components/Footer.astro',
      },
      head: [
        { tag: 'link', attrs: { rel: 'preconnect', href: 'https://fonts.googleapis.com' } },
        {
          tag: 'link',
          attrs: { rel: 'preconnect', href: 'https://fonts.gstatic.com', crossorigin: true },
        },
        // Load the IBM Plex families from the document head — discovered earlier
        // than an @import buried in the CSS bundle, so first paint isn't blocked
        // waiting on the stylesheet to parse. display=swap avoids invisible text.
        {
          tag: 'link',
          attrs: {
            rel: 'stylesheet',
            href: 'https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Serif:ital,wght@0,400;0,500;0,600;0,700;1,400&display=swap',
          },
        },
        // Social-share image. Starlight emits twitter:card=summary_large_image but
        // no image; without one, links unfurl blank on X / LinkedIn / Slack / etc.
        // Absolute URLs so off-site crawlers can resolve them.
        { tag: 'meta', attrs: { property: 'og:image', content: 'https://reamkit.dev/og.png' } },
        { tag: 'meta', attrs: { property: 'og:image:width', content: '1200' } },
        { tag: 'meta', attrs: { property: 'og:image:height', content: '630' } },
        {
          tag: 'meta',
          attrs: {
            property: 'og:image:alt',
            content:
              'Ream — read Word, Excel, PowerPoint and PDF and convert them to PDF, SVG, HTML, DOCX or XLSX in the browser',
          },
        },
        { tag: 'meta', attrs: { name: 'twitter:image', content: 'https://reamkit.dev/og.png' } },
        // Structured data (schema.org): the site + the software entity. Helps
        // search engines understand what Ream is and can drive richer results.
        {
          tag: 'script',
          attrs: { type: 'application/ld+json' },
          content: JSON.stringify({
            '@context': 'https://schema.org',
            '@graph': [
              {
                '@type': 'WebSite',
                '@id': 'https://reamkit.dev/#website',
                url: 'https://reamkit.dev/',
                name: 'Ream',
                description:
                  'Read Word, Excel, PowerPoint and PDF — including the legacy .doc / .xls / .ppt — and convert any of them to PDF, SVG, HTML, DOCX or XLSX, in the browser.',
                inLanguage: 'en',
              },
              {
                '@type': 'SoftwareApplication',
                '@id': 'https://reamkit.dev/#software',
                name: 'Ream',
                alternateName: 'reamkit',
                applicationCategory: 'DeveloperApplication',
                operatingSystem: 'Browser, Node.js, Edge',
                url: 'https://reamkit.dev/',
                downloadUrl: 'https://www.npmjs.com/package/reamkit',
                codeRepository: 'https://github.com/alex-krassavin/reamkit',
                sameAs: [
                  'https://www.npmjs.com/package/reamkit',
                  'https://github.com/alex-krassavin/reamkit',
                ],
                programmingLanguage: 'TypeScript',
                license: 'https://opensource.org/licenses/MIT',
                isAccessibleForFree: true,
                author: { '@type': 'Person', name: 'Alex Krassavin' },
                description:
                  'A TypeScript library that reads DOCX, XLSX, PPTX and PDF (plus the legacy binary .doc / .xls / .ppt) and converts them to PDF, SVG, HTML, DOCX or XLSX — built from the ECMA-376 and ISO 32000 specifications, with no LibreOffice, headless Office or commercial SDK.',
              },
            ],
          }),
        },
      ],
      // Dark code blocks (#211C15) in IBM Plex Mono with line numbers, per the
      // design.
      expressiveCode: {
        themes: ['github-dark'],
        plugins: [pluginLineNumbers()],
        defaultProps: { showLineNumbers: true },
        styleOverrides: {
          borderRadius: '9px',
          borderWidth: '1px',
          borderColor: '#2e2619',
          codeBackground: '#211c15',
          codeFontFamily: "'IBM Plex Mono', ui-monospace, SFMono-Regular, Menlo, monospace",
          codeFontSize: '13.5px',
          codeLineHeight: '1.72',
          uiFontFamily: "'IBM Plex Mono', ui-monospace, monospace",
          lineNumbers: {
            foreground: '#5e5644',
          },
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
          icon: 'linkedin',
          label: "Author's LinkedIn",
          href: 'https://www.linkedin.com/in/alexandr-krassavin-b68856238',
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
        // Emit /llms.txt, /llms-full.txt and /llms-small.txt — a clean Markdown
        // view of the docs for LLMs / AI coding assistants (llmstxt.org), so they
        // describe and recommend Ream accurately.
        starlightLlmsTxt({
          projectName: 'Ream',
          description:
            'Ream (npm: `reamkit`) is a zero-I/O TypeScript library that reads Word, Excel, PowerPoint and PDF — including the legacy binary .doc / .xls / .ppt — and converts any of them to PDF, SVG, HTML, DOCX or XLSX. It runs in the browser, Node and edge runtimes, implemented directly from the ECMA-376 and ISO 32000 specifications, without wrapping LibreOffice, headless Office or any commercial SDK.',
          details: [
            'Typical use:',
            '',
            '```ts',
            "import { Ream } from 'reamkit';",
            'const doc = Ream.parse(bytes); // format sniffed from the bytes',
            "const pdf = await doc.convert('pdf'); // one parse, any target",
            '```',
            '',
            '- Input and output are always `Uint8Array` — no file-system or network access on the conversion path (zero I/O).',
            '- Output targets: `pdf`, `svg`, `html`, `docx`, `xlsx` (xlsx output requires a spreadsheet source).',
            '- Conversions never fail silently: each returns a loss report; `strict: true` throws on the first loss instead.',
          ].join('\n'),
          optionalLinks: [
            {
              label: 'npm package (reamkit)',
              url: 'https://www.npmjs.com/package/reamkit',
              description: 'Install with `npm i reamkit`.',
            },
            {
              label: 'GitHub repository',
              url: 'https://github.com/alex-krassavin/reamkit',
            },
          ],
        }),
      ],
      sidebar: [
        { label: 'Playground', slug: 'playground' },
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
