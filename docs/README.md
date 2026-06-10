# Ream docs site

[Astro Starlight](https://starlight.astro.build) documentation site for **Ream**
(`reamkit`), deployed to GitHub Pages at
<https://alex-krassavin.github.io/reamkit/>.

- **Guides** (`src/content/docs/guides/`) — hand-written.
- **API Reference** — generated from the typed public surface (`../src/index.ts`
  + the `reamkit/document-model` subpath) by
  [`starlight-typedoc`](https://github.com/HiDeoo/starlight-typedoc) on every
  build. Lives under `src/content/docs/api/` and is **git-ignored** (rebuilt, not
  committed).

## Commands

Run from this `docs/` directory:

| Command | Action |
| :-- | :-- |
| `npm install` | Install docs dependencies |
| `npm run dev` | Local dev server at `localhost:4321` |
| `npm run build` | Build the production site to `./dist/` |
| `npm run preview` | Preview the build locally |

The API reference resolves types from the module source, so the root package
must be installed too (`npm ci` in the repo root) before building.

## Deploy

`.github/workflows/docs.yml` builds this site and publishes `docs/dist` to
GitHub Pages on every push to `main` that touches `docs/**` or `src/**`. Enable
Pages once in repo **Settings → Pages → Source: GitHub Actions**.
