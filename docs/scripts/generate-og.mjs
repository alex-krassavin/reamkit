// Generates the social-share / Open Graph image (public/og.png, 1200×630) from
// an inline SVG, rasterised with sharp (already a docs dependency). Run after
// changing the branding:
//
//   node scripts/generate-og.mjs
//
// The PNG is committed as a static asset; og:image / twitter:image in
// astro.config.mjs point at https://reamkit.dev/og.png. System fonts (Georgia +
// a monospace) are used so the render is reproducible without bundling webfonts.

import sharp from 'sharp';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const out = join(here, '..', 'public', 'og.png');

const svg = `<svg width="1200" height="630" viewBox="0 0 1200 630" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#fdfcf9"/>
      <stop offset="1" stop-color="#f5eddf"/>
    </linearGradient>
  </defs>
  <rect width="1200" height="630" fill="url(#bg)"/>
  <rect x="24" y="24" width="1152" height="582" rx="28" fill="#ffffff" fill-opacity="0.5" stroke="#e6dcc9" stroke-width="2"/>
  <rect x="24" y="24" width="1152" height="9" rx="4.5" fill="#c2632b"/>

  <!-- brand -->
  <g transform="translate(84,86)">
    <rect x="0" y="0" width="54" height="54" rx="13" fill="#c2632b"/>
    <rect x="14" y="15" width="26" height="4.5" rx="2.25" fill="#ffffff"/>
    <rect x="14" y="25" width="26" height="4.5" rx="2.25" fill="#ffffff"/>
    <rect x="14" y="35" width="17" height="4.5" rx="2.25" fill="#ffffff"/>
    <text x="74" y="40" font-family="Georgia, 'Times New Roman', serif" font-size="42" font-weight="700" fill="#221c13">Ream</text>
  </g>

  <!-- headline -->
  <text x="84" y="278" font-family="Georgia, 'Times New Roman', serif" font-size="74" font-weight="700" fill="#221c13">Read any office document.</text>
  <text x="84" y="364" font-family="Georgia, 'Times New Roman', serif" font-size="74" font-weight="700" fill="#c2632b">Convert it to anything.</text>

  <!-- reads / writes -->
  <text x="84"  y="452" font-family="Menlo, 'Courier New', monospace" font-size="21" font-weight="700" fill="#c2632b" letter-spacing="1.5">READS</text>
  <text x="210" y="452" font-family="Menlo, 'Courier New', monospace" font-size="25" fill="#5a513f">docx · xlsx · pptx · pdf · doc · xls · ppt</text>
  <text x="84"  y="496" font-family="Menlo, 'Courier New', monospace" font-size="21" font-weight="700" fill="#c2632b" letter-spacing="1.5">WRITES</text>
  <text x="210" y="496" font-family="Menlo, 'Courier New', monospace" font-size="25" fill="#5a513f">pdf · svg · html · docx · xlsx</text>

  <!-- footer -->
  <line x1="84" y1="548" x2="1116" y2="548" stroke="#e6dcc9" stroke-width="2"/>
  <text x="84"   y="586" font-family="Menlo, 'Courier New', monospace" font-size="24" font-weight="700" fill="#c2632b">reamkit.dev</text>
  <text x="1116" y="585" text-anchor="end" font-family="Menlo, 'Courier New', monospace" font-size="20" fill="#8a7c64">Built from ECMA-376 &amp; ISO 32000 · MIT · TypeScript</text>
</svg>`;

await sharp(Buffer.from(svg)).png().toFile(out);
console.log('wrote', out);
