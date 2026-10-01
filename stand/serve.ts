// The stand's page — every file on the stand beside what was made of it, the
// gold and ours, and where the two differ. Three stands share it:
//
//   npm run stand        PDF → DOCX   http://localhost:4477   stand/files
//   npm run stand:html   xlsx → HTML  http://localhost:4478   stand/files-html
//   npm run stand:svg    xlsx → SVG   http://localhost:4479   stand/files-html
//
// A rebuild runs the stand's build script in a child process, so it reads the
// source as it is NOW: edit, press R (or run the build on the file), and the
// page redraws itself when the new report lands.

import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { dirname, extname, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { ServerResponse } from 'node:http';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');

/** One picture column of a sheet stand: the file prefix of its pictures and what it shows. */
interface Layer {
  readonly key: string;
  readonly label: string;
}

/** One stand: where its files and its output are, what builds them, and the page. */
interface Bench {
  readonly page: string;
  readonly build: string;
  readonly files: string;
  readonly out: string;
  /** What a file on the stand is. */
  readonly kind: RegExp;
  readonly port: number;
  /** For a page that shows any of several stands: its title, columns and links. */
  readonly title?: string;
  readonly layers?: ReadonlyArray<Layer>;
  /** A file of a build to open beside its pictures, by its path in the build's folder. */
  readonly links?: ReadonlyArray<{ readonly label: string; readonly path: string }>;
}

const DIFF_LAYER: Layer = { key: 'diff', label: 'Diff · red gold only · blue ours only' };

const BENCHES: Readonly<Record<string, Bench>> = {
  docx: {
    page: 'index.html',
    build: 'build.ts',
    files: resolve(here, 'files'),
    out: resolve(here, 'out'),
    kind: /\.pdf$/iu,
    port: 4477,
  },
  html: {
    page: 'html.html',
    build: 'html-build.ts',
    files: resolve(here, 'files-html'),
    out: resolve(here, 'out-html'),
    kind: /\.(?:xlsx|xlsm|xls)$/iu,
    port: 4478,
    title: 'XLSX → HTML stand',
    layers: [
      { key: 'source', label: 'Source · LibreOffice, the sheet whole' },
      { key: 'gold', label: 'Gold · LibreOffice HTML' },
      { key: 'ours', label: 'Ours · Ream HTML' },
      DIFF_LAYER,
    ],
    links: [
      { label: 'gold page', path: 'gold/gold.html' },
      { label: 'our page', path: 'ours.html' },
      { label: 'source.pdf', path: 'source.pdf' },
    ],
  },
  svg: {
    page: 'html.html',
    build: 'svg-build.ts',
    files: resolve(here, 'files-html'),
    out: resolve(here, 'out-svg'),
    kind: /\.(?:xlsx|xlsm|xls)$/iu,
    port: 4479,
    title: 'XLSX → SVG stand',
    layers: [
      { key: 'gold', label: 'Gold · LibreOffice, the sheet whole' },
      { key: 'ours', label: 'Ours · Ream SVG' },
      DIFF_LAYER,
    ],
    links: [
      { label: 'our svg', path: 'ours.svg' },
      { label: 'gold.pdf', path: 'gold.pdf' },
    ],
  },
};

function pick(name: string): Bench {
  const found = BENCHES[name];
  if (!found) throw new Error(`no stand "${name}": ${Object.keys(BENCHES).join(' or ')}`);
  return found;
}

const which = process.argv[2] ?? 'docx';
const bench = pick(which);
const FILES_DIR = bench.files;
const OUT_DIR = bench.out;
const PORT = Number(process.env.STAND_PORT ?? bench.port);

const TYPES: Readonly<Record<string, string>> = {
  '.html': 'text/html; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ttf': 'font/ttf',
  '.pdf': 'application/pdf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

/** The builds asked for, run one at a time in the order asked. */
const queue: Array<string> = [];
let building: string | null = null;
let log = '';

/** The files on the stand, by the name each goes by there (its own, minus the extension). */
function files(): Map<string, string> {
  if (!existsSync(FILES_DIR)) return new Map();
  return new Map(
    readdirSync(FILES_DIR)
      .filter((f) => bench.kind.test(f))
      .map((f) => [f.replace(bench.kind, ''), f] as const)
      .sort((a, b) => a[0].localeCompare(b[0])),
  );
}

function names(): Array<string> {
  return [...files().keys()];
}

function state(): unknown {
  const docs = names().map((name) => {
    const file = resolve(OUT_DIR, name, 'report.json');
    return {
      name,
      report: existsSync(file) ? (JSON.parse(readFileSync(file, 'utf8')) as unknown) : null,
    };
  });
  return {
    building,
    queue,
    log,
    docs,
    bench: { title: bench.title, layers: bench.layers, links: bench.links },
  };
}

function next(): void {
  if (building !== null) return;
  const name = queue.shift();
  if (name === undefined) return;
  building = name;
  log = '';
  const file = files().get(name);
  if (file === undefined) {
    building = null;
    next();
    return;
  }
  const child = spawn(
    resolve(root, 'node_modules/.bin/tsx'),
    [resolve(here, bench.build), resolve(FILES_DIR, file)],
    { cwd: root },
  );
  const take = (chunk: Buffer): void => {
    log = (log + chunk.toString()).slice(-20_000);
  };
  child.stdout.on('data', take);
  child.stderr.on('data', take);
  child.on('close', () => {
    building = null;
    next();
  });
}

function ask(name: string): void {
  const wanted = name === '*' ? names() : names().filter((n) => n === name);
  for (const n of wanted) if (n !== building && !queue.includes(n)) queue.push(n);
  next();
}

function send(res: ServerResponse, status: number, type: string, body: string | Buffer): void {
  res.writeHead(status, { 'content-type': type, 'cache-control': 'no-store' });
  res.end(body);
}

const server = createServer((req, res) => {
  const url = new URL(req.url ?? '/', `http://localhost:${String(PORT)}`);
  const path = decodeURIComponent(url.pathname);
  if (path === '/' || path === '/index.html') {
    send(res, 200, TYPES['.html']!, readFileSync(resolve(here, bench.page)));
    return;
  }
  if (path === '/api/state') {
    send(res, 200, TYPES['.json']!, JSON.stringify(state()));
    return;
  }
  if (path === '/api/build' && req.method === 'POST') {
    ask(url.searchParams.get('name') ?? '');
    send(res, 202, TYPES['.json']!, JSON.stringify(state()));
    return;
  }
  if (path.startsWith('/files/')) {
    const file = resolve(FILES_DIR, `.${path.slice('/files'.length)}`);
    // The file itself, to open beside what was made of it.
    if (file.startsWith(FILES_DIR + sep) && existsSync(file) && statSync(file).isFile()) {
      send(
        res,
        200,
        TYPES[extname(file).toLowerCase()] ?? 'application/octet-stream',
        readFileSync(file),
      );
      return;
    }
  }
  if (path.startsWith('/out/')) {
    const file = resolve(OUT_DIR, `.${path.slice('/out'.length)}`);
    // Nothing outside the stand's output is served.
    if (file.startsWith(OUT_DIR + sep) && existsSync(file) && statSync(file).isFile()) {
      send(
        res,
        200,
        TYPES[extname(file).toLowerCase()] ?? 'application/octet-stream',
        readFileSync(file),
      );
      return;
    }
  }
  send(res, 404, 'text/plain; charset=utf-8', 'not found');
});

server.listen(PORT, () => {
  process.stdout.write(`stand (${which}): http://localhost:${String(PORT)}\n`);
});
