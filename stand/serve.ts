// The stand's page — http://localhost:4477 — every file on the stand with its
// source page, LibreOffice's .docx (the gold), ours, and where the two differ.
//
// A rebuild runs `stand/build.ts` in a child process, so it reads the reader's
// source as it is NOW: edit, press R (or run `npm run stand:build -- <file>`),
// and the page redraws itself when the new report lands.
//
//   npm run stand

import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { dirname, extname, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { ServerResponse } from 'node:http';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');
const FILES_DIR = resolve(here, 'files');
const OUT_DIR = resolve(here, 'out');
const PORT = Number(process.env.STAND_PORT ?? 4477);

const TYPES: Readonly<Record<string, string>> = {
  '.html': 'text/html; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.pdf': 'application/pdf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};

/** The builds asked for, run one at a time in the order asked. */
const queue: Array<string> = [];
let building: string | null = null;
let log = '';

function names(): Array<string> {
  if (!existsSync(FILES_DIR)) return [];
  return readdirSync(FILES_DIR)
    .filter((f) => /\.pdf$/iu.test(f))
    .map((f) => f.replace(/\.pdf$/iu, ''))
    .sort((a, b) => a.localeCompare(b));
}

function state(): unknown {
  const docs = names().map((name) => {
    const file = resolve(OUT_DIR, name, 'report.json');
    return {
      name,
      report: existsSync(file) ? (JSON.parse(readFileSync(file, 'utf8')) as unknown) : null,
    };
  });
  return { building, queue, log, docs };
}

function next(): void {
  if (building !== null) return;
  const name = queue.shift();
  if (name === undefined) return;
  building = name;
  log = '';
  const child = spawn(
    resolve(root, 'node_modules/.bin/tsx'),
    [resolve(here, 'build.ts'), resolve(FILES_DIR, `${name}.pdf`)],
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
    send(res, 200, TYPES['.html']!, readFileSync(resolve(here, 'index.html')));
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
  if (path.startsWith('/out/')) {
    const file = resolve(OUT_DIR, `.${path.slice('/out'.length)}`);
    // Nothing outside the stand's output is served.
    if (file.startsWith(OUT_DIR + sep) && existsSync(file) && statSync(file).isFile()) {
      send(res, 200, TYPES[extname(file)] ?? 'application/octet-stream', readFileSync(file));
      return;
    }
  }
  send(res, 404, 'text/plain; charset=utf-8', 'not found');
});

server.listen(PORT, () => {
  process.stdout.write(`stand: http://localhost:${String(PORT)}\n`);
});
