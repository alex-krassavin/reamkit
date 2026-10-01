// Headless Chrome over the DevTools protocol — what the HTML stand draws its
// pages with.
//
// One browser for a whole build: a tab goes from file to file, and the parts
// of a page that are wanted — a sheet apiece — are cut out of it as pictures
// of their own. The protocol rather than `--screenshot`: that flag draws the
// window and no more, where a sheet runs on past it, and a headless Chrome
// asked for one at times never exits.

import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import type { ChildProcessByStdio } from 'node:child_process';
import type { Readable } from 'node:stream';

const CANDIDATES = [
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  '/Applications/Chromium.app/Contents/MacOS/Chromium',
];

/** How long Chrome may take to say where it listens. */
const START_TIMEOUT_MS = 20_000;
/** How long a page may take to load; past it, it is drawn as far as it got. */
const LOAD_TIMEOUT_MS = 20_000;
/** How long any one command may take. */
const CALL_TIMEOUT_MS = 60_000;

/** A rectangle of the page, in CSS pixels from its top left. */
export interface Box {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

type Fields = Record<string, unknown>;

interface Message {
  readonly id?: number;
  readonly method?: string;
  readonly sessionId?: string;
  readonly result?: Fields;
  readonly error?: { readonly message: string };
}

interface Waiter {
  readonly ok: (result: Fields) => void;
  readonly no: (e: Error) => void;
}

export class Chrome {
  private last = 0;
  private session: string | undefined;
  private readonly waiting = new Map<number, Waiter>();
  private readonly listeners = new Set<(m: Message) => void>();

  private constructor(
    private readonly proc: ChildProcessByStdio<null, null, Readable>,
    private readonly socket: WebSocket,
    private readonly profile: string,
  ) {
    socket.addEventListener('message', (e: MessageEvent) => {
      const m = JSON.parse(String(e.data)) as Message;
      if (m.id !== undefined) {
        const w = this.waiting.get(m.id);
        if (!w) return;
        this.waiting.delete(m.id);
        if (m.error) w.no(new Error(m.error.message));
        else w.ok(m.result ?? {});
        return;
      }
      for (const f of this.listeners) f(m);
    });
  }

  /** A fresh headless Chrome on a profile of its own, with one tab the size of `viewport`. */
  static async launch(viewport: { width: number; height: number }): Promise<Chrome> {
    const binary = CANDIDATES.find((c) => existsSync(c)) ?? 'google-chrome';
    const profile = mkdtempSync(join(tmpdir(), 'ream-stand-chrome-'));
    const proc = spawn(
      binary,
      [
        '--headless=new',
        '--disable-gpu',
        '--hide-scrollbars',
        '--no-first-run',
        '--no-default-browser-check',
        `--user-data-dir=${profile}`,
        '--remote-debugging-port=0',
        'about:blank',
      ],
      { stdio: ['ignore', 'ignore', 'pipe'] },
    );
    let address: string;
    try {
      address = await new Promise<string>((ok, no) => {
        let said = '';
        const timer = setTimeout(() => {
          no(new Error('Chrome gave no DevTools address'));
        }, START_TIMEOUT_MS);
        proc.stderr.on('data', (chunk: Buffer) => {
          said += chunk.toString();
          const m = /ws:\/\/\S+/u.exec(said);
          if (m) {
            clearTimeout(timer);
            ok(m[0]);
          }
        });
        proc.on('exit', () => {
          clearTimeout(timer);
          no(new Error(`Chrome exited before it listened: ${said.slice(-300)}`));
        });
      });
    } catch (e) {
      proc.kill('SIGKILL');
      rmSync(profile, { recursive: true, force: true });
      throw e;
    }
    const socket = new WebSocket(address);
    await new Promise<void>((ok, no) => {
      socket.addEventListener('open', () => ok(), { once: true });
      socket.addEventListener('error', () => no(new Error('no DevTools connection')), {
        once: true,
      });
    });
    const chrome = new Chrome(proc, socket, profile);
    const { targetId } = await chrome.call('Target.createTarget', { url: 'about:blank' });
    const { sessionId } = await chrome.call('Target.attachToTarget', { targetId, flatten: true });
    chrome.session = String(sessionId);
    await chrome.call('Page.enable');
    await chrome.size(viewport.width, viewport.height);
    return chrome;
  }

  /** One protocol command — to the tab, once there is one — and its answer. */
  call(method: string, params: Fields = {}): Promise<Fields> {
    const id = ++this.last;
    return new Promise<Fields>((ok, no) => {
      const timer = setTimeout(() => {
        this.waiting.delete(id);
        no(new Error(`${method} took over ${String(CALL_TIMEOUT_MS / 1000)} s`));
      }, CALL_TIMEOUT_MS);
      this.waiting.set(id, {
        ok: (r) => {
          clearTimeout(timer);
          ok(r);
        },
        no: (e) => {
          clearTimeout(timer);
          no(e);
        },
      });
      this.socket.send(
        JSON.stringify({
          id,
          method,
          params,
          ...(this.session !== undefined ? { sessionId: this.session } : {}),
        }),
      );
    });
  }

  /** The tab's window, in CSS pixels at one device pixel apiece. */
  async size(width: number, height: number): Promise<void> {
    await this.call('Emulation.setDeviceMetricsOverride', {
      width: Math.round(width),
      height: Math.round(height),
      deviceScaleFactor: 1,
      mobile: false,
    });
  }

  /** Load a file and wait for it, its pictures with it — or for the timeout. */
  async open(file: string): Promise<void> {
    let listen: ((m: Message) => void) | undefined;
    const loaded = new Promise<void>((ok) => {
      const timer = setTimeout(ok, LOAD_TIMEOUT_MS);
      listen = (m: Message): void => {
        if (m.method !== 'Page.loadEventFired' || m.sessionId !== this.session) return;
        clearTimeout(timer);
        ok();
      };
      this.listeners.add(listen);
    });
    try {
      await this.call('Page.navigate', { url: pathToFileURL(file).href });
      await loaded;
    } finally {
      if (listen) this.listeners.delete(listen);
    }
  }

  /** An expression evaluated in the page, and its value as JSON brings it back. */
  async evaluate(expression: string): Promise<unknown> {
    const r = await this.call('Runtime.evaluate', {
      expression,
      returnByValue: true,
      awaitPromise: true,
    });
    const thrown = r.exceptionDetails as { text?: string } | undefined;
    if (thrown) throw new Error(`in the page: ${thrown.text ?? 'an exception'}`);
    return (r.result as { value?: unknown }).value;
  }

  /** A rectangle of the page as a PNG, drawn past the window where it runs on. */
  async shoot(box: Box): Promise<Uint8Array> {
    const r = await this.call('Page.captureScreenshot', {
      format: 'png',
      captureBeyondViewport: true,
      clip: { x: box.x, y: box.y, width: box.width, height: box.height, scale: 1 },
    });
    return new Uint8Array(Buffer.from(String(r.data), 'base64'));
  }

  close(): void {
    for (const w of this.waiting.values()) w.no(new Error('Chrome closed'));
    this.waiting.clear();
    this.socket.close();
    this.proc.kill('SIGKILL');
    try {
      rmSync(this.profile, { recursive: true, force: true });
    } catch {
      // Still being let go of by the browser; the system's temp sweep takes it.
    }
  }
}
