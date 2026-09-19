/*
 * Fygaro Catalog Automation
 * Loads the extension into a real Chrome and checks that its service worker
 * actually starts.
 *
 * Nothing else covers this. The integration suite loads the worker into a page
 * and stubs importScripts to a no-op, because the shared files are already there
 * as script tags, so the worker's own imports are never resolved by any test. A
 * path that points at nothing therefore passes the entire suite and then stops
 * the extension dead the first time it is loaded, with
 * "Failed to execute 'importScripts' on 'WorkerGlobalScope'" and nothing else
 * working at all.
 *
 * npm run lint resolves those paths statically, which catches a typo. This
 * catches the rest: a file that fails to parse, a shared file that throws while
 * it evaluates, a manifest Chrome refuses outright.
 *
 * The extension is loaded over the DevTools protocol rather than with
 * --load-extension, because that flag is ignored under --headless=new. That
 * needs the Extensions domain, which needs --enable-unsafe-extension-debugging,
 * and a Chrome new enough to have it. When it is not available this says so and
 * exits 0 rather than failing a build over a missing tool.
 */
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(join(dirname(fileURLToPath(import.meta.url)), '..'));

// The same places the browser test runner looks.
const CANDIDATES = [
  process.env.CHROME_PATH,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
  process.env.LOCALAPPDATA && `${process.env.LOCALAPPDATA}/Google/Chrome/Application/chrome.exe`,
  '/usr/bin/google-chrome',
  '/usr/bin/chromium',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
].filter(Boolean);

const chromePath = CANDIDATES.find((p) => existsSync(p));
if (!chromePath) {
  console.error('Chrome was not found. Set CHROME_PATH to its executable.');
  process.exit(2);
}

const profile = mkdtempSync(join(tmpdir(), 'fyg-extension-'));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const chrome = spawn(chromePath, [
  '--headless=new',
  '--disable-gpu',
  '--no-sandbox',
  '--no-first-run',
  '--no-default-browser-check',
  '--enable-unsafe-extension-debugging',
  `--user-data-dir=${profile}`,
  '--remote-debugging-port=0',
  'about:blank'
], { stdio: ['ignore', 'pipe', 'pipe'] });

let chromeErr = '';
chrome.stderr.on('data', (d) => { chromeErr += d.toString(); });

/* Chrome writes the port it actually took into the profile, so nothing here has
 * to guess at one that might already be busy. */
async function debuggingPort() {
  const file = join(profile, 'DevToolsActivePort');
  for (let i = 0; i < 100; i++) {
    if (existsSync(file)) {
      const port = readFileSync(file, 'utf8').split('\n')[0].trim();
      if (port) return port;
    }
    await sleep(100);
  }
  throw new Error(`Chrome never opened a debugging port.\n${chromeErr}`);
}

const waiting = new Map();
const exceptions = [];
let nextId = 1;
let workerSession = null;
let workerUrl = '';

function call(ws, method, params, sessionId) {
  const id = nextId++;
  ws.send(JSON.stringify({ id, method, params: params || {}, sessionId }));
  return new Promise((r) => waiting.set(id, r));
}

async function main() {
  const port = await debuggingPort();
  const version = await (await fetch(`http://127.0.0.1:${port}/json/version`)).json();

  const ws = new WebSocket(version.webSocketDebuggerUrl);
  await new Promise((r, j) => { ws.onopen = r; ws.onerror = () => j(new Error('could not attach to Chrome')); });

  ws.onmessage = (event) => {
    const msg = JSON.parse(event.data);

    if (msg.id && waiting.has(msg.id)) {
      waiting.get(msg.id)(msg);
      waiting.delete(msg.id);
      return;
    }

    if (msg.method === 'Target.attachedToTarget') {
      const { sessionId, targetInfo } = msg.params;
      if (targetInfo.type === 'service_worker' && targetInfo.url.includes('/src/background/')) {
        workerSession = sessionId;
        workerUrl = targetInfo.url;
      }
      call(ws, 'Runtime.enable', {}, sessionId);
    }

    if (msg.method === 'Runtime.exceptionThrown') {
      const d = msg.params.exceptionDetails;
      const text = d.exception ? d.exception.description || d.exception.value : d.text;
      exceptions.push(`${text}${d.url ? `\n      at ${d.url}:${d.lineNumber}` : ''}`);
    }
  };

  await call(ws, 'Target.setDiscoverTargets', { discover: true });
  await call(ws, 'Target.setAutoAttach', { autoAttach: true, waitForDebuggerOnStart: false, flatten: true });

  const loaded = await call(ws, 'Extensions.loadUnpacked', { path: root });
  if (loaded.error) {
    const why = loaded.error.message || '';
    if (/not found|wasn't found|Extensions\.loadUnpacked/i.test(why)) {
      console.log(`Skipped: this Chrome cannot load an extension over the DevTools protocol.\n  ${why}`);
      return true;
    }
    console.error(`Chrome refused to load the extension:\n  ${why}`);
    return false;
  }

  // The worker installs as soon as the extension loads. Give it a moment.
  for (let i = 0; i < 60 && !workerSession; i++) await sleep(100);

  if (!workerSession) {
    console.error('The service worker never started at all.');
    return false;
  }

  /*
   * Attaching to the worker is not the same as the worker having finished
   * evaluating, so this asks until the answer settles rather than once. Asking
   * once reported an empty namespace on a worker that was perfectly fine.
   */
  const question = 'JSON.stringify({' +
    ' util: typeof (self.FYG && self.FYG.util),' +
    ' state: typeof (self.FYG && self.FYG.state),' +
    ' idb: typeof (self.FYG && self.FYG.idb),' +
    ' routes: !!(self.FYG && self.FYG.state && self.FYG.state.ROUTES)' +
    '})';

  let present = {};
  let missing = ['util', 'state', 'idb'];
  for (let i = 0; i < 50 && (missing.length || !present.routes); i++) {
    const answer = await call(ws, 'Runtime.evaluate',
      { expression: question, returnByValue: true }, workerSession);
    const value = answer.result && answer.result.result && answer.result.result.value;
    present = value ? JSON.parse(value) : {};
    missing = ['util', 'state', 'idb'].filter((key) => present[key] !== 'object');
    if (!missing.length && present.routes) break;
    if (exceptions.length) break;
    await sleep(100);
  }

  ws.close();

  if (exceptions.length) {
    console.error('The service worker threw as it started:\n');
    exceptions.forEach((e) => console.error(`  ${e}`));
    return false;
  }

  if (missing.length || !present.routes) {
    console.error(`The service worker started but its imports did not take: FYG.${missing.join(', FYG.')}` +
      ` ${missing.length === 1 ? 'is' : 'are'} not there.`);
    return false;
  }

  console.log(`Extension loads in ${version.Browser}. The service worker starts, and` +
    ' FYG.util, FYG.state and FYG.idb are all there.');
  console.log(`  worker: ${workerUrl.replace(/^chrome-extension:\/\/[a-z]+/, 'chrome-extension://<id>')}`);
  return true;
}

let ok = false;
try {
  ok = await main();
} catch (err) {
  console.error(`Could not check the extension: ${err.message}`);
} finally {
  chrome.kill();
  await sleep(300);
  try { rmSync(profile, { recursive: true, force: true }); } catch { /* a temp dir Chrome still holds */ }
}

process.exit(ok ? 0 : 1);
