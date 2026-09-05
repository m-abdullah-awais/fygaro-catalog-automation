/*
 * Fygaro Catalog Automation
 * Runs the browser test pages in headless Chrome and reports the result.
 *
 * Each page writes its outcome into document.title as "PASS 0/n" or "FAIL k/n",
 * so --dump-dom is enough and no browser driver is needed.
 *
 * The pages are served over http rather than opened from disk. Chrome gives a
 * file:// page no IndexedDB at all, and the extension keeps the workbook and its
 * images there, so on file:// every one of those calls would quietly time out
 * and the storage layer would look tested while never having run.
 *
 *   node tools/run-browser-tests.mjs
 */
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createReadStream, existsSync, statSync } from 'node:fs';
import { createServer } from 'node:http';
import { fileURLToPath } from 'node:url';
import { dirname, join, normalize, extname } from 'node:path';

const run = promisify(execFile);
const root = join(dirname(fileURLToPath(import.meta.url)), '..');

const CANDIDATES = [
  process.env.CHROME_PATH,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
  process.env.LOCALAPPDATA && `${process.env.LOCALAPPDATA}/Google/Chrome/Application/chrome.exe`,
  '/usr/bin/google-chrome',
  '/usr/bin/chromium',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
].filter(Boolean);

const chrome = CANDIDATES.find((p) => existsSync(p));
if (!chrome) {
  console.error('Chrome was not found. Set CHROME_PATH to the browser executable and try again.');
  process.exit(2);
}

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.png': 'image/png'
};

const server = createServer((req, res) => {
  const wanted = decodeURIComponent(req.url.split('?')[0]);
  const path = normalize(join(root, wanted));
  // Nothing outside the project is servable, however the url is written.
  if (!path.startsWith(normalize(root))) {
    res.writeHead(403);
    res.end('forbidden');
    return;
  }
  try {
    const info = statSync(path);
    if (!info.isFile()) throw new Error('not a file');
    res.writeHead(200, {
      'Content-Type': TYPES[extname(path)] || 'application/octet-stream',
      'Content-Length': info.size
    });
    createReadStream(path).pipe(res);
  } catch {
    res.writeHead(404);
    res.end('not found');
  }
});

await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const port = server.address().port;

const PAGES = [
  { file: 'tests/selectors.test.html', name: 'locators' },
  { file: 'tests/xlsx.test.html', name: 'workbook' },
  { file: 'tests/integration.test.html', name: 'end to end' }
];

let failed = 0;

for (const page of PAGES) {
  const url = `http://127.0.0.1:${port}/${page.file}`;
  let dom = '';
  try {
    // Async on purpose. The static server lives in this process, so a
    // synchronous spawn would block the event loop and Chrome would wait for a
    // reply that could never be sent.
    const result = await run(chrome, [
      '--headless=new',
      '--disable-gpu',
      '--no-sandbox',
      // Virtual time also advances Date.now(), so this has to cover every wait
      // the suites perform, not just their real elapsed time.
      '--virtual-time-budget=90000',
      '--dump-dom',
      url
    ], {
      encoding: 'utf8',
      maxBuffer: 64 * 1024 * 1024,
      // A safety net: never let a stuck browser hang the whole test run.
      timeout: 120000,
      killSignal: 'SIGKILL'
    });
    dom = result.stdout;
  } catch (err) {
    console.error(`${page.name}: Chrome could not run the page. ${err.message}`);
    failed++;
    continue;
  }

  const title = (/<title>(.*?)<\/title>/s.exec(dom) || [, ''])[1].trim();
  const passed = title.startsWith('PASS');
  const total = title.split('/')[1] || '?';

  if (passed) {
    console.log(`  ok  ${page.name.padEnd(10)} ${total} checks passed`);
    continue;
  }

  failed++;
  console.error(`FAIL  ${page.name.padEnd(10)} ${title || 'the page did not report a result'}`);
  for (const [, body] of dom.matchAll(/<li class="bad">(.*?)<\/li>/gs)) {
    const text = body.replace(/<div class="detail">/, ' -- ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    console.error(`        ${text}`);
  }
}

server.close();

if (failed) {
  console.error(`\n${failed} browser test page${failed === 1 ? '' : 's'} failed.`);
  process.exit(1);
}
console.log('\nAll browser tests passed.');
