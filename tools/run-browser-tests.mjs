/*
 * Fygaro Catalog Automation
 * Runs the browser test pages in headless Chrome and reports the result.
 *
 * The pages are served over http and each one posts its outcome back to
 * /__result when it finishes. Two earlier approaches did not survive contact
 * with this suite, and both failures are worth recording:
 *
 *   Opening the pages from disk gave them no IndexedDB at all, so every call
 *   into the workbook and picture store quietly timed out and the storage layer
 *   looked tested while never having run.
 *
 *   Reading the outcome out of document.title with --dump-dom needed a virtual
 *   clock to let the asynchronous suites finish first, and Chrome's virtual
 *   clock fast forwards past IndexedDB completion callbacks whenever the page
 *   looks idle. Transactions started and never finished.
 *
 * So the pages run in real time and say when they are done.
 *
 *   node tools/run-browser-tests.mjs
 */
import { spawn } from 'node:child_process';
import { createReadStream, existsSync, mkdtempSync, rmSync, statSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { dirname, join, normalize, extname } from 'node:path';

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

/* Set while a page is running, so its posted result reaches the right waiter. */
let reportResult = null;

const server = createServer((req, res) => {
  const wanted = decodeURIComponent(req.url.split('?')[0]);

  if (req.method === 'POST' && wanted === '/__result') {
    let body = '';
    req.on('data', (piece) => { body += piece; });
    req.on('end', () => {
      res.writeHead(204);
      res.end();
      let parsed = { title: '', failures: ['the page posted something unreadable'] };
      try { parsed = JSON.parse(body); } catch { /* keep the fallback */ }
      if (reportResult) reportResult(parsed);
    });
    return;
  }

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

/** Runs one page in a real browser and waits for it to report back. */
function runPage(url) {
  // A fresh profile each time, so one page's IndexedDB cannot be seen by the
  // next and a stale database cannot make a test pass for the wrong reason.
  const profile = mkdtempSync(join(tmpdir(), 'fygaro-test-'));
  const child = spawn(chrome, [
    '--headless=new',
    '--disable-gpu',
    '--no-sandbox',
    `--user-data-dir=${profile}`,
    url
  ], { stdio: 'ignore' });

  return new Promise((resolve) => {
    const finish = (value) => {
      reportResult = null;
      clearTimeout(timer);
      child.kill('SIGKILL');
      try { rmSync(profile, { recursive: true, force: true }); } catch { /* windows holds it briefly */ }
      resolve(value);
    };
    const timer = setTimeout(() => {
      finish({ title: '', failures: ['the page did not report a result within 120 seconds'] });
    }, 120000);
    reportResult = finish;
  });
}

let failed = 0;

for (const page of PAGES) {
  const report = await runPage(`http://127.0.0.1:${port}/${page.file}`);
  const title = String(report.title || '');
  const total = title.split('/')[1] || '?';

  if (title.startsWith('PASS')) {
    console.log(`  ok  ${page.name.padEnd(10)} ${total} checks passed`);
    continue;
  }

  failed++;
  console.error(`FAIL  ${page.name.padEnd(10)} ${title || 'the page did not report a result'}`);
  for (const line of report.failures || []) console.error(`        ${line}`);
}

server.close();

if (failed) {
  console.error(`\n${failed} browser test page${failed === 1 ? '' : 's'} failed.`);
  process.exit(1);
}
console.log('\nAll browser tests passed.');
