/*
 * Fygaro Catalog Automation
 * Runs the browser test pages in headless Chrome and reports the result.
 *
 * Each page writes its outcome into document.title as "PASS 0/n" or "FAIL k/n",
 * so --dump-dom is enough and no browser driver is needed.
 *
 *   node tools/run-browser-tests.mjs
 */
import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { dirname, join } from 'node:path';

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

const PAGES = [
  { file: 'tests/selectors.test.html', name: 'locators' },
  { file: 'tests/xlsx.test.html', name: 'workbook' },
  { file: 'tests/integration.test.html', name: 'end to end' }
];

let failed = 0;

for (const page of PAGES) {
  const url = pathToFileURL(join(root, page.file)).href;
  let dom = '';
  try {
    dom = execFileSync(chrome, [
      '--headless=new',
      '--disable-gpu',
      '--no-sandbox',
      // The workbook test reads the catalog straight off disk.
      '--allow-file-access-from-files',
      '--virtual-time-budget=20000',
      '--dump-dom',
      url
    ], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
      maxBuffer: 64 * 1024 * 1024,
      // A safety net: never let a stuck browser hang the whole test run.
      timeout: 120000,
      killSignal: 'SIGKILL'
    });
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

if (failed) {
  console.error(`\n${failed} browser test page${failed === 1 ? '' : 's'} failed.`);
  process.exit(1);
}
console.log('\nAll browser tests passed.');
