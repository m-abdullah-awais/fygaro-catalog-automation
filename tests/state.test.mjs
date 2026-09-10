/*
 * Fygaro Catalog Automation
 * Run state tests.  node --test tests/state.test.mjs
 *
 * These exist because of a real bug: settings read back from storage replaced
 * the defaults wholesale instead of merging into them, so a setting added in a
 * later version arrived as undefined for anyone who had already used the
 * extension. The zoom silently did nothing as a result. The merge is now
 * checked against every key in DEFAULT_SETTINGS, so the next setting added
 * cannot repeat it.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');

let store = {};
globalThis.chrome = {
  storage: {
    local: {
      get(keys) {
        const wanted = typeof keys === 'string' ? [keys] : keys;
        const out = {};
        for (const k of wanted) if (k in store) out[k] = structuredClone(store[k]);
        return Promise.resolve(out);
      },
      set(items) {
        Object.assign(store, structuredClone(items));
        return Promise.resolve();
      },
      remove(keys) {
        for (const k of (typeof keys === 'string' ? [keys] : keys)) delete store[k];
        return Promise.resolve();
      }
    }
  }
};

new Function(readFileSync(join(root, 'src', 'shared', 'util.js'), 'utf8'))();
new Function(readFileSync(join(root, 'src', 'shared', 'state.js'), 'utf8'))();
const S = globalThis.FYG.state;
const U = globalThis.FYG.util;

test('an empty profile reads the documented defaults', async () => {
  store = {};
  const { run, rows, log } = await S.read();
  assert.equal(run.status, S.STATUS.IDLE);
  assert.equal(run.settings.zoomPercent, 67);
  assert.equal(run.settings.minDelayMs, 1200);
  assert.equal(run.settings.maxDelayMs, 3000);
  assert.deepEqual(rows, []);
  assert.deepEqual(log, []);
});

test('a run stored by an older version gains every new setting', async () => {
  // Exactly the shape saved before zoomPercent existed.
  store = {
    [S.KEY_RUN]: {
      status: 'paused',
      cursor: 41,
      settings: { minDelayMs: 800, maxDelayMs: 1500, stepTimeoutMs: 30000, maxAttempts: 3, dryRun: false }
    }
  };

  const { run } = await S.read();

  assert.equal(run.settings.zoomPercent, 67, 'the new setting must fall back to its default');
  for (const key of Object.keys(S.DEFAULT_SETTINGS)) {
    assert.notEqual(run.settings[key], undefined, `settings.${key} came back undefined`);
  }

  // What the user had chosen must survive the merge.
  assert.equal(run.settings.minDelayMs, 800);
  assert.equal(run.settings.stepTimeoutMs, 30000);
  assert.equal(run.settings.maxAttempts, 3);

  // And the rest of the run is still whatever was stored.
  assert.equal(run.status, 'paused');
  assert.equal(run.cursor, 41);
});

test('a zoom of 100 saved by an older build is repaired, not honoured', async () => {
  // Builds before the settings version existed could persist 100 because a blank
  // field fell back to a literal. The user never chose it, and honouring it would
  // zoom the page in rather than out.
  store = {
    [S.KEY_RUN]: {
      settings: { minDelayMs: 900, maxDelayMs: 2000, stepTimeoutMs: 20000, maxAttempts: 2, zoomPercent: 100, dryRun: false }
    }
  };

  const { run } = await S.read();
  assert.equal(run.settings.zoomPercent, 67, 'the stray 100 should have been repaired');
  assert.equal(run.settings.settingsVersion, S.SETTINGS_VERSION);
  assert.equal(run.settings.minDelayMs, 900, 'unrelated settings must not be touched');
});

test('a zoom the user deliberately chose is kept once the settings are versioned', async () => {
  store = {
    [S.KEY_RUN]: {
      settings: Object.assign({}, S.DEFAULT_SETTINGS, { zoomPercent: 100, settingsVersion: S.SETTINGS_VERSION })
    }
  };

  const { run } = await S.read();
  assert.equal(run.settings.zoomPercent, 100, 'a deliberate choice must be respected');
});

test('the repair runs once, then leaves the value alone', async () => {
  store = { [S.KEY_RUN]: { settings: { zoomPercent: 100 } } };

  const first = (await S.read()).run;
  assert.equal(first.settings.zoomPercent, 67);

  // Simulate the repaired settings being written back, then read again.
  store[S.KEY_RUN] = { settings: Object.assign({}, first.settings, { zoomPercent: 120 }) };
  const second = (await S.read()).run;
  assert.equal(second.settings.zoomPercent, 120, 'a later choice must survive');
});

test('a zoom of 67 percent actually resolves to a real zoom factor', async () => {
  store = {};
  const { run } = await S.read();
  const factor = (run.settings.zoomPercent || 100) / 100;
  assert.equal(factor, 0.67, 'this is the value handed to chrome.tabs.setZoom');
  assert.notEqual(factor, 1, 'a factor of 1 would be a silent no-op');
});

test('fields added to the run itself also fall back to their defaults', async () => {
  store = { [S.KEY_RUN]: { status: 'running', cursor: 3 } };
  const { run } = await S.read();
  for (const key of Object.keys(S.defaultRun())) {
    assert.ok(key in run, `run.${key} went missing`);
  }
  assert.equal(run.originalZoom, null);
  assert.equal(run.waitingSince, null);
  assert.deepEqual(run.durations, []);
});

test('stats are filled in even when a stored run predates a counter', async () => {
  store = { [S.KEY_RUN]: { stats: { total: 10, done: 4 } } };
  const { run } = await S.read();
  assert.equal(run.stats.total, 10);
  assert.equal(run.stats.done, 4);
  assert.equal(run.stats.failed, 0);
  assert.equal(run.stats.skipped, 0);
});

test('recount counts each row once and excludes failed rows from the work left', () => {
  const rows = [
    { status: S.ROW.DONE }, { status: S.ROW.DONE },
    { status: S.ROW.SKIPPED },
    { status: S.ROW.FAILED },
    { status: S.ROW.PENDING }, { status: S.ROW.PENDING },
    { status: S.ROW.ACTIVE }
  ];
  assert.deepEqual(S.recount(rows), { total: 7, toProcess: 3, done: 2, failed: 1, skipped: 1 });
});

test('routeOf names every page the automation drives, and nothing else', () => {
  assert.equal(S.routeOf('/en/app/dashboard/'), 'dashboard');
  assert.equal(S.routeOf('/es/app/dashboard/'), 'dashboard');
  assert.equal(S.routeOf('/en/app/products/'), 'productList');
  assert.equal(S.routeOf('/en/app/products/add/'), 'productAdd');
  assert.equal(S.routeOf('/en/app/products/c8faacb1-fafc-4672-a50d-ee8edc325bd0/permalink/product/'), 'productDetail');
  assert.equal(S.routeOf('/en/app/payment-buttons/payments/payment-buttons/add/'), 'linkAdd');
  assert.equal(S.routeOf('/en/app/payment-buttons/payments/75fed96c-a3d5-4f52-a494-8fb0498cc68e/payment-buttons/permalink/'), 'linkDone');
  assert.equal(S.routeOf('/en/app/clients/'), '');
  assert.equal(S.routeOf('/en/pb/75fed96c-a3d5-4f52-a494-8fb0498cc68e/'), '');
});

test('every step declares the page it is allowed to act on', () => {
  for (const step of S.STEP_ORDER) {
    const route = S.STEP_ROUTE[step];
    assert.ok(route, `${step} has no route`);
    assert.ok(route in S.ROUTES, `${step} points at unknown route ${route}`);
    assert.ok(S.STEP_LABEL[step], `${step} has no label for the interface`);
  }
  assert.equal(S.STEP_ORDER.length, 7);
});

/*
 * Recognising a link the sheet already carries lives beside the run state
 * because S.SKIP.HAD_LINK is what it decides. A row wrongly called linked is
 * never created, and a row wrongly called unlinked is created a second time and
 * refused for a duplicate code tens of seconds later, so both directions matter.
 */
test('a Fygaro payment link is recognised whatever language prefix it carries', () => {
  const links = [
    'https://www.fygaro.com/en/pb/65df9668-0a2e-4c0b-a8f7-0789e47be346/',
    'https://www.fygaro.com/es/pb/65df9668-0a2e-4c0b-a8f7-0789e47be346/',
    'https://www.fygaro.com/en/pb/65df9668-0a2e-4c0b-a8f7-0789e47be346',
    'http://www.fygaro.com/en/pb/65DF9668-0A2E-4C0B-A8F7-0789E47BE346/'
  ];
  for (const link of links) {
    assert.ok(U.isFygaroLink(link), `should be a link: ${link}`);
    assert.equal(U.fygaroLink(link), link);
  }
});

test('text that is not a Fygaro link never counts as one', () => {
  const notLinks = [
    '',
    '   ',
    null,
    undefined,
    'pending',
    'N/A',
    '2026-09-10',
    'B/.625,00',
    // Both of these are real values this extension writes into the Nota column
    // one place over. A mapping slip of a single column used to retire every
    // row that carried one.
    'Fuera del rango elegido (1020 a 1030).',
    'Ya existe en Fygaro. Nothing was created and no link was captured.',
    // Fygaro pages that are not payment buttons.
    'https://www.fygaro.com/en/app/dashboard/',
    'https://www.fygaro.com/en/app/products/65df9668-0a2e-4c0b-a8f7-0789e47be346/permalink/product/',
    // Right shape, wrong length in the uuid slot.
    'https://www.fygaro.com/en/pb/65df9668/'
  ];
  for (const value of notLinks) {
    assert.ok(!U.isFygaroLink(value), `should not be a link: ${JSON.stringify(value)}`);
    assert.equal(U.fygaroLink(value), '', `should extract nothing from ${JSON.stringify(value)}`);
  }
});

test('a link is pulled out of the text around it, and comes back clean', () => {
  const link = 'https://www.fygaro.com/en/pb/65df9668-0a2e-4c0b-a8f7-0789e47be346/';
  assert.equal(U.fygaroLink('Link: ' + link), link);
  assert.equal(U.fygaroLink(link + ' (sent to the client)'), link);
  assert.equal(U.fygaroLink('  ' + link + '  '), link);
  // Extraction is what decides the row is done, so it must agree with the
  // stricter whole value test on anything the whole value test accepts.
  assert.ok(U.isFygaroLink(U.fygaroLink('Link: ' + link)));
});

test('the skip reasons the panel tells apart are all still declared', () => {
  assert.equal(S.SKIP.HAD_LINK, 'hadLink');
  assert.equal(S.SKIP.EXISTS, 'exists');
  assert.equal(S.SKIP.OUT_OF_RANGE, 'range');
});
