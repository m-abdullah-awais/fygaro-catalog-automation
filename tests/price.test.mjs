/*
 * Fygaro Catalog Automation
 * Price parser tests. Zero dependencies, run with:  node --test tests/
 *
 * The fixture holds all 699 "Precio Total" values taken straight out of the real
 * workbook, with expected results produced by a separate reference
 * implementation, so this is a real cross check and not a snapshot of itself.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));

// price.js is a classic script, so it is evaluated rather than imported.
new Function(readFileSync(join(here, '..', 'src', 'shared', 'price.js'), 'utf8'))();
const price = globalThis.FYG.price;

const cases = JSON.parse(readFileSync(join(here, 'fixtures', 'prices.json'), 'utf8'));

test('fixture covers the whole catalog', () => {
  assert.equal(cases.length, 699);
});

test('every real catalog price parses to the expected number', () => {
  const failures = [];
  for (const c of cases) {
    const got = price.parse(c.raw);
    if (!got.ok || got.value !== c.value || got.text !== c.text) {
      failures.push(`row ${c.row}: ${JSON.stringify(c.raw)} expected ${c.value} got ${got.ok ? got.value : 'FAIL:' + got.reason}`);
    }
  }
  assert.deepEqual(failures, [], failures.slice(0, 10).join('\n'));
});

test('comma decimal', () => {
  assert.equal(price.parse('B/.625,00').value, 625);
  assert.equal(price.parse('B/.25,00').value, 25);
  assert.equal(price.parse('B/.1.312,50').value, 1312.5);
  assert.equal(price.parse('B/.2.875,00').value, 2875);
});

test('dot decimal', () => {
  assert.equal(price.parse('B/.100.00').value, 100);
  assert.equal(price.parse('B/.80.00').value, 80);
  assert.equal(price.parse('B/.96.00').value, 96);
});

test('bare thousands group reads as thousands, not as a fraction', () => {
  assert.equal(price.parse('B/.1.000').value, 1000);
  assert.equal(price.parse('1.234.567').value, 1234567);
  assert.equal(price.parse('1,234,567').value, 1234567);
});

test('when both separators appear, the last one is the decimal point', () => {
  assert.equal(price.parse('1.125,00').value, 1125);
  assert.equal(price.parse('1,299.99').value, 1299.99);
  assert.equal(price.parse('1.234.567,89').value, 1234567.89);
  assert.equal(price.parse('1,234,567.89').value, 1234567.89);
});

test('plain numbers and numeric input', () => {
  assert.equal(price.parse('150').value, 150);
  assert.equal(price.parse(123.45).value, 123.45);
  assert.equal(price.parse('$ 1,299.99').value, 1299.99);
});

test('the text handed to the Fygaro price field is always dot decimal with two places', () => {
  assert.equal(price.parse('B/.1.125,00').text, '1125.00');
  assert.equal(price.parse('B/.100.00').text, '100.00');
  assert.equal(price.parse('B/.1.312,50').text, '1312.50');
});

test('bad input is rejected rather than guessed', () => {
  for (const bad of ['', '   ', 'B/.', 'gratis', null, undefined]) {
    const r = price.parse(bad);
    assert.equal(r.ok, false, `expected ${JSON.stringify(bad)} to be rejected`);
    assert.ok(r.reason, 'a reason should be given');
  }
});
