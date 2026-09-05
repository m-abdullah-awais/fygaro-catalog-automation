/*
 * Fygaro Catalog Automation
 * ZIP and worksheet patching tests, run against the real 2073 row catalog.
 * Zero dependencies:  node --test tests/xlsx.test.mjs
 *
 * The DOMParser based reader in xlsx-read.js cannot run under node, so that half
 * is covered by tests/selectors.test.html and tests/xlsx.test.html in Chrome.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');
const load = (p) => new Function(readFileSync(join(root, p), 'utf8'))();

load('src/shared/util.js');
load('src/shared/zip.js');
load('src/shared/xlsx-read.js');
load('src/shared/xlsx-write.js');
const { zip, xlsx } = globalThis.FYG;

const WORKBOOK = 'docs/Catálogo de Productos y Servicios Fygaro.xlsx';
const SHEET_PART = 'xl/worksheets/sheet1.xml';
const original = new Uint8Array(readFileSync(join(root, WORKBOOK)));

const cellIn = (xml, ref) =>
  new RegExp(`<c\\b[^>]*r="${ref}"[^>]*?(?:/>|>[\\s\\S]*?</c>)`).exec(xml)?.[0] ?? null;

test('reads every entry of the real workbook', async () => {
  const arch = await zip.read(original);
  assert.equal(arch.order.length, 66);
  assert.ok(arch.files.has('xl/workbook.xml'));
  assert.ok(arch.files.has('xl/sharedStrings.xml'));
  assert.ok(arch.files.has(SHEET_PART));
  assert.equal(arch.order.length, arch.files.size);
});

test('a rewrite with no edits reproduces every part byte for byte', async () => {
  const arch = await zip.read(original);
  const blob = await zip.write(arch.order, arch.files);
  const again = await zip.read(new Uint8Array(await blob.arrayBuffer()));

  assert.deepEqual(again.order, arch.order, 'entry order must be preserved');
  for (const name of arch.order) {
    assert.deepEqual(again.files.get(name), arch.files.get(name), `content changed for ${name}`);
  }
});

test('crc32 matches the well known check value', () => {
  assert.equal(zip.crc32(new TextEncoder().encode('123456789')), 0xcbf43926);
  assert.equal(zip.crc32(new Uint8Array(0)), 0);
});

test('patching column H keeps styles, other cells, and every other part', async () => {
  const arch = await zip.read(original);
  const xml = new TextDecoder().decode(arch.files.get(SHEET_PART));

  // Rows 2, 3 and 700 exist with an empty styled H cell. Row 3000 is past the
  // end of the sheet, which exercises the row creation path.
  const updates = [
    { row: 2, value: 'https://www.fygaro.com/en/pb/aaaaaaaa-0000-0000-0000-000000000002/' },
    { row: 3, value: 'https://www.fygaro.com/en/pb/aaaaaaaa-0000-0000-0000-000000000003/' },
    { row: 700, value: 'https://www.fygaro.com/en/pb/aaaaaaaa-0000-0000-0000-000000000700/' },
    { row: 3000, value: 'https://www.fygaro.com/en/pb/aaaaaaaa-0000-0000-0000-000000003000/' }
  ];

  assert.equal(cellIn(xml, 'H2'), '<c r="H2" s="9"/>', 'precondition: H2 starts empty and styled');
  assert.equal(cellIn(xml, 'H3000'), null, 'precondition: row 3000 does not exist');

  const patched = xlsx.patchSheetXml(xml, 'H', updates);

  assert.match(cellIn(patched, 'H2'), /^<c r="H2" s="9" t="inlineStr">/, 'style s="9" must survive');
  assert.match(cellIn(patched, 'H3'), /^<c r="H3" s="9" t="inlineStr">/, 'style s="9" must survive');
  assert.match(cellIn(patched, 'H700'), /^<c r="H700" s="9" t="inlineStr">/, 'style s="9" must survive');
  assert.ok(cellIn(patched, 'H3000'), 'a missing row must be created');
  for (const u of updates) {
    assert.ok(patched.includes(u.value), `link for row ${u.row} must be present`);
  }

  // Rows that were not targeted keep their cell exactly as it was.
  assert.equal(cellIn(patched, 'H18'), cellIn(xml, 'H18'), 'untargeted rows must be untouched');

  // Neighbouring columns are left alone.
  for (const ref of ['C2', 'D2', 'E2', 'G18', 'C700']) {
    assert.equal(cellIn(patched, ref), cellIn(xml, ref), `${ref} must be untouched`);
  }

  const files = new Map(arch.files);
  files.set(SHEET_PART, new TextEncoder().encode(patched));
  const blob = await zip.write(arch.order, files);
  const rebuilt = await zip.read(new Uint8Array(await blob.arrayBuffer()));

  assert.deepEqual(rebuilt.order, arch.order);
  const changed = arch.order.filter(
    (n) => Buffer.compare(Buffer.from(rebuilt.files.get(n)), Buffer.from(arch.files.get(n))) !== 0
  );
  assert.deepEqual(changed, [SHEET_PART], 'exactly one part may differ');
});

test('the loaded workbook is never mutated, so repeated exports stay correct', async () => {
  const arch = await zip.read(original);
  const before = arch.files.get(SHEET_PART);
  const xml = new TextDecoder().decode(before);
  xlsx.patchSheetXml(xml, 'H', [{ row: 5, value: 'https://example.test/x' }]);
  assert.deepEqual(arch.files.get(SHEET_PART), before);
});

test('patching handles self closing rows, absent cells and cleared values', () => {
  const head = '<worksheet><sheetData>';
  const tail = '</sheetData></worksheet>';

  const selfClosingRow = head + '<row r="4"/>' + tail;
  assert.match(
    xlsx.patchSheetXml(selfClosingRow, 'H', [{ row: 4, value: 'x' }]),
    /<row r="4"><c r="H4" t="inlineStr"><is><t xml:space="preserve">x<\/t><\/is><\/c><\/row>/
  );

  // A new cell must land in column order, not simply be appended.
  const partialRow = head + '<row r="4"><c r="A4"/><c r="Z4"/></row>' + tail;
  const filled = xlsx.patchSheetXml(partialRow, 'H', [{ row: 4, value: 'x' }]);
  assert.ok(filled.indexOf('r="A4"') < filled.indexOf('r="H4"'), 'H must come after A');
  assert.ok(filled.indexOf('r="H4"') < filled.indexOf('r="Z4"'), 'H must come before Z');

  // A new row must land in row order.
  const rows = head + '<row r="1"><c r="A1"/></row><row r="9"><c r="A9"/></row>' + tail;
  const inserted = xlsx.patchSheetXml(rows, 'H', [{ row: 5, value: 'x' }]);
  assert.ok(inserted.indexOf('r="A1"') < inserted.indexOf('r="H5"'), 'row 5 must come after row 1');
  assert.ok(inserted.indexOf('r="H5"') < inserted.indexOf('r="A9"'), 'row 5 must come before row 9');

  // An empty value clears the cell but keeps its style.
  const populated = head + '<row r="4"><c r="H4" s="7" t="inlineStr"><is><t>old</t></is></c></row>' + tail;
  assert.equal(cellIn(xlsx.patchSheetXml(populated, 'H', [{ row: 4, value: '' }]), 'H4'), '<c r="H4" s="7"/>');
});

test('values that would break the XML are escaped', () => {
  const xml = '<worksheet><sheetData><row r="2"><c r="H2" s="7"/></row></sheetData></worksheet>';
  const out = xlsx.patchSheetXml(xml, 'H', [{ row: 2, value: 'a&b<c>"d\'e' }]);
  assert.ok(out.includes('a&amp;b&lt;c&gt;&quot;d&apos;e'));
  assert.ok(!out.includes('<c>'), 'raw angle brackets must not leak into the markup');
});

test('a column no row has a cell for is created on every row, in column order', () => {
  const rows = [1, 2, 3].map((r) =>
    `<row r="${r}"><c r="G${r}" s="4"/><c r="H${r}" s="9"/></row>`).join('');
  const xml = `<worksheet><sheetData>${rows}</sheetData></worksheet>`;

  const out = xlsx.patchSheetXml(xml, 'I', [
    { row: 1, value: 'Link' }, { row: 2, value: 'one' }, { row: 3, value: 'two' }
  ]);

  for (const r of [1, 2, 3]) {
    assert.ok(cellIn(out, `I${r}`), `I${r} must be created`);
    assert.ok(out.indexOf(`r="H${r}"`) < out.indexOf(`r="I${r}"`), `I${r} must follow H${r}`);
    assert.equal(cellIn(out, `G${r}`), `<c r="G${r}" s="4"/>`, `G${r} must be untouched`);
    assert.equal(cellIn(out, `H${r}`), `<c r="H${r}" s="9"/>`, `H${r} must be untouched`);
  }
});

test('a created cell inherits the style of the cell to its left', () => {
  // The header cell and the body cells carry different styles in the real
  // sheet, so a new column has to pick up whichever one it is sitting beside.
  const xml = '<worksheet><sheetData>' +
    '<row r="1"><c r="H1" s="6" t="s"><v>15</v></c></row>' +
    '<row r="2"><c r="H2" s="9"/></row>' +
    '</sheetData></worksheet>';

  const out = xlsx.patchSheetXml(xml, 'I', [{ row: 1, value: 'Link' }, { row: 2, value: 'x' }]);
  assert.match(cellIn(out, 'I1'), /^<c r="I1" s="6" t="inlineStr">/, 'I1 must inherit s="6"');
  assert.match(cellIn(out, 'I2'), /^<c r="I2" s="9" t="inlineStr">/, 'I2 must inherit s="9"');
});

test('a cell that already exists in the new column keeps its own style and neighbours', () => {
  // Row 1017 of the real catalog already carries empty I through X cells.
  const xml = '<worksheet><sheetData><row r="1017">' +
    '<c r="H1017" s="52"/><c r="I1017" s="65"/><c r="J1017" s="65"/>' +
    '</row></sheetData></worksheet>';

  const out = xlsx.patchSheetXml(xml, 'I', [{ row: 1017, value: 'link' }]);
  assert.match(cellIn(out, 'I1017'), /^<c r="I1017" s="65" t="inlineStr">/, 'its own style must win');
  assert.equal(cellIn(out, 'J1017'), '<c r="J1017" s="65"/>', 'the right neighbour must survive');
  assert.ok(out.indexOf('r="I1017"') < out.indexOf('r="J1017"'), 'column order must hold');
  assert.equal((out.match(/r="I1017"/g) || []).length, 1, 'the cell must not be duplicated');
});

test('writing a brand new column into a whole sheet is not quadratic', () => {
  // Locating each row with a fresh scan of the sheet made this 15 seconds on the
  // real catalog while the same batch against an existing column took 230 ms.
  const rows = [];
  for (let r = 1; r <= 2400; r++) {
    rows.push(`<row r="${r}"><c r="A${r}" s="1"/><c r="H${r}" s="9"/></row>`);
  }
  const xml = `<worksheet><sheetData>${rows.join('')}</sheetData></worksheet>`;
  const updates = [];
  for (let r = 1; r <= 2074; r++) updates.push({ row: r, value: 'https://example.test/pb/' + r + '/' });

  const started = Date.now();
  const out = xlsx.patchSheetXml(xml, 'I', updates);
  const elapsed = Date.now() - started;

  assert.equal((out.match(/r="I\d+"/g) || []).length, 2074, 'every update must land');
  assert.ok(out.includes('https://example.test/pb/2074/'), 'the last row must be written');
  assert.equal(cellIn(out, 'H9'), '<c r="H9" s="9"/>', 'untargeted columns must be untouched');
  assert.ok(elapsed < 1500, `took ${elapsed} ms, which means the row scan went quadratic again`);
});

test('two columns are written in one pass and only the worksheet part changes', async () => {
  const arch = await zip.read(original);
  const wb = { order: arch.order, files: arch.files, findSheet: () => ({ path: SHEET_PART }) };

  const blob = await xlsx.writeColumns(wb, 'Logros ', [
    { column: 'I', updates: [{ row: 1, value: 'Link' }, { row: 2, value: 'https://example.test/a/' }] },
    { column: 'J', updates: [{ row: 1, value: 'Nota' }, { row: 3, value: 'El precio es cero.' }] }
  ]);
  const rebuilt = await zip.read(new Uint8Array(await blob.arrayBuffer()));
  const xml = new TextDecoder().decode(rebuilt.files.get(SHEET_PART));

  assert.match(cellIn(xml, 'I1'), /Link/, 'the link header must be written');
  assert.match(cellIn(xml, 'I2'), /example\.test/, 'the link must be written');
  assert.match(cellIn(xml, 'J1'), /Nota/, 'the note header must be written');
  assert.match(cellIn(xml, 'J3'), /precio/, 'the note must be written');

  // The table, the drawing that anchors 1030 product images, and the package
  // metadata must all come through untouched, because widening the table or
  // disturbing the drawing is what would make Excel offer to repair the file.
  const changed = arch.order.filter(
    (n) => Buffer.compare(Buffer.from(rebuilt.files.get(n)), Buffer.from(arch.files.get(n))) !== 0
  );
  assert.deepEqual(changed, [SHEET_PART], 'only the worksheet may differ');

  const table = new TextDecoder().decode(rebuilt.files.get('xl/tables/table1.xml'));
  assert.ok(table.includes('ref="A1:H2408"'), 'the table range must still stop at H');
  assert.ok(table.includes('<tableColumns count="8">'), 'the table must still declare 8 columns');
});

test('the export file name is derived from the original', () => {
  const name = xlsx.exportName('Catálogo de Productos y Servicios Fygaro.xlsx');
  assert.match(name, /^Catálogo de Productos y Servicios Fygaro \(con Links\) \d{4}-\d{2}-\d{2}\.xlsx$/);
});

test('a non zip input fails with a clear message instead of hanging', async () => {
  await assert.rejects(
    () => zip.read(new TextEncoder().encode('this is not a zip file at all')),
    /not a zip file/i
  );
});
