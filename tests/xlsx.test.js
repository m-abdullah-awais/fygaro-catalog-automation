/*
 * Fygaro Catalog Automation
 * Browser side workbook tests, driven by tests/xlsx.test.html.
 *
 * These cover the DOMParser based reader and the row building that the side
 * panel does, which is the part node cannot exercise.
 */
(function () {
  'use strict';

  var U = FYG.util;
  var S = FYG.state;
  var X = FYG.xlsx;

  var WORKBOOK_PATH = '../docs/Cat%C3%A1logo%20de%20Productos%20y%20Servicios%20Fygaro.xlsx';
  var SHEET = 'Logros ';

  var results = [];

  function assert(condition, message) {
    if (!condition) throw new Error(message);
  }

  function check(name, fn) {
    return Promise.resolve()
      .then(fn)
      .then(function (detail) { results.push({ ok: true, name: name, detail: detail || '' }); })
      .catch(function (err) {
        results.push({ ok: false, name: name, detail: String(err && err.message ? err.message : err) });
      });
  }

  /** Mirrors the row building in src/sidepanel/sidepanel.js. */
  function buildRows(sheetData, header, mapping) {
    var rows = [];
    sheetData.rows.forEach(function (r) {
      if (r.r <= header.headerRow) return;
      var name = U.normText(mapping.name ? r.cells[mapping.name] : '');
      var code = U.normText(mapping.code ? r.cells[mapping.code] : '');
      var priceRaw = U.normText(mapping.price ? r.cells[mapping.price] : '');
      var link = U.normText(mapping.link ? r.cells[mapping.link] : '');
      if (!name && !code) return;
      var parsed = FYG.price.parse(priceRaw);
      rows.push({
        sheetRow: r.r, name: name, code: code, priceRaw: priceRaw,
        priceText: parsed.ok ? parsed.text : '', link: link,
        blocked: !link && (!name || !code || !parsed.ok)
      });
    });
    return rows;
  }

  function run(bytes) {
    var wb = null;
    var sheetData = null;
    var header = null;
    var mapping = null;

    return check('loads the workbook', function () {
      return X.load(bytes).then(function (loaded) {
        wb = loaded;
        assert(wb.sheets.length === 6, 'expected 6 sheets, saw ' + wb.sheets.length);
        assert(wb.sharedStrings.length > 1000, 'shared strings look empty');
        return wb.sheets.length + ' sheets, ' + wb.sharedStrings.length + ' shared strings';
      });
    })
    .then(function () {
      return check('finds "Logros " despite the trailing space in its name', function () {
        assert(wb.sheetNames.indexOf('Logros ') !== -1, 'the raw sheet name should still carry its space');
        assert(wb.findSheet('Logros ').name === SHEET, 'exact match failed');
        assert(wb.findSheet('Logros').name === SHEET, 'trimmed match failed');
        assert(wb.findSheet('logros').name === SHEET, 'case insensitive match failed');
        assert(wb.findSheet('El Consultorio').name === 'El Consultorio', 'a different sheet resolved wrongly');
        return 'exact, trimmed and case insensitive all resolve';
      });
    })
    .then(function () {
      return check('reads the sheet and its header', function () {
        sheetData = wb.readSheet(SHEET);
        header = X.readHeader(sheetData);
        assert(header.headerRow === 1, 'header row is ' + header.headerRow);
        var labels = header.labels.map(function (l) { return l.col + '=' + l.label; }).join(' ');
        assert(labels.indexOf('C=Código') !== -1, 'Código not found: ' + labels);
        assert(labels.indexOf('D=Servicios') !== -1, 'Servicios not found: ' + labels);
        assert(labels.indexOf('E=Precio Total') !== -1, 'Precio Total not found: ' + labels);
        assert(labels.indexOf('H=Link') !== -1, 'Link not found: ' + labels);
        return labels;
      });
    })
    .then(function () {
      return check('maps the columns the automation needs', function () {
        mapping = {
          name: X.findColumn(header, S.HEADERS.name),
          code: X.findColumn(header, S.HEADERS.code),
          price: X.findColumn(header, S.HEADERS.price),
          link: X.findColumn(header, S.HEADERS.link)
        };
        assert(mapping.name === 'D', 'Servicios mapped to ' + mapping.name);
        assert(mapping.code === 'C', 'Código mapped to ' + mapping.code);
        assert(mapping.price === 'E', 'Precio Total mapped to ' + mapping.price);
        assert(mapping.link === 'H', 'Link mapped to ' + mapping.link);
        return 'Servicios=D Código=C Precio=E Link=H';
      });
    })
    .then(function () {
      return check('builds all 699 catalog rows', function () {
        var rows = buildRows(sheetData, header, mapping);
        assert(rows.length === 699, 'built ' + rows.length + ' rows, expected 699');
        assert(rows[0].sheetRow === 2, 'first row is sheet row ' + rows[0].sheetRow);
        assert(rows[rows.length - 1].sheetRow === 700, 'last row is sheet row ' + rows[rows.length - 1].sheetRow);
        return rows.length + ' rows, sheet rows 2 to 700';
      });
    })
    .then(function () {
      return check('accented text survives the shared string table', function () {
        var rows = buildRows(sheetData, header, mapping);
        var row = rows[0];
        assert(row.code === 'CT-ING-PRE-EQ-01', 'code is "' + row.code + '"');
        assert(row.name === 'Cita de Ingreso Presencial con Psicóloga Clínica del Equipo',
          'name is "' + row.name + '"');
        assert(row.priceRaw === 'B/.100.00', 'price is "' + row.priceRaw + '"');
        assert(row.priceText === '100.00', 'parsed price is "' + row.priceText + '"');
        return row.name;
      });
    })
    .then(function () {
      return check('every row has a code, a name and a readable price', function () {
        var rows = buildRows(sheetData, header, mapping);
        var blocked = rows.filter(function (r) { return r.blocked; });
        assert(blocked.length === 0, blocked.length + ' rows cannot be processed, first is row ' +
          (blocked[0] && blocked[0].sheetRow));
        return 'no blocked rows';
      });
    })
    .then(function () {
      return check('the one row that already has a link is detected', function () {
        var rows = buildRows(sheetData, header, mapping);
        var withLink = rows.filter(function (r) { return r.link; });
        assert(withLink.length === 1, withLink.length + ' rows carry a link, expected 1');
        assert(withLink[0].sheetRow === 18, 'the link is on row ' + withLink[0].sheetRow);
        assert(withLink[0].link.indexOf('https://www.fygaro.com/es/pb/') === 0, 'unexpected link value');
        return 'row 18 will be skipped';
      });
    })
    .then(function () {
      return check('both price conventions in the sheet parse correctly', function () {
        var rows = buildRows(sheetData, header, mapping);
        var byRaw = {};
        rows.forEach(function (r) { byRaw[r.priceRaw] = r.priceText; });
        var expected = { 'B/.100.00': '100.00', 'B/.625,00': '625.00', 'B/.1.125,00': '1125.00' };
        Object.keys(expected).forEach(function (raw) {
          assert(byRaw[raw] === expected[raw], raw + ' parsed to ' + byRaw[raw] + ', expected ' + expected[raw]);
        });
        return Object.keys(byRaw).length + ' distinct price strings, all parsed';
      });
    })
    .then(function () {
      return check('patching column H survives a full write and re read', function () {
        var updates = [
          { row: 2, value: 'https://www.fygaro.com/en/pb/11111111-1111-1111-1111-111111111111/' },
          { row: 700, value: 'https://www.fygaro.com/en/pb/22222222-2222-2222-2222-222222222222/' }
        ];
        return X.writeColumn(wb, SHEET, 'H', updates)
          .then(function (blob) { return blob.arrayBuffer(); })
          .then(function (buffer) { return X.load(new Uint8Array(buffer)); })
          .then(function (again) {
            var data = again.readSheet(SHEET);
            var byRow = {};
            data.rows.forEach(function (r) { byRow[r.r] = r.cells; });

            assert(byRow[2].H === updates[0].value, 'row 2 link is "' + byRow[2].H + '"');
            assert(byRow[700].H === updates[1].value, 'row 700 link is "' + byRow[700].H + '"');
            // The pre existing link and the neighbouring data must be untouched.
            assert(byRow[18].H.indexOf('/es/pb/5c03f135') !== -1, 'row 18 was disturbed');
            assert(byRow[2].C === 'CT-ING-PRE-EQ-01', 'row 2 code changed');
            assert(byRow[2].D === 'Cita de Ingreso Presencial con Psicóloga Clínica del Equipo', 'row 2 name changed');
            assert(byRow[2].E === 'B/.100.00', 'row 2 price changed');
            assert(again.sheets.length === 6, 'sheets were lost: ' + again.sheets.length);
            assert(again.readSheet('El Consultorio').rows.length > 0, 'another sheet lost its rows');
            return 'links written, other cells and all 6 sheets intact';
          });
      });
    })
    .then(function () {
      return check('the loaded workbook is not mutated by exporting', function () {
        var data = wb.readSheet(SHEET);
        var byRow = {};
        data.rows.forEach(function (r) { byRow[r.r] = r.cells; });
        assert(!byRow[2].H, 'row 2 in the original should still be empty, saw "' + byRow[2].H + '"');
        return 'exporting twice stays correct';
      });
    })
    .then(report);
  }

  function report() {
    var failed = results.filter(function (r) { return !r.ok; });
    var summary = document.getElementById('summary');
    summary.className = failed.length ? 'fail' : 'pass';
    summary.textContent = failed.length
      ? failed.length + ' of ' + results.length + ' checks FAILED'
      : 'All ' + results.length + ' checks passed';
    document.title = (failed.length ? 'FAIL ' + failed.length + '/' : 'PASS 0/') + results.length;

    var list = document.getElementById('results');
    list.innerHTML = '';
    results.forEach(function (r) {
      var li = document.createElement('li');
      li.className = r.ok ? 'ok' : 'bad';
      var tag = document.createElement('span');
      tag.className = 'tag';
      tag.textContent = r.ok ? 'PASS' : 'FAIL';
      var name = document.createElement('span');
      name.textContent = r.name;
      var detail = document.createElement('div');
      detail.className = 'detail';
      detail.textContent = r.detail;
      li.appendChild(tag);
      li.appendChild(name);
      li.appendChild(detail);
      list.appendChild(li);
    });
  }

  // Headless runs use --allow-file-access-from-files and fetch the workbook
  // directly. Opened by hand that is blocked, so fall back to a file picker.
  fetch(WORKBOOK_PATH)
    .then(function (response) {
      if (!response.ok) throw new Error('status ' + response.status);
      return response.arrayBuffer();
    })
    .then(function (buffer) { return run(new Uint8Array(buffer)); })
    .catch(function () {
      document.getElementById('picker').classList.remove('hidden');
      document.getElementById('summary').textContent =
        'Pick the workbook above to run the tests. Chrome blocks reading local files without ' +
        '--allow-file-access-from-files.';
      document.getElementById('file').addEventListener('change', function (e) {
        var file = e.target.files[0];
        if (!file) return;
        document.getElementById('summary').textContent = 'Running...';
        file.arrayBuffer().then(function (buffer) { return run(new Uint8Array(buffer)); });
      });
    });
})();
