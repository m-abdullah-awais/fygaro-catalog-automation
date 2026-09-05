/*
 * Fygaro Catalog Automation
 * Browser side workbook tests, driven by tests/xlsx.test.html.
 *
 * These cover the DOMParser based reader and the row building that the side
 * panel does, which is the part node cannot exercise.
 *
 * They run against tests/fixtures/catalog-sample.xlsx rather than the real
 * catalog, because headless Chrome does not finish fetching a 63.5 MB file
 * before it dumps the page. The fixture is cut from the real one by
 * tools/build-sample-workbook.py and keeps every structural feature that
 * matters, so this is a trim rather than a mock.
 */
(function () {
  'use strict';

  var U = FYG.util;
  var S = FYG.state;
  var X = FYG.xlsx;

  var WORKBOOK_PATH = 'fixtures/catalog-sample.xlsx';
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
        assert(labels.indexOf('H=Columna 1') !== -1, 'Columna 1 not found: ' + labels);
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
        assert(mapping.link === '', 'this sheet has no Link column, but one mapped to ' + mapping.link);
        return 'Servicios=D Código=C Precio=E, and no Link column';
      });
    })
    .then(function () {
      return check('builds every data row and no header row', function () {
        var rows = buildRows(sheetData, header, mapping);
        assert(rows.length === 30, 'built ' + rows.length + ' rows, expected 30');
        assert(rows[0].sheetRow === 2, 'first row is sheet row ' + rows[0].sheetRow);
        assert(rows[rows.length - 1].sheetRow === 31, 'last row is sheet row ' + rows[rows.length - 1].sheetRow);
        return rows.length + ' rows, sheet rows 2 to 31';
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
      return check('no row carries a link yet, so nothing is pre skipped', function () {
        var rows = buildRows(sheetData, header, mapping);
        var withLink = rows.filter(function (r) { return r.link; });
        assert(withLink.length === 0, withLink.length + ' rows carry a link, expected none');
        return 'every row is still to do';
      });
    })
    .then(function () {
      return check('both price conventions in the sheet parse correctly', function () {
        var rows = buildRows(sheetData, header, mapping);
        var byRaw = {};
        rows.forEach(function (r) { byRaw[r.priceRaw] = r.priceText; });
        var expected = { 'B/.100.00': '100.00', 'B/.80.00': '80.00', 'B/.120.00': '120.00' };
        Object.keys(expected).forEach(function (raw) {
          assert(byRaw[raw] === expected[raw], raw + ' parsed to ' + byRaw[raw] + ', expected ' + expected[raw]);
        });
        return Object.keys(byRaw).length + ' distinct price strings, all parsed';
      });
    })
    .then(function () {
      return check('writing a new column survives a full write and re read', function () {
        var updates = [
          { row: 2, value: 'https://www.fygaro.com/en/pb/11111111-1111-1111-1111-111111111111/' },
          { row: 31, value: 'https://www.fygaro.com/en/pb/22222222-2222-2222-2222-222222222222/' }
        ];
        return X.writeColumn(wb, SHEET, 'I', updates)
          .then(function (blob) { return blob.arrayBuffer(); })
          .then(function (buffer) { return X.load(new Uint8Array(buffer)); })
          .then(function (again) {
            var data = again.readSheet(SHEET);
            var byRow = {};
            data.rows.forEach(function (r) { byRow[r.r] = r.cells; });

            assert(byRow[2].I === updates[0].value, 'row 2 link is "' + byRow[2].I + '"');
            assert(byRow[31].I === updates[1].value, 'row 31 link is "' + byRow[31].I + '"');
            // Column H holds the product images, so it must come through empty.
            assert(!byRow[2].H, 'row 2 column H was written into: "' + byRow[2].H + '"');
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
        assert(!byRow[2].I, 'row 2 in the original should still have no link, saw "' + byRow[2].I + '"');
        assert(!byRow[2].H, 'row 2 column H should still be empty, saw "' + byRow[2].H + '"');
        return 'exporting twice stays correct';
      });
    })
    .then(function () {
      return check('reads the pictures anchored to each row', function () {
        var found = X.readImages(wb, SHEET);
        assert(found.byRow.size === 5, 'expected 5 rows with pictures, saw ' + found.byRow.size);
        assert(found.images.size === 3, 'expected 3 distinct pictures, saw ' + found.images.size);
        assert(found.byRow.get(2).length === 1, 'row 2 should carry one picture');
        assert(found.byRow.get(5).length === 2, 'row 5 should carry two pictures');
        assert(!found.byRow.has(7), 'a row with no picture should be absent, not empty');
        return found.byRow.size + ' rows, ' + found.images.size + ' distinct pictures';
      });
    })
    .then(function () {
      return check('one picture shared by several rows is stored once', function () {
        var found = X.readImages(wb, SHEET);
        var first = found.byRow.get(2)[0];
        assert(found.byRow.get(3)[0] === first, 'rows 2 and 3 should share a picture');
        assert(found.images.get(first), 'the shared picture is missing from the table');
        assert(found.images.get(first).bytes.length > 0, 'the picture has no bytes');
        return first + ' is used by rows 2 and 3';
      });
    })
    .then(function () {
      return check('the format is read from the bytes, not from the file name', function () {
        var found = X.readImages(wb, SHEET);
        var byName = {};
        found.images.forEach(function (img) { byName[img.name] = img; });
        assert(byName['image1.png'].type === 'image/png', 'image1.png is ' + byName['image1.png'].type);
        assert(byName['image2.jpg'].type === 'image/jpeg', 'image2.jpg is ' + byName['image2.jpg'].type);
        // Deliberately JPEG bytes behind a .png name. Trusting the extension
        // would upload a JPEG labelled as a PNG.
        assert(byName['image3.png'].type === 'image/jpeg',
          'image3.png should sniff as JPEG, saw ' + byName['image3.png'].type);
        return 'PNG, JPEG, and a JPEG wearing a .png name';
      });
    })
    .then(function () {
      return check('a picture stretched across cells belongs to the row it starts in', function () {
        var found = X.readImages(wb, SHEET);
        assert(found.byRow.has(6), 'the twoCellAnchor picture was lost');
        assert(found.byRow.get(6).length === 1, 'row 6 should carry one picture');
        return 'twoCellAnchor attributed to its from row';
      });
    })
    .then(function () {
      return check('a picture floating free of any row is reported, not guessed at', function () {
        var found = X.readImages(wb, SHEET);
        var loose = found.skipped.filter(function (s) { return s.reason === 'unanchored'; });
        assert(loose.length === 1, 'expected 1 unanchored picture, saw ' + loose.length);
        var rows = [];
        found.byRow.forEach(function (ids, row) { rows.push(row); });
        assert(rows.indexOf(null) === -1, 'a null row leaked into the map');
        return 'absoluteAnchor skipped with a reason';
      });
    })
    .then(function () {
      return check('every anchored picture sits in the images column', function () {
        var anchors = X.parseDrawingAnchors(
          new TextDecoder('utf-8').decode(wb.files.get('xl/drawings/drawing1.xml')));
        var placed = anchors.filter(function (a) { return a.row !== null; });
        placed.forEach(function (a) {
          assert(a.col === 7, 'a picture is anchored at column index ' + a.col + ', expected 7 (H)');
        });
        return placed.length + ' pictures, all in column H';
      });
    })
    .then(function () {
      return check('a sheet nobody put pictures on yields an empty map, not an error', function () {
        // El Consultorio has a drawing relationship pointing at a part with no
        // anchors and no relationships file of its own.
        var found = X.readImages(wb, 'El Consultorio');
        assert(found.byRow.size === 0, 'expected no pictures, saw ' + found.byRow.size);
        assert(found.images.size === 0, 'expected no image table');
        return 'empty and quiet';
      });
    })
    .then(function () {
      return check('reading the pictures twice returns the same work', function () {
        assert(X.readImages(wb, SHEET) === X.readImages(wb, SHEET), 'the result should be cached');
        return 'cached per sheet';
      });
    })
    .then(function () {
      return check('relationship targets resolve relative to the part that declared them', function () {
        assert(X.resolvePart('xl/drawings/drawing1.xml', '../media/image1.png') === 'xl/media/image1.png',
          'parent relative failed');
        assert(X.resolvePart('xl/drawings/drawing1.xml', './image1.png') === 'xl/drawings/image1.png',
          'same directory failed');
        assert(X.resolvePart('xl/drawings/drawing1.xml', 'image1.png') === 'xl/drawings/image1.png',
          'bare name failed');
        assert(X.resolvePart('xl/drawings/drawing1.xml', '/xl/media/image1.png') === 'xl/media/image1.png',
          'package absolute failed');
        assert(X.relsPathFor('xl/worksheets/sheet1.xml') === 'xl/worksheets/_rels/sheet1.xml.rels',
          'rels path failed');
        return 'relative, absolute and bare targets all resolve';
      });
    })
    .then(function () {
      return check('a sheet with no Link column gets one proposed past the images', function () {
        assert(X.findColumn(header, S.HEADERS.link) === '', 'this sheet should have no Link column');
        assert(header.byLabel[U.foldText('Columna 1')] === 'H', 'the images column should be H');
        var next = X.nextFreeColumn(header, sheetData);
        assert(next === 'I', 'proposed ' + next + ', expected I');
        assert(next !== 'H', 'the images column must never be proposed');
        return 'links would go into a new column I';
      });
    })
    .then(function () {
      return check('column letters round trip through colName and colIndex', function () {
        ['A', 'H', 'I', 'Z', 'AA', 'AZ', 'BA', 'ZZ', 'AAA'].forEach(function (col) {
          assert(X.colName(X.colIndex(col)) === col, col + ' round tripped to ' + X.colName(X.colIndex(col)));
        });
        return 'A, H, I, Z, AA, AZ, BA, ZZ and AAA all survive';
      });
    })
    .then(function () {
      return check('writing the header makes a restart skip the rows already done', function () {
        // This is the resume guarantee. The header written on the first run has
        // to be the one findColumn looks for on the second, or every finished
        // row would be created a second time.
        var link = 'https://www.fygaro.com/en/pb/33333333-3333-3333-3333-333333333333/';
        return X.writeColumn(wb, SHEET, 'I', [
          { row: 1, value: S.LINK_HEADER },
          { row: 2, value: link }
        ])
          .then(function (blob) { return blob.arrayBuffer(); })
          .then(function (buffer) { return X.load(new Uint8Array(buffer)); })
          .then(function (again) {
            var data = again.readSheet(SHEET);
            var reread = X.readHeader(data);
            var mapped = X.findColumn(reread, S.HEADERS.link);
            assert(mapped === 'I', 'on reload the Link column resolved to "' + mapped + '"');

            var rows = buildRows(data, reread, {
              name: 'D', code: 'C', price: 'E', link: mapped
            });
            var withLink = rows.filter(function (r) { return r.link; });
            assert(withLink.length === 1, withLink.length + ' rows carry a link, expected 1');
            assert(withLink[0].sheetRow === 2, 'the link came back on row ' + withLink[0].sheetRow);
            assert(rows.length === 30, 'the header row leaked into the data: ' + rows.length + ' rows');
            return 'the Link header is found again and row 2 is skipped';
          });
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
    // Report back to the test runner. The pages are driven in real time rather
    // than under a virtual clock, because Chrome's virtual time fast forwards
    // past IndexedDB completion callbacks whenever the page looks idle, which
    // left transactions hanging for ever.
    if (location.protocol.indexOf('http') === 0) {
      fetch('/__result', {
        method: 'POST',
        body: JSON.stringify({
          title: document.title,
          failures: results.filter(function (r) { return !r.ok; })
            .map(function (r) { return r.name + ' -- ' + r.detail; })
        })
      }).catch(function () {});
    }


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
