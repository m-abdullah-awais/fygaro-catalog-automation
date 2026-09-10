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

  /**
   * Builds a PNG of random pixels, which barely compresses, so a modest canvas
   * yields a genuinely large file to test the size fitter against.
   */
  function makeNoisyPng(side) {
    var canvas = document.createElement('canvas');
    canvas.width = side;
    canvas.height = side;
    var ctx = canvas.getContext('2d');
    var picture = ctx.createImageData(side, side);
    for (var i = 0; i < picture.data.length; i += 4) {
      picture.data[i] = Math.random() * 256;
      picture.data[i + 1] = Math.random() * 256;
      picture.data[i + 2] = Math.random() * 256;
      picture.data[i + 3] = 255;
    }
    ctx.putImageData(picture, 0, 0);
    return new Promise(function (resolve) {
      canvas.toBlob(function (blob) {
        blob.arrayBuffer().then(function (buffer) { resolve(new Uint8Array(buffer)); });
      }, 'image/png');
    });
  }

  /*
   * Mirrors the row building in src/sidepanel/sidepanel.js, including the rule
   * that decides a row is already done. The panel version also carries the row
   * range and the zero price check, which nothing here exercises.
   */
  function linkAnywhereIn(cells) {
    for (var col in cells) {
      if (!Object.prototype.hasOwnProperty.call(cells, col)) continue;
      var found = U.fygaroLink(cells[col]);
      if (found) return found;
    }
    return '';
  }

  function buildRows(sheetData, header, mapping, searchWholeRow) {
    var rows = [];
    sheetData.rows.forEach(function (r) {
      if (r.r <= header.headerRow) return;
      var name = U.normText(mapping.name ? r.cells[mapping.name] : '');
      var code = U.normText(mapping.code ? r.cells[mapping.code] : '');
      var priceRaw = U.normText(mapping.price ? r.cells[mapping.price] : '');
      if (!name && !code) return;
      var linkCell = U.normText(mapping.link ? r.cells[mapping.link] : '');
      var link = U.fygaroLink(linkCell);
      if (!link && searchWholeRow) link = linkAnywhereIn(r.cells);
      var parsed = FYG.price.parse(priceRaw);
      rows.push({
        sheetRow: r.r, name: name, code: code, priceRaw: priceRaw,
        priceText: parsed.ok ? parsed.text : '', link: link,
        linkJunk: !link && !!linkCell,
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
      return check('a picture within the limit is stored byte for byte', function () {
        var found = X.readImages(wb, SHEET);
        var img = found.images.get(found.byRow.get(2)[0]);
        return FYG.imagefit.fit(img.bytes, img.type, img.name).then(function (fitted) {
          assert(fitted.changed === false, 'a small picture should not be re-encoded');
          assert(fitted.bytes === img.bytes, 'the original bytes should be handed straight back');
          assert(fitted.name === img.name, 'the name changed to ' + fitted.name);
          return 'left alone at ' + img.bytes.length + ' bytes';
        });
      });
    })
    .then(function () {
      return check('a picture over the limit is brought under it', function () {
        // Fygaro refuses anything over 2.5 MB, counting a megabyte as 1,000,000
        // bytes. Rather than build a 2.5 MB fixture, the target is lowered so
        // the same code path runs against a picture that really is too big.
        var realTarget = FYG.imagefit.TARGET_BYTES;
        return makeNoisyPng(220).then(function (big) {
          FYG.imagefit.TARGET_BYTES = Math.floor(big.length / 3);
          return FYG.imagefit.fit(big, 'image/png', 'huge.png')
            .then(function (fitted) {
              FYG.imagefit.TARGET_BYTES = realTarget;
              assert(fitted.changed === true, 'an oversized picture should be re-encoded');
              assert(fitted.bytes.length <= Math.floor(big.length / 3),
                'still ' + fitted.bytes.length + ' bytes, over the target');
              assert(fitted.bytes.length < big.length, 'it did not actually get smaller');
              assert(fitted.type === 'image/jpeg', 'it should come out as JPEG, saw ' + fitted.type);
              assert(fitted.name === 'huge.jpg', 'the name should follow the format, saw ' + fitted.name);
              assert(fitted.bytes[0] === 0xFF && fitted.bytes[1] === 0xD8, 'the bytes are not a JPEG');
              return big.length + ' bytes down to ' + fitted.bytes.length;
            })
            .catch(function (err) {
              FYG.imagefit.TARGET_BYTES = realTarget;
              throw err;
            });
        });
      });
    })
    .then(function () {
      return check('a picture that cannot be decoded is left alone rather than lost', function () {
        var notAnImage = new Uint8Array(64);
        var realTarget = FYG.imagefit.TARGET_BYTES;
        FYG.imagefit.TARGET_BYTES = 8;
        return FYG.imagefit.fit(notAnImage, 'image/png', 'broken.png')
          .then(function (fitted) {
            FYG.imagefit.TARGET_BYTES = realTarget;
            assert(fitted.changed === false, 'nothing should have been re-encoded');
            assert(fitted.bytes === notAnImage, 'the original bytes should survive');
            assert(fitted.note, 'it should say why it was left alone');
            return fitted.note;
          })
          .catch(function (err) {
            FYG.imagefit.TARGET_BYTES = realTarget;
            throw err;
          });
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
    .then(function () {
      return check('text in the Link column that is not a link does not retire the row', function () {
        /*
         * The Nota column sits one place past the Link column and says things
         * like "Fuera del rango elegido". Reading any non empty cell as a link
         * meant a mapping off by one column quietly retired every row in the
         * sheet, and a run would report there was nothing left to do.
         */
        var link = 'https://www.fygaro.com/en/pb/44444444-4444-4444-4444-444444444444/';
        return X.writeColumn(wb, SHEET, 'I', [
          { row: 1, value: S.LINK_HEADER },
          { row: 2, value: link },
          { row: 3, value: 'Fuera del rango elegido (5 a 10).' },
          { row: 4, value: 'Ya existe en Fygaro. Nothing was created and no link was captured.' },
          { row: 5, value: 'pending' }
        ])
          .then(function (blob) { return blob.arrayBuffer(); })
          .then(function (buffer) { return X.load(new Uint8Array(buffer)); })
          .then(function (again) {
            var data = again.readSheet(SHEET);
            var reread = X.readHeader(data);
            var rows = buildRows(data, reread, {
              name: 'D', code: 'C', price: 'E', link: X.findColumn(reread, S.HEADERS.link)
            });
            var byRow = {};
            rows.forEach(function (r) { byRow[r.sheetRow] = r; });

            assert(byRow[2].link === link, 'row 2 should carry the link, saw "' + byRow[2].link + '"');
            assert(!byRow[2].linkJunk, 'row 2 is a real link, not junk');
            [3, 4, 5].forEach(function (n) {
              assert(byRow[n].link === '', 'row ' + n + ' should carry no link, saw "' + byRow[n].link + '"');
              assert(byRow[n].linkJunk, 'row ' + n + ' should be flagged as junk in the link column');
            });

            var withLink = rows.filter(function (r) { return r.link; });
            assert(withLink.length === 1, withLink.length + ' rows counted as linked, expected 1');
            return 'one real link kept, three notes rejected and left to do';
          });
      });
    })
    .then(function () {
      return check('a link in another column is found only when no Link column exists', function () {
        /*
         * When the sheet has no Link column the mapping points at one that does
         * not exist yet, so every cell under it reads as empty. A link written
         * by an earlier run under a heading this build does not know would be
         * invisible, and the row would go to Fygaro and come back refused for a
         * duplicate code. Looking across the row closes that. It must not happen
         * when a real Link column is present, or a link quoted anywhere else
         * would retire a row that still needs doing.
         */
        var link = 'https://www.fygaro.com/en/pb/55555555-5555-5555-5555-555555555555/';
        return X.writeColumn(wb, SHEET, 'P', [{ row: 5, value: link }])
          .then(function (blob) { return blob.arrayBuffer(); })
          .then(function (buffer) { return X.load(new Uint8Array(buffer)); })
          .then(function (again) {
            var data = again.readSheet(SHEET);
            var reread = X.readHeader(data);
            assert(X.findColumn(reread, S.HEADERS.link) === '', 'this sheet should still have no Link column');

            // 'I' is what nextFreeColumn would propose. Nothing is under it.
            var mapping = { name: 'D', code: 'C', price: 'E', link: 'I' };

            var scanned = buildRows(data, reread, mapping, true);
            var found = scanned.filter(function (r) { return r.link; });
            assert(found.length === 1, found.length + ' rows found a stray link, expected 1');
            assert(found[0].sheetRow === 5, 'the stray link came back on row ' + found[0].sheetRow);
            assert(found[0].link === link, 'the link came back as "' + found[0].link + '"');

            var plain = buildRows(data, reread, mapping, false);
            assert(plain.filter(function (r) { return r.link; }).length === 0,
              'without the scan no row should count as linked');
            return 'row 5 recognised from column P, and ignored when a Link column is mapped';
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
