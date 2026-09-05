/*
 * Fygaro Catalog Automation
 * Writes generated links back into one column of one sheet.
 *
 * The worksheet XML is edited as text rather than reparsed and reserialised, so
 * styles, formulas, data validations, hyperlinks, drawings and every other sheet
 * survive byte for byte. Values are written as inline strings, which means the
 * shared string table is never touched and cannot fall out of sync.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var X = FYG.xlsx || (FYG.xlsx = {});

  function escapeXml(s) { return FYG.util.escapeXml(s); }

  /** Rebuilds one cell, keeping its style attribute so formatting is preserved. */
  function buildCell(existingMarkup, ref, value) {
    var style = '';
    if (existingMarkup) {
      var openTag = existingMarkup.slice(0, existingMarkup.indexOf('>') + 1);
      var m = /\bs="([^"]*)"/.exec(openTag);
      if (m) style = ' s="' + m[1] + '"';
    }
    if (value === '' || value == null) return '<c r="' + ref + '"' + style + '/>';
    return '<c r="' + ref + '"' + style + ' t="inlineStr"><is><t xml:space="preserve">' +
      escapeXml(value) + '</t></is></c>';
  }

  function cellRegex(ref, flags) {
    return new RegExp('<c\\b[^>]*\\br="' + ref + '"[^>]*?(?:/>|>[\\s\\S]*?</c>)', flags || '');
  }

  /**
   * Records where every row element starts and ends, in one pass.
   *
   * Locating rows by running a fresh regex over the whole sheet once per cell is
   * quadratic. That costs nothing while the target cells already exist, and it
   * is ruinous when they do not: writing a brand new column into the real
   * catalog measured 15 seconds against 230 ms for a column that was already
   * there. Everything is found in this single scan instead.
   *
   * @param {string} xml worksheet part, as text
   * @returns {Map<number, {start: number, end: number}>}
   */
  X.indexRows = function (xml) {
    var byRow = new Map();
    var open = /<row\b[^>]*>/g;
    var m;
    while ((m = open.exec(xml)) !== null) {
      var tag = m[0];
      var num = /\br="(\d+)"/.exec(tag);
      if (!num) continue;
      var end;
      if (tag.charAt(tag.length - 2) === '/') {
        end = m.index + tag.length;
      } else {
        var close = xml.indexOf('</row>', m.index + tag.length);
        if (close === -1) continue;
        end = close + '</row>'.length;
      }
      byRow.set(parseInt(num[1], 10), { start: m.index, end: end });
      // Rows never nest, so the body is skipped wholesale, which is what keeps
      // this linear in the size of the sheet.
      open.lastIndex = end;
    }
    return byRow;
  };

  /**
   * Returns one row element with the target cell written into it: replaced when
   * the cell is already there, inserted in column order when it is not.
   *
   * A cell that has to be created copies the style of the nearest cell to its
   * left, so a brand new column matches the sheet it is joining rather than
   * arriving unformatted beside a fully styled table.
   */
  function writeCellIntoRow(rowMarkup, column, rowNumber, value) {
    var ref = column + rowNumber;

    // Every cell inside one row ends with that row number, so this cannot
    // collide with a longer reference such as I10 while patching row 1.
    var existing = cellRegex(ref).exec(rowMarkup);
    if (existing) {
      return rowMarkup.slice(0, existing.index) +
        buildCell(existing[0], ref, value) +
        rowMarkup.slice(existing.index + existing[0].length);
    }

    if (/\/>$/.test(rowMarkup)) {
      // An empty self closing row has to become a container first.
      return rowMarkup.slice(0, -2) + '>' + buildCell(null, ref, value) + '</row>';
    }

    var targetIndex = X.colIndex(column);
    var bodyStart = rowMarkup.indexOf('>') + 1;
    var head = rowMarkup.slice(0, bodyStart);
    var body = rowMarkup.slice(bodyStart, rowMarkup.length - '</row>'.length);

    var insertAt = body.length;
    var neighbour = null;
    var scan = /<c\b[^>]*\br="([A-Z]+)\d+"[^>]*?(?:\/>|>)/g;
    var found;
    while ((found = scan.exec(body)) !== null) {
      if (X.colIndex(found[1]) > targetIndex) { insertAt = found.index; break; }
      neighbour = found[0];
    }

    return head + body.slice(0, insertAt) + buildCell(neighbour, ref, value) +
      body.slice(insertAt) + '</row>';
  }

  /** Slow path used only when the sheet has no row element at all for a number. */
  function insertRow(xml, column, rowNumber, value) {
    var newRow = '<row r="' + rowNumber + '">' + buildCell(null, column + rowNumber, value) + '</row>';
    var rowScan = /<row\b[^>]*\br="(\d+)"/g;
    var hit;
    while ((hit = rowScan.exec(xml)) !== null) {
      if (parseInt(hit[1], 10) > rowNumber) {
        return xml.slice(0, hit.index) + newRow + xml.slice(hit.index);
      }
    }
    if (xml.indexOf('</sheetData>') !== -1) {
      return xml.replace('</sheetData>', newRow + '</sheetData>');
    }
    return xml.replace(/<sheetData\s*\/>/, '<sheetData>' + newRow + '</sheetData>');
  }

  /**
   * Applies a batch of single column updates to worksheet XML.
   * @param {string} xml worksheet part, as text
   * @param {string} column column letter, for example "H"
   * @param {Array<{row: number, value: string}>} updates
   * @returns {string} the patched XML
   */
  X.patchSheetXml = function (xml, column, updates) {
    // Last write wins for a repeated row, so the forward pass stays monotonic.
    var wanted = new Map();
    updates.forEach(function (u) { wanted.set(u.row, u.value); });
    var sorted = Array.from(wanted.keys()).sort(function (a, b) { return a - b; });

    var index = X.indexRows(xml);
    var parts = [];
    var cursor = 0;
    var missingRows = [];

    for (var i = 0; i < sorted.length; i++) {
      var number = sorted[i];
      var slot = index.get(number);
      if (!slot) { missingRows.push(number); continue; }
      parts.push(xml.slice(cursor, slot.start));
      parts.push(writeCellIntoRow(xml.slice(slot.start, slot.end), column, number, wanted.get(number)));
      cursor = slot.end;
    }
    parts.push(xml.slice(cursor));

    var result = parts.join('');
    for (var j = 0; j < missingRows.length; j++) {
      result = insertRow(result, column, missingRows[j], wanted.get(missingRows[j]));
    }
    return result;
  };

  /**
   * Applies several column batches, then rebuilds the file once.
   * @param {object} wb workbook handle from FYG.xlsx.load
   * @param {string} sheetName
   * @param {Array<{column: string, updates: Array<{row: number, value: string}>}>} batches
   * @returns {Promise<Blob>}
   */
  X.writeColumns = function (wb, sheetName, batches) {
    return Promise.resolve().then(function () {
      var sheet = wb.findSheet(sheetName);
      if (!sheet) throw new Error('Sheet "' + sheetName + '" was not found in this workbook.');
      var bytes = wb.files.get(sheet.path);
      if (!bytes) throw new Error('Sheet part "' + sheet.path + '" is missing from the file.');

      var xml = new TextDecoder('utf-8').decode(bytes);
      (batches || []).forEach(function (batch) {
        if (!batch || !batch.column || !batch.updates || !batch.updates.length) return;
        xml = X.patchSheetXml(xml, batch.column, batch.updates);
      });

      // Write into a copy of the file table so the loaded workbook stays intact
      // and the user can export again after more rows finish.
      var files = new Map(wb.files);
      files.set(sheet.path, new TextEncoder().encode(xml));
      return FYG.zip.write(wb.order, files);
    });
  };

  /**
   * Produces a new workbook file with the given column filled in.
   * The original bytes handed to X.load are never modified.
   * @param {object} wb workbook handle from FYG.xlsx.load
   * @param {string} sheetName
   * @param {string} column
   * @param {Array<{row: number, value: string}>} updates
   * @returns {Promise<Blob>}
   */
  X.writeColumn = function (wb, sheetName, column, updates) {
    return X.writeColumns(wb, sheetName, [{ column: column, updates: updates }]);
  };

  /** Builds the download name for the patched workbook. */
  X.exportName = function (originalName) {
    var base = String(originalName || 'catalogo.xlsx').replace(/\.xlsx$/i, '');
    return base + ' (con Links) ' + FYG.util.todayStamp() + '.xlsx';
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
