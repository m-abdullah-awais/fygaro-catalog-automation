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

  function rowRegex(rowNumber) {
    return new RegExp('<row\\b[^>]*\\br="' + rowNumber + '"[^>]*?(?:/>|>[\\s\\S]*?</row>)');
  }

  /** Slow path used only when the target cell does not already exist. */
  function insertCell(xml, column, rowNumber, value) {
    var ref = column + rowNumber;
    var newCell = buildCell(null, ref, value);
    var targetIndex = X.colIndex(column);

    var rowMatch = rowRegex(rowNumber).exec(xml);
    if (rowMatch) {
      var rowMarkup = rowMatch[0];
      var updatedRow;
      if (/\/>$/.test(rowMarkup)) {
        // An empty self closing row has to become a container first.
        updatedRow = rowMarkup.slice(0, -2) + '>' + newCell + '</row>';
      } else {
        var bodyStart = rowMarkup.indexOf('>') + 1;
        var head = rowMarkup.slice(0, bodyStart);
        var body = rowMarkup.slice(bodyStart, rowMarkup.length - '</row>'.length);
        var insertAt = body.length;
        var scan = /<c\b[^>]*\br="([A-Z]+)\d+"/g;
        var found;
        while ((found = scan.exec(body)) !== null) {
          if (X.colIndex(found[1]) > targetIndex) { insertAt = found.index; break; }
        }
        updatedRow = head + body.slice(0, insertAt) + newCell + body.slice(insertAt) + '</row>';
      }
      return xml.slice(0, rowMatch.index) + updatedRow + xml.slice(rowMatch.index + rowMarkup.length);
    }

    // The row itself is missing, so create it in row order.
    var newRow = '<row r="' + rowNumber + '">' + newCell + '</row>';
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
    var sorted = updates.slice().sort(function (a, b) { return a.row - b.row; });
    var parts = [];
    var cursor = 0;
    var missed = [];

    // Rows appear in ascending order, so one forward pass handles the normal
    // case where every target cell already exists.
    for (var i = 0; i < sorted.length; i++) {
      var u = sorted[i];
      var ref = column + u.row;
      var re = cellRegex(ref, 'g');
      re.lastIndex = cursor;
      var m = re.exec(xml);
      if (!m) { missed.push(u); continue; }
      parts.push(xml.slice(cursor, m.index));
      parts.push(buildCell(m[0], ref, u.value));
      cursor = m.index + m[0].length;
    }
    parts.push(xml.slice(cursor));

    var result = parts.join('');
    for (var j = 0; j < missed.length; j++) {
      result = insertCell(result, column, missed[j].row, missed[j].value);
    }
    return result;
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
    return Promise.resolve().then(function () {
      var sheet = wb.findSheet(sheetName);
      if (!sheet) throw new Error('Sheet "' + sheetName + '" was not found in this workbook.');
      var bytes = wb.files.get(sheet.path);
      if (!bytes) throw new Error('Sheet part "' + sheet.path + '" is missing from the file.');

      var xml = new TextDecoder('utf-8').decode(bytes);
      var patched = X.patchSheetXml(xml, column, updates);

      // Write into a copy of the file table so the loaded workbook stays intact
      // and the user can export again after more rows finish.
      var files = new Map(wb.files);
      files.set(sheet.path, new TextEncoder().encode(patched));
      return FYG.zip.write(wb.order, files);
    });
  };

  /** Builds the download name for the patched workbook. */
  X.exportName = function (originalName) {
    var base = String(originalName || 'catalogo.xlsx').replace(/\.xlsx$/i, '');
    return base + ' (con Links) ' + FYG.util.todayStamp() + '.xlsx';
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
