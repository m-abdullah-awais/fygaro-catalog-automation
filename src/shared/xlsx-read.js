/*
 * Fygaro Catalog Automation
 * Reads an xlsx workbook far enough to pull a sheet out as rows of text.
 *
 * Only the parts that matter are parsed: the sheet list, the shared string
 * table and one worksheet at a time. Every entry of the archive is kept in
 * memory untouched so the writer can hand back a file that differs in exactly
 * one place.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var X = FYG.xlsx || (FYG.xlsx = {});

  var MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  function decode(bytes) {
    return new TextDecoder('utf-8').decode(bytes);
  }

  function parseXml(text, label) {
    var doc = new DOMParser().parseFromString(text, 'application/xml');
    var err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('Could not parse ' + label + ': ' + FYG.util.normText(err.textContent));
    return doc;
  }

  /** Joins every <t> in a node, skipping phonetic hints. */
  function textFromRuns(node) {
    if (!node) return '';
    var out = '';
    var ts = node.getElementsByTagNameNS(MAIN_NS, 't');
    for (var i = 0; i < ts.length; i++) {
      if (ts[i].parentNode && ts[i].parentNode.localName === 'rPh') continue;
      out += ts[i].textContent;
    }
    return out;
  }

  /** "H18" -> "H" */
  X.colOf = function (ref) {
    var m = /^([A-Z]+)/.exec(String(ref || '').toUpperCase());
    return m ? m[1] : '';
  };

  /** "H18" -> 18 */
  X.rowOf = function (ref) {
    var m = /(\d+)$/.exec(String(ref || ''));
    return m ? parseInt(m[1], 10) : 0;
  };

  /** "A" -> 1, "H" -> 8, "AA" -> 27 */
  X.colIndex = function (letters) {
    var n = 0;
    letters = String(letters || '').toUpperCase();
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  };

  /**
   * Loads a workbook from raw bytes.
   * @param {ArrayBuffer|Uint8Array} buffer
   * @returns {Promise<object>} workbook handle
   */
  X.load = function (buffer) {
    return FYG.zip.read(buffer).then(function (archive) {
      var wb = {
        order: archive.order,
        files: archive.files,
        sheets: [],
        sharedStrings: null,
        _sheetCache: new Map()
      };

      var workbookXml = wb.files.get('xl/workbook.xml');
      if (!workbookXml) throw new Error('This file is not a valid xlsx workbook: xl/workbook.xml is missing.');

      // rId -> part path
      var relMap = new Map();
      var relsBytes = wb.files.get('xl/_rels/workbook.xml.rels');
      if (relsBytes) {
        var relsDoc = parseXml(decode(relsBytes), 'xl/_rels/workbook.xml.rels');
        var rels = relsDoc.getElementsByTagName('Relationship');
        for (var i = 0; i < rels.length; i++) {
          var target = rels[i].getAttribute('Target') || '';
          var path = target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
          relMap.set(rels[i].getAttribute('Id'), path);
        }
      }

      var wbDoc = parseXml(decode(workbookXml), 'xl/workbook.xml');
      var sheetNodes = wbDoc.getElementsByTagNameNS(MAIN_NS, 'sheet');
      for (var s = 0; s < sheetNodes.length; s++) {
        var rid = sheetNodes[s].getAttributeNS(REL_NS, 'id') || sheetNodes[s].getAttribute('r:id');
        wb.sheets.push({
          name: sheetNodes[s].getAttribute('name') || '',
          state: sheetNodes[s].getAttribute('state') || 'visible',
          path: relMap.get(rid) || ''
        });
      }
      if (!wb.sheets.length) throw new Error('This workbook has no sheets.');

      var sstBytes = wb.files.get('xl/sharedStrings.xml');
      if (sstBytes) {
        var sstDoc = parseXml(decode(sstBytes), 'xl/sharedStrings.xml');
        var sis = sstDoc.getElementsByTagNameNS(MAIN_NS, 'si');
        wb.sharedStrings = new Array(sis.length);
        for (var k = 0; k < sis.length; k++) wb.sharedStrings[k] = textFromRuns(sis[k]);
      } else {
        wb.sharedStrings = [];
      }

      wb.sheetNames = wb.sheets.map(function (x) { return x.name; });

      /** Finds a sheet by exact name, then by trimmed name, then case insensitively. */
      wb.findSheet = function (name) {
        var want = String(name == null ? '' : name);
        var exact = wb.sheets.filter(function (x) { return x.name === want; })[0];
        if (exact) return exact;
        var trimmed = wb.sheets.filter(function (x) { return x.name.trim() === want.trim(); })[0];
        if (trimmed) return trimmed;
        var lower = want.trim().toLowerCase();
        return wb.sheets.filter(function (x) { return x.name.trim().toLowerCase() === lower; })[0] || null;
      };

      /**
       * Reads a sheet as rows of plain text.
       * @returns {{name: string, path: string, rows: Array<{r: number, cells: object}>, maxRow: number}}
       */
      wb.readSheet = function (name) {
        if (wb._sheetCache.has(name)) return wb._sheetCache.get(name);

        var sheet = wb.findSheet(name);
        if (!sheet) throw new Error('Sheet "' + name + '" was not found in this workbook.');
        var bytes = wb.files.get(sheet.path);
        if (!bytes) throw new Error('Sheet part "' + sheet.path + '" is missing from the file.');

        var doc = parseXml(decode(bytes), sheet.path);
        var rowNodes = doc.getElementsByTagNameNS(MAIN_NS, 'row');
        var rows = [];
        var maxRow = 0;

        for (var i = 0; i < rowNodes.length; i++) {
          var rowNode = rowNodes[i];
          var rNum = parseInt(rowNode.getAttribute('r') || '0', 10);
          var cells = {};
          var cNodes = rowNode.getElementsByTagNameNS(MAIN_NS, 'c');
          var any = false;

          for (var j = 0; j < cNodes.length; j++) {
            var c = cNodes[j];
            var ref = c.getAttribute('r') || '';
            var col = X.colOf(ref);
            if (!col) continue;
            var t = c.getAttribute('t');
            var value = '';

            if (t === 's') {
              var vNode = c.getElementsByTagNameNS(MAIN_NS, 'v')[0];
              var idx = vNode ? parseInt(vNode.textContent, 10) : -1;
              value = (idx >= 0 && idx < wb.sharedStrings.length) ? wb.sharedStrings[idx] : '';
            } else if (t === 'inlineStr') {
              value = textFromRuns(c.getElementsByTagNameNS(MAIN_NS, 'is')[0]);
            } else if (t === 'b') {
              var bNode = c.getElementsByTagNameNS(MAIN_NS, 'v')[0];
              value = bNode && bNode.textContent === '1' ? 'TRUE' : 'FALSE';
            } else {
              // Numbers, dates and cached formula results all live in <v>.
              var vv = c.getElementsByTagNameNS(MAIN_NS, 'v')[0];
              value = vv ? vv.textContent : '';
            }

            value = value == null ? '' : String(value);
            if (value !== '') any = true;
            cells[col] = value;
          }

          if (rNum > maxRow) maxRow = rNum;
          if (any) rows.push({ r: rNum, cells: cells });
        }

        var result = { name: sheet.name, path: sheet.path, rows: rows, maxRow: maxRow };
        wb._sheetCache.set(name, result);
        return result;
      };

      return wb;
    });
  };

  /**
   * Builds a header label to column letter map from the first populated row.
   * Matching folds case and accents so "Código" still resolves if the file is
   * saved with a different accent form.
   * @returns {{headerRow: number, byLabel: object, labels: Array<{col: string, label: string}>}}
   */
  X.readHeader = function (sheetData) {
    var first = sheetData.rows[0];
    if (!first) return { headerRow: 0, byLabel: {}, labels: [] };
    var byLabel = {};
    var labels = [];
    Object.keys(first.cells).sort(function (a, b) { return X.colIndex(a) - X.colIndex(b); }).forEach(function (col) {
      var label = FYG.util.normText(first.cells[col]);
      if (!label) return;
      labels.push({ col: col, label: label });
      var key = FYG.util.foldText(label);
      if (!(key in byLabel)) byLabel[key] = col;
    });
    return { headerRow: first.r, byLabel: byLabel, labels: labels };
  };

  /** Resolves one of several acceptable header spellings to a column letter. */
  X.findColumn = function (header, candidates) {
    for (var i = 0; i < candidates.length; i++) {
      var key = FYG.util.foldText(candidates[i]);
      if (key in header.byLabel) return header.byLabel[key];
    }
    return '';
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
