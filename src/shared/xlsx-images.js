/*
 * Fygaro Catalog Automation
 * Finds the pictures anchored to each row of a sheet.
 *
 * A picture in a spreadsheet is not a cell value. It is a floating shape listed
 * in a drawing part, pinned to a cell by a row and column index, and pointing at
 * its bytes through two levels of relationship file. So none of this can be read
 * through readSheet, and all of it has to be resolved by hand:
 *
 *   xl/worksheets/sheet1.xml
 *     -> xl/worksheets/_rels/sheet1.xml.rels   (the drawing)
 *       -> xl/drawings/drawing1.xml            (the anchors)
 *         -> xl/drawings/_rels/drawing1.xml.rels
 *           -> xl/media/image8.png             (the bytes)
 *
 * Pictures are keyed by their part path, because the real catalog reuses 35
 * images across 1030 rows and holding one copy of each rather than one per row
 * is the difference between 65 MB and about 2 GB.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var X = FYG.xlsx || (FYG.xlsx = {});

  var REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /**
   * Resolves a relationship target against the part that declared it.
   * Handles "../media/x.png", "./x.png", a bare name, and a package absolute
   * "/xl/media/x.png".
   */
  X.resolvePart = function (fromPart, target) {
    var clean = String(target || '').replace(/\\/g, '/');
    if (!clean) return '';
    if (clean.charAt(0) === '/') return clean.slice(1);

    var base = String(fromPart || '').split('/');
    base.pop();
    clean.split('/').forEach(function (piece) {
      if (!piece || piece === '.') return;
      if (piece === '..') base.pop();
      else base.push(piece);
    });
    return base.join('/');
  };

  /** "xl/worksheets/sheet1.xml" becomes "xl/worksheets/_rels/sheet1.xml.rels". */
  X.relsPathFor = function (partPath) {
    var pieces = String(partPath || '').split('/');
    var file = pieces.pop();
    return pieces.concat(['_rels', file + '.rels']).join('/');
  };

  /** Reads one part's relationships, or an empty list when it has none. */
  X.readRels = function (wb, partPath) {
    var bytes = wb.files.get(X.relsPathFor(partPath));
    if (!bytes) return [];

    var doc = new DOMParser().parseFromString(new TextDecoder('utf-8').decode(bytes), 'application/xml');
    var out = [];
    var nodes = doc.getElementsByTagName('*');
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      if (el.localName !== 'Relationship') continue;
      out.push({
        id: el.getAttribute('Id') || '',
        type: el.getAttribute('Type') || '',
        target: el.getAttribute('Target') || '',
        external: (el.getAttribute('TargetMode') || '') === 'External'
      });
    }
    return out;
  };

  function childByLocalName(parent, name) {
    if (!parent) return null;
    var kids = parent.children || [];
    for (var i = 0; i < kids.length; i++) {
      if (kids[i].localName === name) return kids[i];
    }
    return null;
  }

  function intOf(parent, name) {
    var el = childByLocalName(parent, name);
    if (!el) return null;
    var n = parseInt(el.textContent, 10);
    return isNaN(n) ? null : n;
  }

  function relIdOf(blip) {
    return blip.getAttributeNS(REL_NS, 'embed') || blip.getAttribute('r:embed') || '';
  }

  function linkIdOf(blip) {
    return blip.getAttributeNS(REL_NS, 'link') || blip.getAttribute('r:link') || '';
  }

  /**
   * Reads every anchor in a drawing part.
   *
   * All three anchor kinds are handled even though the real catalog only uses
   * oneCellAnchor, because the sheet is hand maintained and a picture dragged
   * slightly differently becomes a twoCellAnchor without anyone noticing.
   *
   * @param {string} xmlText the drawing part, as text
   * @returns {Array<{kind, row, col, rId, name, reason}>} row is 1 based, or null
   */
  X.parseDrawingAnchors = function (xmlText) {
    var doc = new DOMParser().parseFromString(xmlText, 'application/xml');
    var root2 = doc.documentElement;
    if (!root2) return [];

    var out = [];
    var anchors = root2.children || [];
    for (var i = 0; i < anchors.length; i++) {
      var anchor = anchors[i];
      var kind = anchor.localName;
      if (kind.indexOf('Anchor') === -1) continue;

      // absoluteAnchor is pinned to the sheet rather than to a cell, so it
      // belongs to no row and cannot be attributed to a product.
      var from = childByLocalName(anchor, 'from');
      var row = from ? intOf(from, 'row') : null;
      var col = from ? intOf(from, 'col') : null;

      // Every blip inside a blipFill, so a group of pictures in one anchor
      // yields all of them rather than only the first. A blip somewhere else,
      // for example decorating a shape's fill, is not a picture and is ignored.
      var blips = anchor.getElementsByTagName('*');
      for (var j = 0; j < blips.length; j++) {
        var blip = blips[j];
        if (blip.localName !== 'blip') continue;
        if (!blip.parentNode || blip.parentNode.localName !== 'blipFill') continue;

        var pic = blip.parentNode.parentNode;
        var nvPr = pic ? childByLocalName(childByLocalName(pic, 'nvPicPr'), 'cNvPr') : null;

        out.push({
          kind: kind,
          row: row === null ? null : row + 1,
          col: col,
          rId: relIdOf(blip),
          linkId: linkIdOf(blip),
          name: nvPr ? (nvPr.getAttribute('name') || '') : '',
          reason: from ? '' : 'anchored to the sheet rather than to a row'
        });
      }
    }
    return out;
  };

  var SIGNATURES = [
    { type: 'image/png', bytes: [0x89, 0x50, 0x4E, 0x47] },
    { type: 'image/jpeg', bytes: [0xFF, 0xD8, 0xFF] },
    { type: 'image/gif', bytes: [0x47, 0x49, 0x46, 0x38] },
    { type: 'image/bmp', bytes: [0x42, 0x4D] }
  ];

  var BY_EXTENSION = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
    gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp', svg: 'image/svg+xml'
  };

  /**
   * Identifies an image from its leading bytes, falling back to its name.
   *
   * The bytes win, because names lie: the real catalog carries a genuine JPEG,
   * and uploading it labelled as a PNG is the kind of thing a strict server
   * rejects for no visible reason.
   */
  X.sniffImageType = function (bytes, fileName) {
    for (var i = 0; i < SIGNATURES.length; i++) {
      var sig = SIGNATURES[i];
      var hit = bytes.length >= sig.bytes.length;
      for (var j = 0; hit && j < sig.bytes.length; j++) {
        if (bytes[j] !== sig.bytes[j]) hit = false;
      }
      if (hit) return sig.type;
    }
    // RIFF....WEBP
    if (bytes.length > 12 && bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 &&
      bytes[3] === 0x46 && bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 &&
      bytes[11] === 0x50) {
      return 'image/webp';
    }
    var ext = String(fileName || '').split('.').pop().toLowerCase();
    return BY_EXTENSION[ext] || '';
  };

  /**
   * Maps every row of a sheet that carries a picture to the pictures it uses.
   *
   * A sheet with no drawing, a drawing with no relationships part, or an empty
   * drawing all yield an empty result rather than an error. All three occur in
   * the real workbook, on the sheets nobody put pictures on.
   *
   * @param {object} wb workbook handle from FYG.xlsx.load
   * @param {string} sheetName
   * @returns {{byRow: Map<number, string[]>,
   *            images: Map<string, {id, name, type, size, bytes}>,
   *            anchors: number, skipped: Array<{reason: string, detail: string}>}}
   */
  X.readImages = function (wb, sheetName) {
    var sheet = wb.findSheet(sheetName);
    if (!sheet) throw new Error('Sheet "' + sheetName + '" was not found in this workbook.');

    if (!wb._imageCache) wb._imageCache = new Map();
    if (wb._imageCache.has(sheet.name)) return wb._imageCache.get(sheet.name);

    var empty = { byRow: new Map(), images: new Map(), anchors: 0, skipped: [] };
    var result = empty;

    var drawingRel = X.readRels(wb, sheet.path).filter(function (rel) {
      return /\/drawing$/.test(rel.type) && !rel.external;
    })[0];

    if (drawingRel) {
      var drawingPath = X.resolvePart(sheet.path, drawingRel.target);
      var drawingBytes = wb.files.get(drawingPath);
      if (drawingBytes) {
        var media = {};
        X.readRels(wb, drawingPath).forEach(function (rel) { media[rel.id] = rel; });

        var byRow = new Map();
        var images = new Map();
        var skipped = [];
        var anchors = X.parseDrawingAnchors(new TextDecoder('utf-8').decode(drawingBytes));

        anchors.forEach(function (anchor) {
          if (anchor.row === null) {
            skipped.push({ reason: 'unanchored', detail: anchor.name || anchor.kind });
            return;
          }
          if (!anchor.rId) {
            skipped.push({
              reason: anchor.linkId ? 'linked' : 'no image',
              detail: 'row ' + anchor.row + ' ' + (anchor.name || anchor.kind)
            });
            return;
          }
          var rel = media[anchor.rId];
          if (!rel || rel.external) {
            skipped.push({ reason: 'external', detail: 'row ' + anchor.row });
            return;
          }

          var id = X.resolvePart(drawingPath, rel.target);
          var bytes = wb.files.get(id);
          if (!bytes) {
            skipped.push({ reason: 'missing', detail: id });
            return;
          }

          if (!images.has(id)) {
            var base = id.split('/').pop();
            var type = X.sniffImageType(bytes, base);
            if (!type) {
              skipped.push({ reason: 'unsupported', detail: base });
              return;
            }
            images.set(id, { id: id, name: base, type: type, size: bytes.length, bytes: bytes });
          }

          var list = byRow.get(anchor.row);
          if (!list) { list = []; byRow.set(anchor.row, list); }
          if (list.indexOf(id) === -1) list.push(id);
        });

        result = { byRow: byRow, images: images, anchors: anchors.length, skipped: skipped };
      }
    }

    wb._imageCache.set(sheet.name, result);
    return result;
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
