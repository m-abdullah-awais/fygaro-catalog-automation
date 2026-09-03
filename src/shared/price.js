/*
 * Fygaro Catalog Automation
 * Price parsing for the "Precio Total" column.
 *
 * The real catalog mixes two number conventions in the same column:
 *   B/.625,00     comma is the decimal separator
 *   B/.1.125,00   dot is the thousands separator, comma is the decimal
 *   B/.100.00     dot is the decimal separator
 * They all have to land on the same number, so the role of each separator is
 * decided per value rather than assumed for the whole sheet.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var P = FYG.price || (FYG.price = {});

  /**
   * Parses a catalog price string.
   * @param {string|number} raw
   * @returns {{ok: boolean, value: number|null, text: string, raw: string, reason: string|null}}
   */
  P.parse = function (raw) {
    var out = { ok: false, value: null, text: '', raw: raw == null ? '' : String(raw), reason: null };

    if (typeof raw === 'number' && isFinite(raw)) {
      out.ok = raw >= 0;
      out.value = raw;
      out.text = P.format(raw);
      if (!out.ok) out.reason = 'negative price';
      return out;
    }

    var s = String(raw == null ? '' : raw).replace(/\u00a0/g, ' ').trim();
    if (!s) { out.reason = 'empty'; return out; }

    // Drop the currency prefix and any trailing label. "B/.1.125,00" -> "1.125,00"
    s = s.replace(/^[^\d-]+/, '').replace(/[^\d.,]+$/, '');
    if (!s) { out.reason = 'no digits'; return out; }

    var negative = /^-/.test(s);
    s = s.replace(/^-/, '');

    var lastComma = s.lastIndexOf(',');
    var lastDot = s.lastIndexOf('.');
    var normalized;

    if (lastComma !== -1 && lastDot !== -1) {
      // Both separators appear, so whichever comes last is the decimal point and
      // the other one is grouping. Covers 1.125,00 and 1,299.99 alike.
      normalized = split(Math.max(lastComma, lastDot));
    } else if (lastComma !== -1 || lastDot !== -1) {
      // Only one kind of separator. A final group of exactly three digits reads
      // as thousands, anything shorter reads as a decimal fraction.
      var pos = lastComma !== -1 ? lastComma : lastDot;
      normalized = s.slice(pos + 1).length === 3 ? s.replace(/[.,]/g, '') : split(pos);
    } else {
      normalized = s;
    }

    function split(at) {
      var intPart = s.slice(0, at).replace(/[.,]/g, '');
      var decPart = s.slice(at + 1).replace(/[^\d]/g, '');
      return intPart + (decPart ? '.' + decPart : '');
    }

    if (!/^\d+(\.\d+)?$/.test(normalized)) { out.reason = 'unrecognised format'; return out; }

    var value = parseFloat(normalized);
    if (!isFinite(value)) { out.reason = 'not a number'; return out; }
    if (negative) value = -value;
    if (value < 0) { out.reason = 'negative price'; out.value = value; return out; }

    out.ok = true;
    out.value = value;
    out.text = P.format(value);
    return out;
  };

  /** Renders a number the way the Fygaro price input expects it. */
  P.format = function (n) {
    return (Math.round(n * 100) / 100).toFixed(2);
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
