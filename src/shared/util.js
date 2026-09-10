/*
 * Fygaro Catalog Automation
 * Shared helpers. Loaded as a classic script in the service worker, the content
 * script and the side panel, so everything hangs off one global namespace.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var U = FYG.util || (FYG.util = {});

  /** Resolves after `ms` milliseconds. */
  U.sleep = function (ms) {
    return new Promise(function (resolve) { setTimeout(resolve, Math.max(0, ms | 0)); });
  };

  /** Inclusive random integer between min and max. */
  U.randInt = function (min, max) {
    if (max < min) { var t = min; min = max; max = t; }
    return min + Math.floor(Math.random() * (max - min + 1));
  };

  U.clamp = function (n, min, max) {
    return n < min ? min : (n > max ? max : n);
  };

  /**
   * Waits a random amount of time inside [min, max]. Used before every click and
   * every field write so the automation never fires at a fixed machine cadence.
   */
  U.randomDelay = function (min, max) {
    return U.sleep(U.randInt(min, max));
  };

  /**
   * Normalises page text for comparison: non breaking spaces become normal
   * spaces, runs of whitespace collapse, and the result is trimmed.
   */
  /*
   * Bytes to base64 and back. Needed because chrome.runtime messages are JSON,
   * so an ArrayBuffer cannot survive the trip to the page.
   *
   * The conversion is chunked because String.fromCharCode.apply throws on a
   * large enough array, and a product picture is comfortably large enough.
   */
  U.bytesToBase64 = function (bytes) {
    var chunk = 0x8000;
    var parts = [];
    for (var i = 0; i < bytes.length; i += chunk) {
      parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunk)));
    }
    return btoa(parts.join(''));
  };

  U.base64ToBytes = function (b64) {
    var binary = atob(b64);
    var out = new Uint8Array(binary.length);
    for (var i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
  };

  U.normText = function (s) {
    return String(s == null ? '' : s)
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  };

  /** Normalised, lowercased, and stripped of accents, for tolerant matching. */
  U.foldText = function (s) {
    return U.normText(s)
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  };

  /** Visible text of an element, normalised. */
  U.textOf = function (el) {
    return el ? U.normText(el.textContent) : '';
  };

  /** Pulls the first uuid out of a string, or null. */
  U.extractUuid = function (s) {
    var m = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i.exec(String(s || ''));
    return m ? m[0].toLowerCase() : null;
  };

  /*
   * A Fygaro payment link, for example
   * https://www.fygaro.com/en/pb/65df9668-0a2e-4c0b-a8f7-0789e47be346/
   *
   * This lives here rather than in the content script because two quite
   * different places need the same answer: the step that reads a freshly
   * generated link off the page, and the sheet reader deciding whether a row is
   * already done. Two copies of this shape would drift, and the way they would
   * drift is that one of them starts accepting something the other rejects,
   * which is exactly how a row gets created in Fygaro twice.
   *
   * The host is not pinned. Fygaro serves the same button under /en/ and /es/,
   * and a link pasted from a custom domain is still a link. What is pinned is
   * the /pb/ path and a uuid shaped segment, which is enough to tell a real
   * link apart from a note, a price, or a half typed cell.
   */
  U.LINK_PATTERN = /^https?:\/\/[^\s]*\/pb\/[0-9a-f-]{36}\/?$/i;

  /** True when the whole value is a Fygaro payment link and nothing else. */
  U.isFygaroLink = function (s) {
    return U.LINK_PATTERN.test(U.normText(s));
  };

  /*
   * The Fygaro link inside a value, or '' when there is not one.
   *
   * Tolerant on purpose. A cell that reads "Link: https://...fygaro.com/pb/x/"
   * or carries a trailing full stop still means the row is done, and treating it
   * as unlinked would send a product Fygaro already has back through the form.
   * The link that comes back is the clean one, so what gets written to the sheet
   * on export is the link and not the sentence around it.
   */
  U.fygaroLink = function (s) {
    var text = U.normText(s);
    if (U.LINK_PATTERN.test(text)) return text;
    var m = /https?:\/\/\S*?\/pb\/[0-9a-f-]{36}\/?/i.exec(text);
    return m ? m[0] : '';
  };

  U.truncate = function (s, n) {
    s = String(s == null ? '' : s);
    return s.length <= n ? s : s.slice(0, Math.max(0, n - 1)) + '\u2026';
  };

  /** Quotes one CSV field per RFC 4180. */
  U.csvCell = function (v) {
    var s = v == null ? '' : String(v);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };

  /** rows is an array of arrays. Emits a BOM so Excel reads UTF-8 correctly. */
  U.toCsv = function (rows) {
    return '\ufeff' + rows.map(function (r) {
      return r.map(U.csvCell).join(',');
    }).join('\r\n');
  };

  /** "1h 04m 12s" style duration, for the run ETA. */
  U.formatDuration = function (ms) {
    if (!isFinite(ms) || ms < 0) return '--';
    var s = Math.round(ms / 1000);
    var h = Math.floor(s / 3600);
    var m = Math.floor((s % 3600) / 60);
    var sec = s % 60;
    function pad(n) { return n < 10 ? '0' + n : String(n); }
    if (h > 0) return h + 'h ' + pad(m) + 'm';
    if (m > 0) return m + 'm ' + pad(sec) + 's';
    return sec + 's';
  };

  /** Local HH:MM:SS for log lines. */
  U.clockTime = function (ts) {
    var d = new Date(ts == null ? Date.now() : ts);
    function pad(n) { return n < 10 ? '0' + n : String(n); }
    return pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  };

  /** YYYY-MM-DD in local time, for the exported file name. */
  U.todayStamp = function () {
    var d = new Date();
    function pad(n) { return n < 10 ? '0' + n : String(n); }
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  };

  /** Appends to a capped log array and returns the new array. */
  U.pushLog = function (log, level, message, cap) {
    log = Array.isArray(log) ? log : [];
    log.push({ ts: Date.now(), level: level || 'info', message: String(message) });
    var max = cap || 500;
    return log.length > max ? log.slice(log.length - max) : log;
  };

  /** Escapes text for safe insertion into XML. */
  U.escapeXml = function (s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
