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
