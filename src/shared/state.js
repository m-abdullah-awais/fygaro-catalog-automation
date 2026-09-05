/*
 * Fygaro Catalog Automation
 * The shape of a run, where it is stored, and the vocabulary every part of the
 * extension uses to talk about it.
 *
 * Storage is split into three keys on purpose. The hot state changes on every
 * step, the row table only changes when a row finishes, and the log is capped.
 * Keeping them apart avoids rewriting a 700 row array thousands of times.
 */
(function (root) {
  'use strict';

  var FYG = root.FYG || (root.FYG = {});
  var S = FYG.state || (FYG.state = {});

  S.KEY_RUN = 'fyg_run';
  S.KEY_ROWS = 'fyg_rows';
  S.KEY_LOG = 'fyg_log';
  /* The workbook bytes are kept too, so the updated file can still be exported
   * after the panel is closed and reopened, or after a browser restart. */
  S.KEY_FILE = 'fyg_file';
  S.LOG_CAP = 500;

  /** Run level status. */
  S.STATUS = {
    IDLE: 'idle',
    RUNNING: 'running',
    PAUSED: 'paused',
    ATTENTION: 'attention',
    DONE: 'done'
  };

  /** Per row status. */
  S.ROW = {
    PENDING: 'pending',
    ACTIVE: 'active',
    DONE: 'done',
    SKIPPED: 'skipped',
    FAILED: 'failed'
  };

  /*
   * Why a row was skipped. Two quite different things share one status, and the
   * reason is what lets the panel tell them apart without adding a sixth status
   * that recount, the stat tiles, the filter chips and the CSS would all have to
   * learn. Rows stored by an earlier build have no reason at all, which every
   * reader treats as "not stated".
   */
  S.SKIP = {
    HAD_LINK: 'hadLink',   // the sheet already carried a link for it
    EXISTS: 'exists'       // Fygaro already has a product with that code
  };

  S.STEP = {
    NAV_TO_PRODUCTS: 'NAV_TO_PRODUCTS',
    CLICK_CREATE: 'CLICK_CREATE',
    FILL_PRODUCT: 'FILL_PRODUCT',
    OPEN_PRODUCT: 'OPEN_PRODUCT',
    CLICK_CREATE_LINK: 'CLICK_CREATE_LINK',
    FILL_LINK: 'FILL_LINK',
    CAPTURE_LINK: 'CAPTURE_LINK'
  };

  S.STEP_ORDER = [
    S.STEP.NAV_TO_PRODUCTS,
    S.STEP.CLICK_CREATE,
    S.STEP.FILL_PRODUCT,
    S.STEP.OPEN_PRODUCT,
    S.STEP.CLICK_CREATE_LINK,
    S.STEP.FILL_LINK,
    S.STEP.CAPTURE_LINK
  ];

  S.STEP_LABEL = {
    NAV_TO_PRODUCTS: 'Open Products',
    CLICK_CREATE: 'Start a new item',
    FILL_PRODUCT: 'Fill the product form',
    OPEN_PRODUCT: 'Open the new product',
    CLICK_CREATE_LINK: 'Start the Fygaro Link',
    FILL_LINK: 'Fill the link form',
    CAPTURE_LINK: 'Capture the link'
  };

  /** Messages exchanged between the side panel, the worker and the page. */
  S.MSG = {
    GET_STATE: 'GET_STATE',
    STATE_CHANGED: 'STATE_CHANGED',
    LOAD_CATALOG: 'LOAD_CATALOG',
    UPDATE_SETTINGS: 'UPDATE_SETTINGS',
    START: 'START',
    PAUSE: 'PAUSE',
    RESUME: 'RESUME',
    STOP: 'STOP',
    RESET: 'RESET',
    RETRY: 'RETRY',
    SKIP_ROW: 'SKIP_ROW',
    REQUEST_JOB: 'REQUEST_JOB',
    STEP_DONE: 'STEP_DONE',
    STEP_FAILED: 'STEP_FAILED',
    CONTENT_LOG: 'CONTENT_LOG',
    OPEN_FYGARO: 'OPEN_FYGARO'
  };

  S.APP_URL = 'https://www.fygaro.com/en/app/dashboard/';
  S.ORIGIN = 'https://www.fygaro.com';

  /** Headers the catalog is expected to use, with tolerated alternatives. */
  S.HEADERS = {
    code: ['Código', 'Codigo', 'Code'],
    name: ['Servicios', 'Servicio', 'Service', 'Services'],
    price: ['Precio Total', 'Precio', 'Price'],
    link: ['Link', 'Enlace', 'URL'],
    note: ['Nota', 'Note', 'Estado del enlace']
  };

  /*
   * Headers written above columns this extension has to create for itself,
   * because the catalog arrived with nowhere to put a link.
   *
   * LINK_HEADER must stay one of S.HEADERS.link. That single fact is what makes
   * a restart safe: the header written on the first run is found by findColumn
   * on the second, so rows that already carry a link are skipped rather than
   * created a second time.
   */
  S.LINK_HEADER = 'Link';

  /*
   * Why a row finished without a link, written beside it so the reason survives
   * in the spreadsheet rather than only in the panel. It is deliberately NOT the
   * link column: a note sitting there would read as a link on the next load and
   * the row would be skipped for the wrong reason.
   */
  S.NOTE_HEADER = 'Nota';

  /*
   * Bumped when a stored settings value needs correcting rather than merely
   * defaulting. Builds before version 2 could persist a page zoom of 100 that
   * the user never chose, because an empty field fell back to a literal instead
   * of to the documented default.
   */
  S.SETTINGS_VERSION = 2;

  S.DEFAULT_SETTINGS = {
    minDelayMs: 1200,
    maxDelayMs: 3000,
    stepTimeoutMs: 20000,
    maxAttempts: 2,
    /* The side panel takes width away from the page, which can push Fygaro's
     * responsive layout over a breakpoint and swap which copy of a duplicated
     * control is on screen. Zooming out gives the page its room back and keeps
     * the desktop layout, so the DOM stays the one the steps were written for. */
    zoomPercent: 67,
    dryRun: false
  };

  /** URL shapes the automation recognises. */
  S.ROUTES = {
    dashboard: /^\/(?:en|es)\/app\/dashboard\/?$/,
    productList: /^\/(?:en|es)\/app\/products\/?$/,
    productAdd: /^\/(?:en|es)\/app\/products\/add\/?$/,
    productDetail: /^\/(?:en|es)\/app\/products\/[0-9a-f-]{36}\/permalink\/product\/?$/i,
    linkAdd: /^\/(?:en|es)\/app\/payment-buttons\/payments\/payment-buttons\/add\/?$/,
    linkDone: /^\/(?:en|es)\/app\/payment-buttons\/payments\/[0-9a-f-]{36}\/payment-buttons\/permalink\/?$/i
  };

  /** Which route each step is allowed to act on. */
  S.STEP_ROUTE = {
    NAV_TO_PRODUCTS: 'dashboard',
    CLICK_CREATE: 'productList',
    FILL_PRODUCT: 'productAdd',
    OPEN_PRODUCT: 'productList',
    CLICK_CREATE_LINK: 'productDetail',
    FILL_LINK: 'linkAdd',
    CAPTURE_LINK: 'linkDone'
  };

  /** Names the current path, or an empty string when it is not one we drive. */
  S.routeOf = function (pathname) {
    var keys = Object.keys(S.ROUTES);
    for (var i = 0; i < keys.length; i++) {
      if (S.ROUTES[keys[i]].test(pathname)) return keys[i];
    }
    return '';
  };

  S.defaultRun = function () {
    return {
      version: 1,
      status: S.STATUS.IDLE,
      settings: Object.assign({}, S.DEFAULT_SETTINGS),
      file: null,
      cursor: -1,
      step: S.STEP.NAV_TO_PRODUCTS,
      attempt: 0,
      pending: null,
      tabId: null,
      startedAt: null,
      finishedAt: null,
      rowStartedAt: null,
      waitingSince: null,
      /* The page zoom before the run touched it, so it can be handed back. */
      originalZoom: null,
      durations: [],
      stats: { total: 0, toProcess: 0, done: 0, failed: 0, skipped: 0 }
    };
  };

  /**
   * Reads the whole run, filling in defaults for anything not stored yet.
   *
   * Settings are merged one level deeper than the rest. A plain Object.assign
   * would let a settings object saved by an older version replace the defaults
   * wholesale, so any setting added later would silently arrive as undefined for
   * anyone who had already used the extension.
   */
  S.read = function () {
    return chrome.storage.local.get([S.KEY_RUN, S.KEY_ROWS, S.KEY_LOG]).then(function (data) {
      var stored = data[S.KEY_RUN] || {};
      var run = Object.assign(S.defaultRun(), stored);
      var storedSettings = stored.settings || {};
      run.settings = S.migrateSettings(
        Object.assign({}, S.DEFAULT_SETTINGS, storedSettings),
        storedSettings.settingsVersion
      );
      run.stats = Object.assign({ total: 0, toProcess: 0, done: 0, failed: 0, skipped: 0 }, stored.stats || {});
      return {
        run: run,
        rows: data[S.KEY_ROWS] || [],
        log: data[S.KEY_LOG] || []
      };
    });
  };

  /**
   * Repairs settings that an older build could have saved wrongly. Merging in
   * defaults is not enough for those, because a wrong value is present rather
   * than missing, so it would win the merge.
   *
   * The version is read from what was actually stored, never from the merged
   * object. `settingsVersion` is deliberately absent from DEFAULT_SETTINGS: if
   * it were there, every merge would look already migrated and nothing would
   * ever be repaired.
   */
  S.migrateSettings = function (settings, storedVersion) {
    if (storedVersion === S.SETTINGS_VERSION) return settings;
    settings.zoomPercent = S.DEFAULT_SETTINGS.zoomPercent;
    settings.settingsVersion = S.SETTINGS_VERSION;
    return settings;
  };

  S.writeRun = function (run) {
    var payload = {};
    payload[S.KEY_RUN] = run;
    return chrome.storage.local.set(payload);
  };

  S.writeRows = function (rows) {
    var payload = {};
    payload[S.KEY_ROWS] = rows;
    return chrome.storage.local.set(payload);
  };

  S.writeLog = function (log) {
    var payload = {};
    payload[S.KEY_LOG] = log;
    return chrome.storage.local.set(payload);
  };

  S.clear = function () {
    return chrome.storage.local.remove([S.KEY_RUN, S.KEY_ROWS, S.KEY_LOG, S.KEY_FILE]);
  };

  /*
   * Wipes the extension's storage outright rather than removing the four keys
   * it currently uses, so a key left behind by an older version cannot survive
   * a clear and quietly come back.
   */
  S.clearAll = function () {
    return chrome.storage.local.clear();
  };

  /** Recomputes the counters shown in the side panel. */
  S.recount = function (rows) {
    var stats = { total: rows.length, toProcess: 0, done: 0, failed: 0, skipped: 0 };
    for (var i = 0; i < rows.length; i++) {
      var s = rows[i].status;
      if (s === S.ROW.DONE) stats.done++;
      else if (s === S.ROW.FAILED) stats.failed++;
      else if (s === S.ROW.SKIPPED) stats.skipped++;
      // Failed rows are not retried automatically, so they are not still to do.
      if (s === S.ROW.PENDING || s === S.ROW.ACTIVE) stats.toProcess++;
    }
    return stats;
  };
})(typeof globalThis !== 'undefined' ? globalThis : self);
